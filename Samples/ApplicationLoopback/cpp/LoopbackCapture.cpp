#include <shlobj.h>
#include <wchar.h>
#include <iostream>
#include <audioclientactivationparams.h>
#include <mmdeviceapi.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <Audiopolicy.h>
#include "LoopbackCapture.h"

#define BITS_PER_BYTE 8

HRESULT CLoopbackCapture::SetDeviceStateErrorIfFailed(HRESULT hr)
{
    if (FAILED(hr))
    {
        m_DeviceState = DeviceState::Error;
    }
    return hr;
}



HRESULT CLoopbackCapture::WriteCapturedDataToPlayback(BYTE* pData, UINT32 numFrames)
{
    // We might need to handle the case where the render buffer has fewer frames than we want to write.
    // For simplicity, let's assume we have enough buffer. In a robust app, you'd loop if needed.
    UINT32 padding = 0;
    RETURN_IF_FAILED(m_RenderClient->GetCurrentPadding(&padding));

    UINT32 framesAvailable = m_RenderBufferFrames - padding;
    if (numFrames > framesAvailable)
    {
        // In a real app, you'd handle this properly (maybe partial write).
        numFrames = framesAvailable;
    }

    BYTE* pRenderBuffer = nullptr;
    RETURN_IF_FAILED(m_AudioRenderClient->GetBuffer(numFrames, &pRenderBuffer));

    // Copy from capture buffer to playback buffer
    size_t bytesToCopy = (size_t)numFrames * m_CaptureFormat.nBlockAlign;
    memcpy_s(pRenderBuffer, bytesToCopy, pData, bytesToCopy);

    // Deliver the data
    RETURN_IF_FAILED(m_AudioRenderClient->ReleaseBuffer(numFrames, 0));

    return S_OK;
}



HRESULT CLoopbackCapture::StartPlayback()
{
    RETURN_IF_FAILED(m_RenderClient->Start());
    return S_OK;
}

HRESULT CLoopbackCapture::InitializePlayback()
{
    // 1. Get the default render device via classic endpoint APIs
    wil::com_ptr_nothrow<IMMDeviceEnumerator> spEnumerator;
    RETURN_IF_FAILED(CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        nullptr,
        CLSCTX_ALL,
        IID_PPV_ARGS(&spEnumerator)));

    wil::com_ptr_nothrow<IMMDevice> spDevice;
    RETURN_IF_FAILED(spEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &spDevice));

    // 2. Activate IAudioClient for rendering
    RETURN_IF_FAILED(spDevice->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&m_RenderClient));

    // 3. Initialize the render client in shared mode with the same format we used for capturing
    //    The buffer duration here is just an example (200ms in 100-ns units).
    REFERENCE_TIME hnsBufferDuration = 2000000; // 200ms
    RETURN_IF_FAILED(m_RenderClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0, // no special stream flags, can use EVENTCALLBACK if you prefer
        hnsBufferDuration,
        0,
        &m_CaptureFormat,
        nullptr));

    // 4. Get the size of the render buffer and the IAudioRenderClient interface
    RETURN_IF_FAILED(m_RenderClient->GetBufferSize(&m_RenderBufferFrames));
    RETURN_IF_FAILED(m_RenderClient->GetService(IID_PPV_ARGS(&m_AudioRenderClient)));

    return S_OK;
}

HRESULT CLoopbackCapture::InitializeLoopbackCapture()
{
    // Create events for sample ready or user stop
    RETURN_IF_FAILED(m_SampleReadyEvent.create(wil::EventOptions::None));

    // Initialize MF
    RETURN_IF_FAILED(MFStartup(MF_VERSION, MFSTARTUP_LITE));

    // Register MMCSS work queue
    DWORD dwTaskID = 0;
    RETURN_IF_FAILED(MFLockSharedWorkQueue(L"Capture", 0, &dwTaskID, &m_dwQueueID));

    // Set the capture event work queue to use the MMCSS queue
    m_xSampleReady.SetQueueID(m_dwQueueID);

    // Create the completion event as auto-reset
    RETURN_IF_FAILED(m_hActivateCompleted.create(wil::EventOptions::None));

    // Create the capture-stopped event as auto-reset
    RETURN_IF_FAILED(m_hCaptureStopped.create(wil::EventOptions::None));

    return S_OK;
}

CLoopbackCapture::~CLoopbackCapture()
{
    if (m_dwQueueID != 0)
    {
        MFUnlockWorkQueue(m_dwQueueID);
    }
}

HRESULT CLoopbackCapture::ActivateAudioInterface(DWORD processId, bool includeProcessTree)
{
    return SetDeviceStateErrorIfFailed([&]() -> HRESULT
        {
            AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
            audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
            audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ?
                PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
            audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

            PROPVARIANT activateParams = {};
            activateParams.vt = VT_BLOB;
            activateParams.blob.cbSize = sizeof(audioclientActivationParams);
            activateParams.blob.pBlobData = (BYTE*)&audioclientActivationParams;

            wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
            RETURN_IF_FAILED(ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activateParams, this, &asyncOp));

            // Wait for activation completion
            m_hActivateCompleted.wait();

            return m_activateResult;
        }());
}

//
//  ActivateCompleted()
//
//  Callback implementation of ActivateAudioInterfaceAsync function.  This will be called on MTA thread
//  when results of the activation are available.
//
HRESULT CLoopbackCapture::ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation)
{
    m_activateResult = SetDeviceStateErrorIfFailed([&]()->HRESULT
        {
            // Check for a successful activation result
            HRESULT hrActivateResult = E_UNEXPECTED;
            wil::com_ptr_nothrow<IUnknown> punkAudioInterface;
            RETURN_IF_FAILED(operation->GetActivateResult(&hrActivateResult, &punkAudioInterface));
            RETURN_IF_FAILED(hrActivateResult);

            // Get the pointer for the Audio Client
            RETURN_IF_FAILED(punkAudioInterface.copy_to(&m_AudioClient));

            // The app can also call m_AudioClient->GetMixFormat instead to get the capture format.
            // 16 - bit PCM format.
            m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
            m_CaptureFormat.nChannels = 2;
            m_CaptureFormat.nSamplesPerSec = 44100;
            m_CaptureFormat.wBitsPerSample = 16;
            m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / BITS_PER_BYTE;
            m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

            // Initialize the AudioClient in Shared Mode with the user specified buffer
            RETURN_IF_FAILED(m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                200000,
                AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                &m_CaptureFormat,
                nullptr));

            // Get the maximum size of the AudioClient Buffer
            RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

            // Get the capture client
            RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

            // Create Async callback for sample events
            RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

            // Tell the system which event handle it should signal when an audio buffer is ready to be processed by the client
            RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

            // Creates the WAV file.
            RETURN_IF_FAILED(CreateWAVFile());

            // Everything is ready.
            m_DeviceState = DeviceState::Initialized;

            return S_OK;
        }());

    // Let ActivateAudioInterface know that m_activateResult has the result of the activation attempt.
    m_hActivateCompleted.SetEvent();
    return S_OK;
}

//
//  CreateWAVFile()
//
//  Creates a WAV file in music folder
//
HRESULT CLoopbackCapture::CreateWAVFile()
{
    return SetDeviceStateErrorIfFailed([&]()->HRESULT
        {
            m_hFile.reset(CreateFile(m_outputFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
            RETURN_LAST_ERROR_IF(!m_hFile);

            // Create and write the WAV header

                // 1. RIFF chunk descriptor
            DWORD header[] = {
                                FCC('RIFF'),        // RIFF header
                                0,                  // Total size of WAV (will be filled in later)
                                FCC('WAVE'),        // WAVE FourCC
                                FCC('fmt '),        // Start of 'fmt ' chunk
                                sizeof(m_CaptureFormat) // Size of fmt chunk
            };
            DWORD dwBytesWritten = 0;
            RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), header, sizeof(header), &dwBytesWritten, NULL));

            m_cbHeaderSize += dwBytesWritten;

            // 2. The fmt sub-chunk
            WI_ASSERT(m_CaptureFormat.cbSize == 0);
            RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_CaptureFormat, sizeof(m_CaptureFormat), &dwBytesWritten, NULL));
            m_cbHeaderSize += dwBytesWritten;

            // 3. The data sub-chunk
            DWORD data[] = { FCC('data'), 0 };  // Start of 'data' chunk
            RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), data, sizeof(data), &dwBytesWritten, NULL));
            m_cbHeaderSize += dwBytesWritten;

            return S_OK;
        }());
}


//
//  FixWAVHeader()
//
//  The size values were not known when we originally wrote the header, so now go through and fix the values
//
HRESULT CLoopbackCapture::FixWAVHeader()
{
    // Write the size of the 'data' chunk first
    DWORD dwPtr = SetFilePointer(m_hFile.get(), m_cbHeaderSize - sizeof(DWORD), NULL, FILE_BEGIN);
    RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == dwPtr);

    DWORD dwBytesWritten = 0;
    RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_cbDataSize, sizeof(DWORD), &dwBytesWritten, NULL));

    // Write the total file size, minus RIFF chunk and size
    // sizeof(DWORD) == sizeof(FOURCC)
    RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == SetFilePointer(m_hFile.get(), sizeof(DWORD), NULL, FILE_BEGIN));

    DWORD cbTotalSize = m_cbDataSize + m_cbHeaderSize - 8;
    RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &cbTotalSize, sizeof(DWORD), &dwBytesWritten, NULL));

    RETURN_IF_WIN32_BOOL_FALSE(FlushFileBuffers(m_hFile.get()));

    return S_OK;
}

HRESULT CLoopbackCapture::StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName)
{
    m_outputFileName = outputFileName;
    auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

    RETURN_IF_FAILED(InitializeLoopbackCapture());
    RETURN_IF_FAILED(ActivateAudioInterface(processId, includeProcessTree));

    // We should be in the initialzied state if this is the first time through getting ready to capture.
    if (m_DeviceState == DeviceState::Initialized)
    {
        m_DeviceState = DeviceState::Starting;
        return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
    }

    return S_OK;
}

//
//  OnStartCapture()
//
//  Callback method to start capture
//
HRESULT CLoopbackCapture::OnStartCapture(IMFAsyncResult* pResult)
{
    return SetDeviceStateErrorIfFailed([&]()->HRESULT
        {
            // Start the capture
            RETURN_IF_FAILED(m_AudioClient->Start());

            // === NEW: Initialize and start playback ===
            RETURN_IF_FAILED(InitializePlayback());
            RETURN_IF_FAILED(StartPlayback());
            // ==========================================

            m_DeviceState = DeviceState::Capturing;
            MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);

            return S_OK;
        }());
}


//
//  StopCaptureAsync()
//
//  Stop capture asynchronously via MF Work Item
//
HRESULT CLoopbackCapture::StopCaptureAsync()
{
    RETURN_HR_IF(E_NOT_VALID_STATE, (m_DeviceState != DeviceState::Capturing) &&
        (m_DeviceState != DeviceState::Error));

    m_DeviceState = DeviceState::Stopping;

    RETURN_IF_FAILED(MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStopCapture, nullptr));

    // Wait for capture to stop
    m_hCaptureStopped.wait();

    return S_OK;
}

//
//  OnStopCapture()
//
//  Callback method to stop capture
//
HRESULT CLoopbackCapture::OnStopCapture(IMFAsyncResult* pResult)
{
    // Stop capture by cancelling Work Item
    // Cancel the queued work item (if any)
    if (0 != m_SampleReadyKey)
    {
        MFCancelWorkItem(m_SampleReadyKey);
        m_SampleReadyKey = 0;
    }

    m_AudioClient->Stop();
    m_SampleReadyAsyncResult.reset();

    return FinishCaptureAsync();
}

//
//  FinishCaptureAsync()
//
//  Finalizes WAV file on a separate thread via MF Work Item
//
HRESULT CLoopbackCapture::FinishCaptureAsync()
{
    // We should be flushing when this is called
    return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xFinishCapture, nullptr);
}

//
//  OnFinishCapture()
//
//  Because of the asynchronous nature of the MF Work Queues and the DataWriter, there could still be
//  a sample processing.  So this will get called to finalize the WAV header.
//
HRESULT CLoopbackCapture::OnFinishCapture(IMFAsyncResult* pResult)
{
    // FixWAVHeader will set the DeviceStateStopped when all async tasks are complete
    HRESULT hr = FixWAVHeader();

    m_DeviceState = DeviceState::Stopped;

    m_hCaptureStopped.SetEvent();

    return hr;
}

//
//  OnSampleReady()
//
//  Callback method when ready to fill sample buffer
//
HRESULT CLoopbackCapture::OnSampleReady(IMFAsyncResult* pResult)
{
    if (SUCCEEDED(OnAudioSampleRequested()))
    {
        // Re-queue work item for next sample
        if (m_DeviceState == DeviceState::Capturing)
        {
            // Re-queue work item for next sample
            return MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
        }
    }
    else
    {
        m_DeviceState = DeviceState::Error;
    }

    return S_OK;
}

//
//  OnAudioSampleRequested()
//
//  Called when audio device fires m_SampleReadyEvent
//
HRESULT CLoopbackCapture::OnAudioSampleRequested()
{
    UINT32 FramesAvailable = 0;
    BYTE* Data = nullptr;
    DWORD dwCaptureFlags;
    UINT64 u64DevicePosition = 0;
    UINT64 u64QPCPosition = 0;
    DWORD cbBytesToCapture = 0;

    auto lock = m_CritSec.lock();

    if (m_DeviceState == DeviceState::Stopping)
    {
        return S_OK;
    }

    // keep retrieving as many frames as available
    while (SUCCEEDED(m_AudioCaptureClient->GetNextPacketSize(&FramesAvailable)) && FramesAvailable > 0)
    {
        cbBytesToCapture = FramesAvailable * m_CaptureFormat.nBlockAlign;

        if ((m_cbDataSize + cbBytesToCapture) < m_cbDataSize)
        {
            StopCaptureAsync();
            break;
        }

        RETURN_IF_FAILED(
            m_AudioCaptureClient->GetBuffer(&Data, &FramesAvailable, &dwCaptureFlags,
                &u64DevicePosition, &u64QPCPosition));

        // ========== NEW: Pass data immediately to playback ==============
        RETURN_IF_FAILED(WriteCapturedDataToPlayback(Data, FramesAvailable));
        // ================================================================

        // The code below is the original WAV-file writing. You can remove or keep for reference:
        /*
        if (m_DeviceState != DeviceState::Stopping)
        {
            DWORD dwBytesWritten = 0;
            WriteFile(
                m_hFile.get(),
                Data,
                cbBytesToCapture,
                &dwBytesWritten,
                NULL);
        }
        m_cbDataSize += cbBytesToCapture;
        */

        // release buffer
        m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
    }

    return S_OK;
}
