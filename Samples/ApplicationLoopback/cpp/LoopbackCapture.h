#pragma once

#include <queue>
#include <vector>

#include <AudioClient.h>
#include <mmdeviceapi.h>
#include <initguid.h>
#include <guiddef.h>
#include <mfapi.h>

#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>

#include "Common.h"

using namespace Microsoft::WRL;

class CLoopbackCapture :
    public RuntimeClass< RuntimeClassFlags< ClassicCom >, FtmBase, IActivateAudioInterfaceCompletionHandler >
{
public:
    CLoopbackCapture() = default;
    ~CLoopbackCapture();

    HRESULT StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName);
    HRESULT StopCaptureAsync();

    // --- existing methods from your code ---
    METHODASYNCCALLBACK(CLoopbackCapture, StartCapture, OnStartCapture);
    METHODASYNCCALLBACK(CLoopbackCapture, StopCapture, OnStopCapture);
    METHODASYNCCALLBACK(CLoopbackCapture, SampleReady, OnSampleReady);
    METHODASYNCCALLBACK(CLoopbackCapture, FinishCapture, OnFinishCapture);

    // IActivateAudioInterfaceCompletionHandler
    STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation* operation);

private:
    enum class DeviceState
    {
        Uninitialized,
        Error,
        Initialized,
        Starting,
        Capturing,
        Stopping,
        Stopped,
    };
     // We'll keep a simple queue of raw bytes
    std::queue<std::vector<BYTE>> m_AudioQueue;

    // mute control
    bool m_muteActive = false;
    DWORD m_muteEndTime = 0;    // GetTickCount time when we stop muting
    DWORD m_nextMuteStart = 0;  // next time we initiate the mute

    // We'll also have a separate event or thread that tries to "drain" the queue to playback.
    wil::unique_handle m_hPlaybackThread;
    static DWORD WINAPI PlaybackThreadProc(LPVOID lpParameter);
    DWORD DoPlaybackThread();
    bool m_bContinuePlayback = false;

    HRESULT OnStartCapture(IMFAsyncResult* pResult);
    HRESULT OnStopCapture(IMFAsyncResult* pResult);
    HRESULT OnFinishCapture(IMFAsyncResult* pResult);
    HRESULT OnSampleReady(IMFAsyncResult* pResult);
    HRESULT InitializeLoopbackCapture();
    HRESULT CreateWAVFile();      // we will ignore usage in Build #1
    HRESULT FixWAVHeader();       // same
    HRESULT OnAudioSampleRequested();
    HRESULT ActivateAudioInterface(DWORD processId, bool includeProcessTree);
    HRESULT FinishCaptureAsync();
    HRESULT SetDeviceStateErrorIfFailed(HRESULT hr);

    wil::com_ptr_nothrow<IAudioClient> m_AudioClient;
    WAVEFORMATEX m_CaptureFormat{};
    UINT32 m_BufferFrames = 0;
    wil::com_ptr_nothrow<IAudioCaptureClient> m_AudioCaptureClient;
    wil::com_ptr_nothrow<IMFAsyncResult> m_SampleReadyAsyncResult;
    wil::unique_event_nothrow m_SampleReadyEvent;
    MFWORKITEM_KEY m_SampleReadyKey = 0;
    wil::unique_hfile m_hFile;
    wil::critical_section m_CritSec;
    DWORD m_dwQueueID = 0;
    DWORD m_cbHeaderSize = 0;
    DWORD m_cbDataSize = 0;
    PCWSTR m_outputFileName = nullptr;
    HRESULT m_activateResult = E_UNEXPECTED;
    DeviceState m_DeviceState{ DeviceState::Uninitialized };
    wil::unique_event_nothrow m_hActivateCompleted;
    wil::unique_event_nothrow m_hCaptureStopped;

    // ===== NEW MEMBERS FOR PLAYBACK =====
    wil::com_ptr_nothrow<IAudioClient>       m_RenderClient;        // separate client for playback
    wil::com_ptr_nothrow<IAudioRenderClient> m_AudioRenderClient;   // for sending captured data to speakers
    UINT32                                   m_RenderBufferFrames = 0;

    HRESULT InitializePlayback();     // initializes the playback device with the same format
    HRESULT StartPlayback();          // starts the playback client
    HRESULT WriteCapturedDataToPlayback(BYTE* pData, UINT32 numFrames);
};
