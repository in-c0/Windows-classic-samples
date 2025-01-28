#include <windows.h>
#include <iostream>

class Timer
{
public:
    Timer()
    {
        // Get the timer frequency once during initialization
        QueryPerformanceFrequency(&m_frequency);
    }

    // Get the current time in seconds
    double GetTimeInSeconds()
    {
        LARGE_INTEGER counter;
        QueryPerformanceCounter(&counter);
        return static_cast<double>(counter.QuadPart) / m_frequency.QuadPart;
    }

private:
    LARGE_INTEGER m_frequency; // Holds the timer frequency (ticks per second)
};
