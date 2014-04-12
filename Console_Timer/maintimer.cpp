/******************************************************************
  timer1.dev, maintimer.cpp - written by Chuck Bolin, July 2004
  Credit:  OpenGL Game Programming, Chapter 19.

*******************************************************************/
#include <iostream>
#include <cstdlib>
#include <windows.h>

using namespace std;

float GetElapsedSeconds ( LARGE_INTEGER, LARGE_INTEGER);

int main(int argc, char *argv[])
{

  //64-bit integers
  LARGE_INTEGER m_ticksPerSecond;
  LARGE_INTEGER m_startTime;
  
  //verifyies that high performance timer will work on 
  //this PC...retrieves frequency of timer
  if (!QueryPerformanceFrequency(&m_ticksPerSecond)){
    cout << "High performance timer not supported..." << endl;
    system("pause");
    return 0;
  }
  else
    cout << "Supported..." << endl;
  
  //load current     
  QueryPerformanceCounter(&m_startTime);

  //calls function 10x
  for(int i=0;i<10;i++){
    cout << "Elapsed time: " 
         << GetElapsedSeconds ( m_startTime, m_ticksPerSecond) << endl;
  }  
        
  system("PAUSE");	
  return 0;
}

// returns the number of seconds that have elapsed since the time 
// passed to the function (1st arg) and the current time
float GetElapsedSeconds ( LARGE_INTEGER previousTime, LARGE_INTEGER ticksPerSecond)
{
  LARGE_INTEGER currentTime;

  //reads timer counter 
  QueryPerformanceCounter(&currentTime);

  //calcs time difference
  float seconds = ((float)currentTime.QuadPart - (float)previousTime.QuadPart)/
                   (float)ticksPerSecond.QuadPart;
  return seconds;
}

