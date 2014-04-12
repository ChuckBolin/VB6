/*******************************************************************************
 EVENTLOG class
 Programmer: Chuck Bolin, 2003
 Purpose:  Allows for data logging of events (strings)
*******************************************************************************/
#ifndef _ENGINE_EVENTLOG_H
#define _ENGINE_EVENTLOG_H

//include files
#include <string>
#include <fstream>
#include <ctime>
#include <iostream>
#include <strstream>

// Class Definition
class CEventLog
{
  public:

  enum EVENTLOG_CONSTANTS
  {
    LOG_TEXT ,
    LOG_TEXT_TIME,
    LOG_TEXT_DATE,
    LOG_TEXT_DATE_TIME
  };
  
  static CEventLog& Instance()
  {
    static CEventLog instance;
    return instance;  
  }
  
  //member functions
  void LogData(string, int);
  string GetTimeString(void);
  string GetDateString(void);
  void SetFilename(string);

  private:
  //constructor
  CEventLog(){}
  
  //destructor
  ~CEventLog(){}   
  
  string log_file;  //stores data logging filename
  
};
#define gEventLog CEventLog::Instance()
#endif _ENGINE_EVENTLOG_H


/***********************************************************************
  Sample usage:
  gEventLog.SetFilename("log.txt");
  gEventLog.LogData("Program starting...",gEventLog.LOG_TEXT);
 
***********************************************************************/
