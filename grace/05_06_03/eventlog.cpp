/*******************************************************************************
 EVENTLOG.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:  Outputs messages to log file
*******************************************************************************/

//Include files
#include "eventlog.h"

//sets filename
void CEventLog::SetFilename(string filename)
{
    log_file=filename;
}

//writes message and/or time/date to log file
void CEventLog::LogData(string text, int status=0)
{
  if (log_file.size() <1) SetFilename("log.txt");
  ofstream file(log_file.c_str(), std::ios::app);
  if(file){
    switch(status)
    {
      case LOG_TEXT:
        file << text.c_str() << endl;
        break;
      case LOG_TEXT_TIME:
        file << GetTimeString() << ": " << text.c_str() << endl;
        break;  
      case LOG_TEXT_DATE:
        file << GetDateString() << ": " << text.c_str() << endl;
        break;  
      case LOG_TEXT_DATE_TIME:
        file << GetDateString() << ", " << GetTimeString() << ": "
			 << text.c_str() << endl;
        break;  
    }
  }
}

//returns a string with time
string CEventLog::GetTimeString(void)
{
  struct tm *ptr;
  string sTime;
  time_t lt;
  lt = time(NULL);
  ptr = localtime(&lt);
  sTime= asctime(ptr);
  return sTime.substr(11,8);
}

//returns a string with date
string CEventLog::GetDateString(void)
{
  struct tm *ptr;
  string sTime;
  time_t lt;
  lt = time(NULL);
  ptr = localtime(&lt);
  sTime= asctime(ptr);
  return sTime.substr(4,6) + ", " + sTime.substr(20,4);
}


