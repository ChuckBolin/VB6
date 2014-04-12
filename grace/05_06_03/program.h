/*******************************************************************************
 PROGRAM class
 Programmer: Chuck Bolin, 2003
 Purpose:  Object contains info about the running program.
*******************************************************************************/
#ifndef _ENGINE_PROGRAM_H
#define _ENGINE_PROGRAM_H

//include files
#include <string>
#include <fstream>
#include <ctime>
#include <iostream>
#include <strstream>

struct PROGRAMINFO
{
  string version_number;
  string program_name;
  string program_description;
  string revision_date;
  string programmer_name;
  string programmer_email;
  string programmer_URL;
  string help_file;
};

// Class Definition
class CProgram
{
  public:

  static CProgram& Instance()
  {
    static CProgram instance;
    return instance;  
  }
  
  //member functions
  string SetVersion(string);
  string SetProgramName(string);
  string SetProgramDescription(string);
  string SetRevisionDate(string);
  string SetProgrammerName(string);
  string SetProgrammerEmail(string);
  string SetProgrammerURL(string);
  string SetHelpFile(string);

  private:

  //constructor
  CProgram(){}
  
  //destructor
  ~CProgram(){}   
  
  string version_number;
  string program_name;
  string program_description;
  string revision_date;
  string programmer_name;
  string programmer_email;
  string programmer_URL;
  string help_file;


};
#define gProgram CProgram::Instance()
#endif _ENGINE_PROGRAM_H


/***********************************************************************
  Sample usage:

  PROGRAMINFO prog;
  
  //program specific information
  prog.version_number = gProgram.SetVersion("0.1");
  prog.program_name = gProgram.SetProgramName("Generic Program");
  prog.program_description = gProgram.SetProgramDescription("This is a basic program.");
  prog.revision_date = gProgram.SetRevisionDate("01.16.03");
  prog.programmer_name = gProgram.SetProgrammerName("Chuck Bolin");
  prog.programmer_email = gProgram.SetProgrammerEmail("cbolin@dycon.com");
  prog.programmer_URL = gProgram.SetProgrammerURL("http://www.clg-net.com");
  prog.help_file = gProgram.SetHelpFile("help.htm");
  
  //datalog info
  gEventLog.LogData("Version: " + prog.version_number, 0);
  gEventLog.LogData("Program: " + prog.program_name,0);
  gEventLog.LogData("Description: " + prog.program_description,0);
  gEventLog.LogData("Revision Date: " + prog.revision_date,0);
  gEventLog.LogData("Programmer: " + prog.programmer_name,0);
  gEventLog.LogData("Email: " + prog.programmer_email,0);
  gEventLog.LogData("URL: " + prog.programmer_URL,0);
  gEventLog.LogData("Help file: " + prog.help_file,0);
  
***********************************************************************/
