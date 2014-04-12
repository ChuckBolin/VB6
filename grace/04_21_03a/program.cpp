/*******************************************************************************
 PROGRAM class
 Programmer: Chuck Bolin, 2003
 Purpose:  Object contains info about the running program.
*******************************************************************************/

//Include files
#include "program.h"

string CProgram::SetVersion(string version)
{
  version_number = version;
  return version_number;
}

string CProgram::SetProgramName(string pname)
{
  program_name = pname;
  return program_name;
}
string CProgram::SetProgramDescription(string pdescript)
{
  program_description = pdescript;
  return program_description;
}

string CProgram::SetRevisionDate(string rdate)
{
  revision_date = rdate;
  return revision_date;
}

string CProgram::SetProgrammerName(string pname)
{
  programmer_name = pname;
  return programmer_name;
}
string CProgram::SetProgrammerEmail(string email)
{
  programmer_email = email;
  return programmer_email;
}

string CProgram::SetProgrammerURL(string URL)
{
  programmer_URL = URL;
  return programmer_URL;
}
string CProgram::SetHelpFile(string file)
{
  help_file = file;
  return help_file;
}

