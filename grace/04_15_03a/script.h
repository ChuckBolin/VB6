/*******************************************************************************
 SCRIPT class
 Programmer: Chuck Bolin, 2003
 Purpose:  Allows for data logging of events (strings)
*******************************************************************************/
#ifndef _ENGINE_SCRIPT_H
#define _ENGINE_SCRIPT_H
#define GLUT_DISABLE_ATEXIT_HACK

//include files
#include <string>
#include <iostream>
#include <fstream>
#include <vector>
#include <strstream>
#include "mymath.h"

typedef struct{
  string name;
  int args;
  string argv[10];
} PARAM;


// Class Definition
class CScript
{
  public:

 
  enum SCRIPT_CONSTANTS
  {
    SYNTAX_OK =0,
    SYNTAX_PARENTHESIS ,   //incorrect num of open parenthesis
    SYNTAX_SEMICOLON  ,     //missing semicolon
    SYNTAX_OPEN_PARENTHESIS , 
    SYNTAX_CLOSED_PARENTHESIS ,
    SYNTAX_QUOTE ,
    SYNTAX_UNKNOWN =99      //unknown error
  };
  
  static CScript& Instance()
  {
    static CScript instance;
    return instance;  
  }
  
  //member functions
  int CheckSyntax(string);
  string CleanUp(string);
  void DisplayResults(string);
  int GetNumErrors(void);
  void Push(string);
  int GetSize(void);
  int GetLength(int);
  string Get(int);
  PARAM ExtractInfo(int);
  void ParseFile(void);
  PARAM GetScriptLine(int);
  
  private:
  //constructor
  CScript(){}
  
  //destructor
  ~CScript(){}   
  
  int number_errors;
  int left_brace;
  int right_brace;
  vector<string> script;     //stores script info
  vector <PARAM> parameters; //stores all functions in order

};
#define gScript CScript::Instance()
#endif _ENGINE_SCRIPT_H


/***********************************************************************
  Sample usage:
  
  //read command line arg
  string scriptfile;
  scriptfile = argv[1];
  gScript.DisplayResults(scriptfile);
  . . .
  string lineout= gScript.CleanUp(line);  //clean up line..spaces, lower case
  status = gScript.CheckSyntax(lineout);  //returns integer indicating error

***********************************************************************/
