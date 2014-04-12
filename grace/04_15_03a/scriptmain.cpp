/********************************************************************
  Program: Sciptreader  Author: Chuck Bolin (cbolin@dycon.com)
  Description: Reads a ASCII file (*.gam) and extracts info required
  for engine operation.                                       
********************************************************************/
#include <cstdlib>
#include <iostream>
#include <fstream>
#include <string>
#include "script.h"
#include "myfunctions.h"

#define GLUT_DISABLE_ATEXIT_HACK

using namespace std;

typedef void GenericFunction(void);      //type of function
extern GenericFunction *gpFunctionList[200]; //array of pointers to function


int main(int argc, char *argv[])
{
  LoadFunctionTable();               //assign pointers to functions
  string scriptfile;

  //opens file specified by command line argument
  if (argc >1 ){
    scriptfile = argv[1];
    gScript.DisplayResults(scriptfile);
  
    if (gScript.GetNumErrors() > 0){
      cout << "Errors reading script file. ABORTING program..." << endl;
      system("pause");
      exit(0);
    }
  }
  
  
  system("cls");
  PARAM par;
  gScript.ParseFile();
  for(int i=0;i< gScript.GetSize();i++){
    par = gScript.GetScriptLine(i);
    cout << "par.name: " << par.name << " " << "par.args: " << par.args << " " << endl; 
    for (int j=0;j< par.args;j++){
      cout << "  " <<  par.argv[j];
    }
    cout << endl ;
  }
  
  
  system("pause");                   //necessary for Dev-C++
  return 0;
}


  /*
#include <fstream>
#include <string>
#include "script.h"


  string scriptfile;
  scriptfile = argv[1];
  gScript.DisplayResults(scriptfile);
  */


    //cout << "Begin" << endl;
  //gpFunctionList[MYFUNCTION_MOVELEFT]();
  //gpFunctionList[MYFUNCTION_MOVERIGHT]();
  //cout << "End" << endl;

