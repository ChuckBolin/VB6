/****************************************************************************
  MYFUNCTIONS.CPP - Stores all functions required by Engine to process
  Script file.
****************************************************************************/  
#include "myfunctions.h"

typedef void GenericFunction(void);      //type of function
GenericFunction *gpFunctionList[200]; //array of pointers to function

void LoadFunctionTable(){        //load table
  gpFunctionList[MYFUNCTION_MOVELEFT] = FuncMoveLeft;
  gpFunctionList[MYFUNCTION_MOVERIGHT] = FuncMoveRight;
  gpFunctionList[MYFUNCTION_COLOR] = FuncColor;
  gpFunctionList[MYFUNCTION_POS3D] = FuncPos3D;
  gpFunctionList[MYFUNCTION_SPHERE] = FuncSphere;  
  
}

void FuncMoveLeft(void){                    
  cout << "Moving to the left..." << endl;
}

void FuncMoveRight(void){                   
  cout << "Moving to the right..." << endl;
}

void FuncPos3D(void){
  //glTranslatef(  );
}
void FuncColor(void){
  //glColor3f(   );
}
void FuncSphere(void){
  //glutSolidSphere(  );
}

/*****************************************************************************
  Sample Usage:
  //call function from function list
  gpFunctionList[MYFUNCTION_MOVELEFT]();
  gpFunctionList[MYFUNCTION_MOVERIGHT]();
  
*****************************************************************************/
