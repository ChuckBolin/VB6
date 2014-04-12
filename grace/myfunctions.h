/*******************************************************************************
 MYFUNCTIONS 
 Programmer: Chuck Bolin, 2003
 Purpose:  Collection of functions that correspond to script
*******************************************************************************/
#ifndef _ENGINE_MYFUNCTIONS_H
#define _ENGINE_MYFUNCTIONS_H
#define GLUT_DISABLE_ATEXIT_HACK

//include files
#include <iostream>
#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>
#include "mymath.h"

enum MYFUNCTION_CONSTANTS
{
  MYFUNCTION_MOVELEFT = 0,
  MYFUNCTION_MOVERIGHT = 1,
  MYFUNCTION_COLOR = 2,
  MYFUNCTION_POS3D = 3,
  MYFUNCTION_SPHERE = 4
};

void LoadFunctionTable(void);  
void FuncMoveLeft(void);
void FuncMoveRight(void);
void FuncPos3D(void);
void FuncColor(void);
void FuncSphere(void);

#endif _ENGINE_MYFUNCTIONS_H

