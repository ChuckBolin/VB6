/*******************************************************************************
 INPUT.H
 Programmer: Chuck Bolin, 2003
 Purpose:  
*******************************************************************************/
#ifndef _ENGINE_INPUT_H
#define _ENGINE_INPUT_H
#define GLUT_DISABLE_ATEXIT_HACK

#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>
#include "gameobj.h"
#include "graphics.h"

// various inputs
const int INPUT_PASSIVE_MOUSE = 1;
const int INPUT_ACTIVE_MOUSE = 2;
const int INPUT_EXTENDED_KEY_DOWN = 3;
const int INPUT_EXTENDED_KEY_UP = 4;
const int INPUT_NORMAL_KEY_DOWN = 5;
const int INPUT_TIMER = 6;

//captures all input information
typedef struct {
  int event;
  int x,y;
  int button;
  int state;
  int key;
  unsigned char uckey;
  int value;
} PROGRAMINPUT;

//experimental --- 
typedef void MyFunc_t(float);
//MyFunc_t *pFunctionList[100];  //holds ptrs to 100 functions


// Class Definition
class CInput
{
  public:

  int mouse_select;  //returns number equal to selection


  static CInput& Instance()
  {
    static CInput instance;
    return instance;  
  }

  //public members
  void ProcessInput (PROGRAMINPUT);  
  void NormalKeyDown(PROGRAMINPUT);
  void PassiveMouse(PROGRAMINPUT input);
  void ActiveMouse(PROGRAMINPUT input);
  void ExtendedKeyDown(PROGRAMINPUT input);
  void ExtendedKeyUp(PROGRAMINPUT input);
  void Timer(PROGRAMINPUT input);
  
  private:

  //constructor
  CInput(){}
  
  //destructor
  ~CInput(){}
 
  //private variables
  
};
#define gInput CInput::Instance()
#endif _ENGINE_INPUT_H



/***********************************************************************
  Sample usage:


***********************************************************************/
