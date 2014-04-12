//firstrc.h
#include "graphics.h"
#include <iostream>
#include <cstdlib>
#include <iomanip>
#include <fstream>
  
#ifndef _FIRST_RC_
#define _FIRST_RC_

struct AutoCommand{
  float startTime;
  int command;
  float stopTime;
  int speed;
};

//function prototypes
void DrawTitle(void);
void DrawLabels(void);
void DrawValues(void);
void DrawInstructions(void);
void PrintTimeDate(void);           
void PrintToFile(void);
void increaseDataPos(void);
void decreaseDataPos(void);
void moveCursor(void);
void shiftDataPosRight(void);
void shiftDataPosLeft(void);
void clearHelp(void);
void displayHelp(void);

#endif
