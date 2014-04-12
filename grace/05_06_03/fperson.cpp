/*******************************************************************************
 FPERSON.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:    Controls first person
*******************************************************************************/

//Include files
#include "fperson.h"

//Initializes player (or camera for first person strategy)
void CFPerson::Initialize(float a, float b, float c,
                          float d, float e, float f,
                          float g, float h, float i){
  x=a;
  y=b;
  z=c;
  lx=d;
  ly=e;
  lz=f;
  up_x=g;
  up_y=h;
  up_z=i;                            
                          
}                          

//change heading and direction of movement                        
void CFPerson::ChangeHeading(float angle){
  oldlx = lx;
  oldlz = lz;
  lx = sin(angle);
  lz = -cos(angle);  
  heading = angle;
}

//Change elevation of camera..up down
void CFPerson::ChangeElevation(float angle){
  oldly=ly;
  ly = sin(angle);
  elevation = angle;
}


//moves left or right...straffing
void CFPerson::MoveSideStep(float i){
  oldx = x;
  oldz = z;
	x = x + i * cos(heading) * sidestep_offset;
	z = z + i * sin(heading) * sidestep_offset;
}

//moves forwards/backwards
void CFPerson::MoveStep(float i){
  oldx = x;
  oldz = z;
  oldlx = lx;
  oldlz = lz;
  lx = sin(heading);
  lz = -cos(heading);
	x = x + i * lx * step_offset;
	z = z + i * lz * step_offset; 
}                               

//sets movement for each step
void CFPerson::SetStepOffset(float i){
  step_offset = i;
}

//sets movement for each sidestep
void CFPerson::SetSideStepOffset(float i){
  sidestep_offset = i;
}  
