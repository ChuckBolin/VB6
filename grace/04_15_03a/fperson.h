/*******************************************************************************
 FPERSON class
 Programmer: Chuck Bolin, 2003
 Purpose:  Controls first person data
*******************************************************************************/
#ifndef _ENGINE_FPERSON_H
#define _ENGINE_FPERSON_H
#define GLUT_DISABLE_ATEXIT_HACK

//include files
#include <gl\gl.h>
#include <gl\glut.h>   
#include <string>
#include <fstream>
#include <iostream>
#include <cmath>


// Class Definition
class CFPerson
{
  public:
 
  float x,y,z;                  //position
  float lx,ly,lz;               //look at position
  float up_x, up_y, up_z;       //up_y = 1 meaning up
  float oldx,oldy,oldz;         //stores previous position
  float oldlx,oldly,oldlz;      //stores previous look at position
  float heading, elevation;     //refers to head position
  float speed;
  int score;                    //score
  int lives;                    //indicates number of lives
  int strength;                 //strength
  float step, sidestep;
  float vel,                    //total velocity..same as speed
        vel_x,                  //velocity in x direction
        vel_y,
        vel_z,
        hdg_rate,               //rate of turn
        hdg,                    //heading or direction of travel
        elev;                   //elevation of travel
  bool headinglock ;
          
  static CFPerson& Instance()
  {
    static CFPerson instance;
    return instance;  
  }
  
  //member functions
  void ChangeHeading(float);
  void ChangeElevation(float);
  void MoveStep(float);
  void MoveSideStep(float);
  void SetStepOffset(float);
  void SetSideStepOffset(float);
  void Initialize(float, float, float, float, float, float, float, float, float);

  private:
  //constructor
  CFPerson(){
    step_offset = 0.2;
    sidestep_offset = 0.2;
    up_x = 0;
    up_y = 1;
    up_z = 0;
  }
  
  //destructor
  ~CFPerson(){}   
  
  float step_offset;
  float sidestep_offset;


  
};
#define gFPerson CFPerson::Instance()
#endif _ENGINE_FPERSON_H


/***********************************************************************
  Sample usage:
 
***********************************************************************/
