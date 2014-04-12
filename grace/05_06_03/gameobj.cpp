/*******************************************************************************
 GAMEOBJ.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:    Game program details...necessary for control
*******************************************************************************/

//Include files
#include <string>
#include <strstream>
#include "gameobj.h"

//sets game mode...play,config
void CGameObject::SetGameMode(int mode){
  game_mode = mode;
}

//returns game mode
int CGameObject::GetGameMode(void){
  return game_mode;
}

//instructs class to update and calculate FPS
void CGameObject::Update (int ntime){
  if (frame==0)
  {
    timebase=ntime;
  }
  frame++;
  if (frame>29)
  {
    fps = (30000 / (ntime-timebase));
    //sprintf(cFPS,"FPS: %d",(int)FPS);
    frame=0;
  }  
  time_factor= 40/fps;  //on my PC this is 1
}

//returns fps (float)
float CGameObject::GetFPS(void){
  return fps;
}

//returns fps as STL string
string CGameObject::GetFPSString(void){
  char buffer[50];
  ostrstream Str(buffer, 50);
  Str << fps << ends;
  string sFPS(Str.str());
  return sFPS;
}

//sets time factor...correction to rendering so all PCs are same rendering spd
void CGameObject::SetTimeFactor(float tfactor){
  time_factor = tfactor;
}

//get current value
float CGameObject::GetTimeFactor(void){
  return time_factor;
}

void CGameObject::EnableRender(void){
  render = true;
}

void CGameObject::DisableRender(void){
  render = false;
}

bool CGameObject::GetRenderStatus(void){
  return render;
}

