/*******************************************************************************
 GAMEOBJ class
 Programmer: Chuck Bolin, 2003
 Purpose:  Game program info.
*******************************************************************************/
#ifndef _ENGINE_GAMEOBJ_H
#define _ENGINE_GAMEOBJ_H
#include <string>

const int GAMEMODE_PAUSE = 1;
const int GAMEMODE_CONFIG = 2;
const int GAMEMODE_PLAY = 3;

// Class Definition
class CGameObject
{
  public:
  float background_color_red;
  float background_color_green;
  float background_color_blue;
  float near_view;
  float far_view;
  float view_angle;
  bool grid_on;
  int grid_min_x;
  int grid_max_x;
  int grid_min_z;
  int grid_max_z;
  float grid_red;
  float grid_green;
  float grid_blue;
  bool stars;
  bool planet;
  bool moons;
  bool asteroids;
  bool hud;
  int star_bmp;
  int icon_bmp;
  int font_bmp;
  string app_path;  
  
  static CGameObject& Instance()
  {
    static CGameObject instance;
    return instance;  
  }

  //public members
  void Update (int);
  float GetFPS(void);
  string GetFPSString(void);
  void SetTimeFactor(float);
  float GetTimeFactor(void);
  void SetGameMode(int);
  int GetGameMode(void);
  void EnableRender(void);
  void DisableRender(void);
  bool GetRenderStatus(void);

  private:

  //constructor
  CGameObject(){}
  
  //destructor
  ~CGameObject(){}
 
  //private variables
  int frame, timef, timebase;
  float fps;
  char  char_fps[6];
  float time_factor;
  int game_mode;
  bool render;

};
#define gGame CGameObject::Instance()
#endif _ENGINE_GAMEOBJ_H


/***********************************************************************
  Sample usage:

***********************************************************************/

