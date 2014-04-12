/*******************************************************************************
 GRAPHICS class
 Programmer: Chuck Bolin, 2003
 Purpose:  Info pertaining to game graphics.
*******************************************************************************/
#ifndef _ENGINE_GRAPHICS_H
#define _ENGINE_GRAPHICS_H
#define GLUT_DISABLE_ATEXIT_HACK

//include files
 
#include <string>
#include <fstream>
#include <ctime>
#include <iostream>
#include <strstream>     
#include <vector>
#include <gl\gl.h>
#include <gl\glut.h>   
#include "mymath.h"
#include "fperson.h"

const int GRAPHICS_ORTHO = 1;


//required for reading bitmap files
#pragma pack(1)
typedef struct {
    unsigned short type;
    unsigned long size;
    unsigned short reserved1;
    unsigned short reserved2;
    unsigned long offset;  
} BITMAPFILEHEADER;

  
typedef struct {
    unsigned long size;
    long width;
    long height;
    unsigned short planes;
    unsigned short bitcount;
    unsigned long compression;
    unsigned long sizeimage;
    long xpelspermeter;
    long ypelspermeter;
    unsigned long clrused;
    unsigned long clrimportant;
} BITMAPINFOHEADER;
#pragma pack()

//bitmap font structure for indivual characters

typedef struct
{
  int ascii;           //ascii character of font
  char chr;            //actual character
  VERTEX2D a,b,c,d;    //corners starting at top-left going ccw
} MYFONT;

// Class Definition
class CGraphics
{
  public:
    float window_width, window_height;
    float aspect_ratio;
    float near_view, far_view, view_angle;
    //float step, sidestep, heading, elevation;
    GLuint texture_objects[20];
    vector <MYFONT> font;
    
  static CGraphics& Instance()
  {
    static CGraphics instance;
    return instance;  
  }
  
  //member functions
  void SetProjection(int mode);
  void ResetProjection(int mode);
  void Print(float, float, void *,char *);
  void PrintSpaced(float,float,int, void *,char *);
  void LoadFont(void);
  void DrawText2D(int, int, int, const string& );
  void Draw2DPanel(int,int, int, int, float,float, float,int, float );
  void Draw2DPanelTexture(int,int, int, int,float, float, float, int);
  int CreateFont(int font_id); 
  int GetMaxTextures(void);
  void UpdateCamera(void);

  private:

  int font_memory_size;  //reserves space for font data
  int font_bmp;
  int max_textures ;
  
  //constructor
  CGraphics(){
    font_memory_size = 70;
    max_textures = 20;
  }
  
  //destructor
  ~CGraphics(){}   
  
};
#define gGraphics CGraphics::Instance()
#endif _ENGINE_GRAPHICS_H


/***********************************************************************
  Sample usage:

  //load font info
  int ret = gGraphics.CreateFont(56);
  if (ret == 0){
    gEventLog.LogData("Font bitmap not found!", 0);
  }
  gGraphics.LoadFont();


***********************************************************************/
