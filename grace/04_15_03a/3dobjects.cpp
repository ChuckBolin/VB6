//3dobjects are drawn here
      //Includes files necessary to make everything below work
#include "stdafx.h"
//#define GLUT_DISABLE_ATEXIT_HACK 
#include <windows.h>
#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>
#include <string>
#include <iostream>
#include <cstdio>
#include <vector>
#include <ctime>
#include "3dstuff.h"
using namespace std;

//Draw box with 4 sides. Reference is northwest corner
//l = distance N-S
//w = distance E-W
//h = height
void DrawBoxEmpty(float x, float y, float z, float l, float h, float w)
{
    DrawQuad(x,y,z,EAST,w,h);   //northern face
    DrawQuad(x,y,z,SOUTH,l, h); //western face
    DrawQuad(x + w,y,z + l,WEST,w,h); //southern face
    DrawQuad(x + w,y,z + l,NORTH,l, h); //northern face
}

//Draw box with 4 sides. Reference is northwest corner. No top or bottom;
//l = distance N-S
//w = distance E-W
void DrawBoxCover(float x, float y, float z, float l, float w)
{
  glPushMatrix();
    glBegin(GL_QUADS);
      glColor3f(.6,.6,.6);
      glVertex3f(x,y,z);
      glVertex3f(x + w,y,z);
      glVertex3f(x + w,y,z + l);
      glVertex3f(x,y,z + l);        
    glEnd();
  glPopMatrix();
}

void DrawCube (float x, float y, float z, float l, float w, float h)
{
  DrawBoxEmpty(x,y,z,l,h,w);
  DrawBoxCover(x,y,z,l,w);
  DrawBoxCover(x,y - h,z,l,w); 
}

void DrawFrame (float x, float y, float z, int orient, float l, float w, float h)
{
  glPushMatrix();
    glBegin(GL_QUADS);
      
      switch(orient){
        case NORTH:
        
          break;
        case EAST:
          DrawCube(x,y,z,l,w,l);
          DrawCube(x,y - h , z,l,w,l);
          DrawCube (x,y,z,l,l,h);
          DrawCube (x + w -l,y,z,l,l,h);
          break;
        case WEST:
        
          break;
        case SOUTH:
          break;
      }   
    glEnd(); 
  glPopMatrix();

}

//Draws quad
void DrawQuad (float x, float y, float z, int orient, float l, float h)
{
  glPushMatrix();
    glBegin(GL_QUADS);
      
      switch(orient){
        case NORTH:
          glColor3f(.5,.5,.5);
          glVertex3f(x,y,z);
          glVertex3f(x,y,z - l);
          glVertex3f(x, y - h, z - l);
          glVertex3f(x,y - h, z);        
        
          break;
        case EAST:
          glColor3f(.5,.5,.5);
          glVertex3f(x,y,z);
          glVertex3f(x + l,y,z);
          glVertex3f(x + l, y - h, z);
          glVertex3f(x,y - h, z);        
          break;
        case WEST:
          glColor3f(.4,.4,.4);
          glVertex3f(x,y,z);
          glVertex3f(x - l,y,z);
          glVertex3f(x - l, y - h, z);
          glVertex3f(x,y - h, z);        
        
          break;
        case SOUTH:
          glColor3f(.4,.4,.4);
          glVertex3f(x,y,z);
          glVertex3f(x,y,z + l);
          glVertex3f(x, y - h, z + l);
          glVertex3f(x,y - h, z);        
          break;
      }   
    glEnd(); 
  glPopMatrix();
}
