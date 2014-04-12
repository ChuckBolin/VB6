#define GLUT_DISABLE_ATEXIT_HACK
#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>
#include <cstdlib>
#include <ctime>
#include "fperson.h"
#include "graphics.h"
#include "gameobj.h"
#include <vector>
#include <cmath>

 typedef struct _asteroid 
{
 float x;
 float y;
 float z;
 float radius;
 bool visible;
 float red;
 float green;
 float blue;
} ASTEROID;

void LoadAsteroids(void);
void DrawAsteroids(void);
void DrawSphere(void);
void DrawCloud(void);
void DrawHUD(float, float,float);

