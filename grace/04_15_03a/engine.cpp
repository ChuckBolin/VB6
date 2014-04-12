/*********************************************************************
 PROGRAM - Written by Chuck Bolin
 This program uses the accompanying Grace Engine v0.1
*********************************************************************/

//Includes files
#define GLUT_DISABLE_ATEXIT_HACK
#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>
#include <string>
#include <iostream>
#include <cstdio>
#include <vector>
#include <ctime>
#include "eventlog.h"
#include "program.h"
#include "fperson.h"
#include "graphics.h"
#include "bitmap.h"
#include "mymath.h"
#include "gameobj.h"
#include "engine.h"
#include "input.h"
#include "script.h"
#include "mymath.h"

#include "drawing.h"

using namespace std;

//****************************************************************************
//  Function prototypes
//****************************************************************************
void LoadFont(void);
void SetupTexture(void);
void LoadBMPTexture(int, string);

//****************************************************************************
//  Global Variables and constants
//****************************************************************************
PROGRAMINFO prog;     //stores program details..name, version, email, etc.

GLuint nDL;           //display list
GLfloat specular [] = { 1.0, 1.0, 1.0, 1.0 };
GLfloat shininess [] = { 122.0 };
GLfloat position [] = { 15.0, 1.75, 0.0, 0.0 };

string bmp[25];
int bmp_count;

typedef struct
{
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
}SCRIPT;

SCRIPT gs;
mymath math;


//****************************************************************************
//  Initialization
//****************************************************************************


//************************************
//Initialize objects and variables
//************************************
void InitializeVariables()
{
  //initialize graphics
  gGraphics.near_view = .01;
  gGraphics.far_view = 200;
  gGraphics.view_angle = 45;
  
  //initialize first person
  gFPerson.Initialize(0,1.5,0,0,0,-1,0,1,0);
  gFPerson.vel = 0;
  gFPerson.hdg = 0;
  gFPerson.hdg_rate = 0;
  gFPerson.vel_x = 0;
  gFPerson.vel_z = 0;
  gFPerson.headinglock = true;
    
  //grid info
  gs.grid_on = true;
  gs.grid_min_x = -25;
  gs.grid_max_x = 25;
  gs.grid_min_z = -25;
  gs.grid_max_z = 25;
  gs.grid_red = 0.4;
  gs.grid_green = 0.4;
  gs.grid_blue = 0.4;

  //program specific information
  prog.version_number = gProgram.SetVersion("0.1");
  prog.program_name = gProgram.SetProgramName("Grace Engine");
  prog.program_description = gProgram.SetProgramDescription("This is a basic program.");
  prog.revision_date = gProgram.SetRevisionDate("01.16.03");
  prog.programmer_name = gProgram.SetProgrammerName("Chuck Bolin");
  prog.programmer_email = gProgram.SetProgrammerEmail("cbolin@dycon.com");
  prog.programmer_URL = gProgram.SetProgrammerURL("http://www.clg-net.com");
  prog.help_file = gProgram.SetHelpFile("help.htm");

  //game program specifics
  gGame.SetTimeFactor(1.0);
  gGame.SetGameMode(GAMEMODE_PLAY);
}

string FloatToString(float n){
  char buffer[50];
  ostrstream Str(buffer, 50);
  Str << n << ends;
  string mynum(Str.str());
  return mynum;
}

//****************************************************************************
//  Graphics
//****************************************************************************

void InitializeLighting(void){
  glMaterialfv(GL_FRONT_AND_BACK, GL_SPECULAR, specular);
  glMaterialfv(GL_FRONT_AND_BACK, GL_SHININESS, shininess);

  // Set the GL_AMBIENT_AND_DIFFUSE color state variable to be the
  // one referred to by all following calls to glColor
  glColorMaterial(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE);
  glEnable(GL_COLOR_MATERIAL);

  // Create a Directional Light Source
  glLightfv(GL_LIGHT0, GL_POSITION, position);
  glEnable(GL_LIGHTING);
  glEnable(GL_LIGHT0);
}

//**************************
// Initializes scene
//**************************
void InitializeGraphics() {
  glClearColor(0,0,0,1);
  glFrontFace(GL_CCW);		// Counter clock-wise polygons face out
  glCullFace(GL_BACK);
  glEnable(GL_CULL_FACE);
  glDepthFunc(GL_LESS);
  glEnable(GL_DEPTH_TEST);
  glEnable ( GL_TEXTURE_2D );
  SetupTexture();
  for(int i = 0; i < bmp_count; i++)
  {
    LoadBMPTexture(i,bmp[i]);
  }  
  glMatrixMode(GL_PROJECTION);
  glLoadIdentity();
  gluPerspective(gGraphics.view_angle,gGraphics.aspect_ratio,
                 gGraphics.near_view, gGraphics.far_view);
  glMatrixMode(GL_MODELVIEW);
  glLoadIdentity();
  glutSetCursor(GLUT_CURSOR_NONE); 
  gGraphics.window_width = glutGet(GLUT_WINDOW_WIDTH);
  gGraphics.window_height = glutGet(GLUT_WINDOW_HEIGHT);
  glutWarpPointer(static_cast<int>(gGraphics.window_width/2), 
                  static_cast<int>(gGraphics.window_height/2));
  gGame.EnableRender();
  nDL = glGenLists(25);  //creates space for 25 lists
  glNewList(1,GL_COMPILE);
   DrawSphere();
  glEndList();
  glNewList(2,GL_COMPILE);
    //DrawCloud();
    DrawAsteroids();
  glEndList();  
  InitializeLighting();
}

//*************************
//Changes size of screen
//*************************
void ResizeWindow(int w1, int h1)
	{

	// Prevent a divide by zero, when window is too short
	// (you cant make a window of zero width).
	if(h1 == 0)
		h1 = 1;
	gGraphics.aspect_ratio = 1.0f * w1 / h1;

	// Reset the coordinate system before modifying
	glMatrixMode(GL_PROJECTION);
	glLoadIdentity();
	
	// Set the viewport to be the entire window
    glViewport(0, 0, w1, h1);

	// Set the clipping volume
	gluPerspective(gGraphics.view_angle,gGraphics.aspect_ratio,
                   gGraphics.near_view, gGraphics.far_view);
	glMatrixMode(GL_MODELVIEW);
	glLoadIdentity();
	gluLookAt(gFPerson.x, gFPerson.y,gFPerson.z, gFPerson.x + gFPerson.lx,
            gFPerson.y + gFPerson.ly,gFPerson.z + gFPerson.lz,
			      gFPerson.up_x, gFPerson.up_y, gFPerson.up_z);
}




//*************************************************************************
//*************************************************************************
//*************************************************************************
//*************************************************************************
//      R E N D E R      R E N D E R     R E N D E R     R E N D E R
//*************************************************************************
//*************************************************************************
//*************************************************************************
//*************************************************************************

void RenderScene(void) {
 
  //if(gGame.GetRenderStatus()==true){
  gGraphics.window_width = glutGet(GLUT_WINDOW_WIDTH);
  gGraphics.window_height = glutGet(GLUT_WINDOW_HEIGHT);

  //allows for setup of program by user in run mode
  if(gGame.GetGameMode() == GAMEMODE_CONFIG){
    //glEnable(GL_LIGHTING);
    //constructs two panels
    glPushMatrix();
      //glClearColor(0,.3,.3,0);         //reinsert to clear background
      //glClear(GL_COLOR_BUFFER_BIT);
      glLoadIdentity();
      gGraphics.Draw2DPanel(100,100,600,400,0,0,.02,0,.01);
      //gGraphics.Draw2DPanel(0,0,static_cast<int>(gGraphics.window_width),
      //            static_cast<int>(gGraphics.window_height),0,0,.02,1,.01);
    glPopMatrix();

    //writes text
    glDisable(GL_LIGHTING);
    glPushMatrix();
    glClearColor(0,0,0,0);
    glClear(GL_DEPTH_BUFFER_BIT);  //allows text to be written
      gGraphics.SetProjection(GRAPHICS_ORTHO);
      glLoadIdentity();

      switch(gInput.mouse_select){
        case 0:
          glColor3f(1,1,1);
          gGraphics.DrawText2D(110,100,25, "OPTION1");
          gGraphics.DrawText2D(110,175,25, "OPTION2");
          gGraphics.DrawText2D(110,250,25, "OPTION3");
          gGraphics.DrawText2D(110,325,25, "CONTINUE...");
          break;
        case 1:
          glColor3f(0,1,0);
          gGraphics.DrawText2D(110,100,25, "OPTION1");
          glColor3f(1,1,1);
          gGraphics.DrawText2D(110,175,25, "OPTION2");
          gGraphics.DrawText2D(110,250,25, "OPTION3");
          gGraphics.DrawText2D(110,325,25, "CONTINUE...");
          break;
        case 2:
          glColor3f(0,1,0);
          gGraphics.DrawText2D(110,175,25, "OPTION2");
          glColor3f(1,1,1);
          gGraphics.DrawText2D(110,100,25, "OPTION1");
          gGraphics.DrawText2D(110,250,25, "OPTION3");
          gGraphics.DrawText2D(110,325,25, "CONTINUE...");
          break;
        case 3:
          glColor3f(0,1,0);
          gGraphics.DrawText2D(110,250,25, "OPTION3");
          glColor3f(1,1,1);
          gGraphics.DrawText2D(110,100,25, "OPTION1");
          gGraphics.DrawText2D(110,175,25, "OPTION2");
          gGraphics.DrawText2D(110,325,25, "CONTINUE...");
          break;
        case 4:
          glColor3f(0,1,0);
          gGraphics.DrawText2D(110,325,25, "CONTINUE...");
          glColor3f(1,1,1);
          gGraphics.DrawText2D(110,100,25, "OPTION1");
          gGraphics.DrawText2D(110,175,25, "OPTION2");
          gGraphics.DrawText2D(110,250,25, "OPTION3");
          break;
      }
      gGraphics.ResetProjection(GRAPHICS_ORTHO);
    glPopMatrix();
}

//render game play
if(gGame.GetGameMode() == GAMEMODE_PLAY){

  //fps
  gGame.Update(glutGet(GLUT_ELAPSED_TIME));

  //Player movement  
 	glLoadIdentity();
  gluLookAt(gFPerson.x, gFPerson.y,gFPerson.z, gFPerson.x + gFPerson.lx,
            gFPerson.y + gFPerson.ly,gFPerson.z + gFPerson.lz,
			      gFPerson.up_x, gFPerson.up_y, gFPerson.up_z);
  
  glClearColor(gs.background_color_red,
               gs.background_color_green,
               gs.background_color_blue,
               0);
  glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT);
  
  //draws ground reference      
  
  
  if (gs.grid_on == true)
  {
    glBegin(GL_LINES);
      glColor3f(gs.grid_red,gs.grid_green,gs.grid_blue);
      for (int i = gs.grid_min_x; i < gs.grid_max_x;i++)
      {
        glVertex3f(i,0,gs.grid_min_z);
        glVertex3f(i,0,gs.grid_max_z);
      }
      for (int j = gs.grid_min_z;j < gs.grid_max_z;j++)
      {
        glVertex3f(gs.grid_min_x,0,j);   
        glVertex3f (gs.grid_max_x,0,j);   
      }
    glEnd();
  }
        
    /*
    //render all triangles
    glPushMatrix();
      glEnable(GL_TEXTURE_2D);
      glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[3]);
       glBegin(GL_TRIANGLES);
        glColor3f(1,1,1);
          glTexCoord2f(0,1);
          glVertex3f(-2,3,-2);
          glTexCoord2f(0,0); 
          glVertex3f(-2,0,-2 );
          glTexCoord2f(1,1);
          glVertex3f(2,3,-2 );
          glTexCoord2f(1,1);
          glVertex3f(2,3,-2 );
          glTexCoord2f(0,0); 
          glVertex3f(-2,0,-2 );
          glTexCoord2f(1,0);
          glVertex3f(2,0,-2 );
      glEnd();
      glDisable(GL_TEXTURE_2D);
    glPopMatrix();
     */
    
    glEnable(GL_LIGHTING);
    glPushMatrix();
      glTranslatef(0,1.75,-7);
      glCallList(1);
    glPopMatrix();
    glPushMatrix();
      glTranslatef(0,2,-100);
      glColor3f(0,0,1);
      glutSolidSphere(3,10,10);
    glPopMatrix();
    
    glCallList(2);
    
    glPushMatrix();
      glLoadIdentity();
      //gGraphics.Draw2DPanel(0,500,800,100,0,0,0,0,.01);
      glColor3f(0,0,0);
      DrawHeadingIndicator(math.RadToDeg(gFPerson.heading),gFPerson.hdg);
    glPopMatrix();
    
    //writes text
    glDisable(GL_LIGHTING);
    glPushMatrix();
      glClearColor(0,0,0,0);
      glClear(GL_DEPTH_BUFFER_BIT);  //allows text to be written
      gGraphics.SetProjection(GRAPHICS_ORTHO);
      glLoadIdentity();
      glColor3f(0,1,0);
      gGraphics.DrawText2D(10,10,20, "FPS: " + gGame.GetFPSString() );
      //gGraphics.DrawText2D(10,518,20,"Heading: " + FloatToString(math.RadToDeg(gFPerson.heading)));
      //gGraphics.DrawText2D(10,540,20, "Hdg: " + FloatToString(gFPerson.hdg));
      //gGraphics.DrawText2D(10,562,20, "Vel: " + FloatToString(gFPerson.vel));
      //gGraphics.DrawText2D(410,540,20, "Vel X: " + FloatToString(gFPerson.vel_x));
      //gGraphics.DrawText2D(410,562,20, "Vel Y: " + FloatToString(gFPerson.vel_y));
     gGraphics.ResetProjection(GRAPHICS_ORTHO);
    glPopMatrix();
    glEnable(GL_LIGHTING);
    
}
     glutSwapBuffers();  
}
//*************************************************************************
//  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ 
//  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ 
//  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ 
//*************************************************************************


//****************************************************************************
//  Input Functions - Keyboard and Mouse
//****************************************************************************

void MousePassiveMotion(int x, int y){
  PROGRAMINPUT input;
  input.event = INPUT_PASSIVE_MOUSE;
  input.x = x;
  input.y = y;
  gInput.ProcessInput(input);
}
//*******************************************************************
// Movement and button clicks - M O U S E  A C T I V E
//*******************************************************************
void MouseFunction(int button, int state, int x, int y)
{
  PROGRAMINPUT input;
  input.event = INPUT_ACTIVE_MOUSE;
  input.button = button;
  input.state = state;
  input.x = x;
  input.y = y;
  gInput.ProcessInput(input);
}

//*****************************************
//Special keys called by glutSpecialFunc()
//*****************************************
void PressExtendedKey(int key, int x, int y) {
  PROGRAMINPUT input;
  input.event = INPUT_EXTENDED_KEY_DOWN;
  input.key = key;
  input.x = x;
  input.y = y;
  gInput.ProcessInput(input);
}

//******************************************
//Special keys called by glutSpecialUpFunc()
//******************************************
void ReleaseExtendedKey(int key, int x, int y){
  PROGRAMINPUT input;
  input.event = INPUT_EXTENDED_KEY_UP;
  input.key = key;
  input.x = x;
  input.y = y;
  gInput.ProcessInput(input);
}

//**********************************************
//Normal key event called by glutKeyboardFunc()
//**********************************************
void PressNormalKey(unsigned char key, int x, int y) {
  PROGRAMINPUT input;
  input.event = INPUT_NORMAL_KEY_DOWN;
  input.uckey = key;
  input.x = x;
  input.y = y;
  gInput.ProcessInput(input);
}

//****************************************************************************
// Timed Interrupts
//****************************************************************************

//Timer function....called each second
void TimerUpdate(int value){

  if((gGame.GetGameMode() == GAMEMODE_PLAY) && (value == 2))
  {
    gFPerson.hdg += gFPerson.hdg_rate;
    if (gFPerson.hdg < 0) 
      gFPerson.hdg = 360;
    if (gFPerson.hdg > 360)
      gFPerson.hdg = 0;
    gFPerson.vel_x = gFPerson.vel * sin(math.DegToRad( gFPerson.hdg ));
    gFPerson.vel_z = -gFPerson.vel * cos(math.DegToRad( gFPerson.hdg ));
    gFPerson.x += gFPerson.vel_x;
    gFPerson.z += gFPerson.vel_z;
    
    //auto reset
    if (gFPerson.x > 50 | gFPerson.x < -50 | gFPerson.z > 50 | gFPerson.z < -50)
    {
      gFPerson.x = 0;
      gFPerson.z = 0;
      gFPerson.vel = 0;
      gFPerson.hdg = 0;
      gFPerson.hdg_rate = 0;
    }
    if(gFPerson.headinglock == true)
    {
      
      gFPerson.heading = math.DegToRad(gFPerson.hdg);            
      gFPerson.ChangeHeading(gFPerson.heading);
    }
    glutTimerFunc(50, TimerUpdate,2);
  }
  
  //gGame.EnableRender();
  //gGame.Update(glutGet(GLUT_ELAPSED_TIME));
  //game clock
  
  /*
   gnTimeLeft -= value; //one second countdown timer
  if (gnTimeLeft<1) {
    gnTimeLeft=0;
  }
  */
  //keep calling timer each second
  
  
}
/***************************************************
 P R O C E S S _ S C R I P T _ F I L E
***************************************************/
void ProcessScriptFile(void)
{

  //output script function names and their matching parameters if any
  PARAM par;
  gScript.ParseFile();
  
  //process one line at a time
  for(int i=0;i< gScript.GetSize();i++)
  {
    //this structure holds name, args and argv
    par = gScript.GetScriptLine(i);

    //script setup stuff
    if (par.name == "background_color")
    {  
      gs.background_color_red = math.StringToFloat(par.argv[0]);
      gs.background_color_green = math.StringToFloat(par.argv[1]);
      gs.background_color_blue = math.StringToFloat(par.argv[2]);
    }  
    if (par.name == "near_view")
    {
      gs.near_view = math.StringToFloat(par.argv[0]);
      gGraphics.near_view = gs.near_view;
    }  
    if (par.name == "far_view")
    {
      gs.far_view = math.StringToFloat(par.argv[0]);
      gGraphics.far_view = gs.far_view;
    }  
    if (par.name == "view_angle")
    {
      gs.view_angle = math.StringToFloat(par.argv[0]);
      gGraphics.view_angle = gs.view_angle;
    }
    if (par.name == "fperson_initialize")
    {
      gFPerson.Initialize(math.StringToFloat(par.argv[0]), // x
                          math.StringToFloat(par.argv[1]), // y
                          math.StringToFloat(par.argv[2]), // z
                          math.StringToFloat(par.argv[3]), // lx
                          math.StringToFloat(par.argv[4]), // ly
                          math.StringToFloat(par.argv[5]), // lz
                          math.StringToFloat(par.argv[6]), // up_x
                          math.StringToFloat(par.argv[7]), // up_y
                          math.StringToFloat(par.argv[8]));// up_z
    }
    if (par.name == "fperson_vel")
      gFPerson.vel = math.StringToFloat(par.argv[0]);
    if (par.name == "fperson_hdg")
      gFPerson.hdg = math.StringToFloat(par.argv[0]);
    if (par.name == "fperson_hdg_rate")
      gFPerson.hdg_rate = math.StringToFloat(par.argv[0]);
    if (par.name == "fperson_vel_x")
      gFPerson.vel_x = math.StringToFloat(par.argv[0]);
    if (par.name == "fperson_vel_z")
      gFPerson.vel_z = math.StringToFloat(par.argv[0]);
    if (par.name == "fperson_headinglock")
    {
      if (par.argv[0] == "true")
      {
        gFPerson.headinglock = true;
      }
      else
      {
        gFPerson.headinglock = false;
      }  
    }
    if (par.name == "grid_on")
    {
      if (par.argv[0] == "true")
      {
        gs.grid_on = true;
      }
      else
      {
        gs.grid_on = false;
      }
    }  
    if (par.name == "grid_min_x")
    {
      gs.grid_min_x = math.StringToInteger(par.argv[0]);
    }
    if (par.name == "grid_max_x")
    {
      gs.grid_max_x = math.StringToInteger(par.argv[0]);
    }
    if (par.name == "grid_min_z")
    {
      gs.grid_min_z = math.StringToInteger(par.argv[0]);
    }
    if (par.name == "grid_max_z")
    {
      gs.grid_max_z = math.StringToInteger(par.argv[0]);
    }
    if (par.name == "grid_red")
    {
      gs.grid_red = math.StringToFloat(par.argv[0]);
    }
    if (par.name == "grid_green")
    {
      gs.grid_green = math.StringToFloat(par.argv[0]);
    }
    if (par.name == "grid_blue")
    {
      gs.grid_blue = math.StringToFloat(par.argv[0]);
    }
    if (par.name == "program_name")
      prog.program_name = gProgram.SetProgramName(par.argv[0]);
    if (par.name == "version")
      prog.version_number = gProgram.SetVersion(par.argv[0]);    
    if (par.name == "program_description")
      prog.program_description = gProgram.SetProgramDescription(par.argv[0]);
    if (par.name == "program_date")
      prog.revision_date = gProgram.SetRevisionDate(par.argv[0]);
    if (par.name == "program_programmer")
      prog.programmer_name = gProgram.SetProgrammerName(par.argv[0]);
    if (par.name == "program_email")
      prog.programmer_email = gProgram.SetProgrammerEmail(par.argv[0]);
    if (par.name == "program_URL")
      prog.programmer_URL = gProgram.SetProgrammerURL(par.argv[0]);
    if (par.name == "program_helpfile")
      prog.help_file = gProgram.SetHelpFile(par.argv[0]);
    if (par.name == "program_logfile")
      gEventLog.SetFilename(par.argv[0]);    
    if (par.name == "load_texture")
    {
      bmp[bmp_count] = par.argv[1];
      gEventLog.LogData("Texture ID: " + par.argv[0] + " BMP Loaded: " + par.argv[1],0);
      bmp_count++;    
      //LoadBMPTexture(math.StringToInteger(par.argv[0]), par.argv[1]);
      //gEventLog.LogData("Texture ID: " + par.argv[0] + " BMP Loaded: " + par.argv[1],0);
    }  
  }
  cout << "Starting application...." << endl;
}

/************************************************************************
  main() - This program begins here!  
************************************************************************/
int main(int argc, char **argv)
{
  //determine if script file exists
  string scriptfile;
  
  if (argc >1 ){
    scriptfile = argv[1];
    gScript.DisplayResults(scriptfile);
  
    if (gScript.GetNumErrors() > 0){
      cout << "Errors reading script file. Aborting program..." << endl;
      cout << "Press any key to continue." << endl;
      system("pause");
      exit(0);
    }
  }
                         
  //setup event logging
  InitializeVariables();
  gEventLog.SetFilename("log2.txt"); //default filename
  ProcessScriptFile(); 
  
  //load font info
  int ret = gGraphics.CreateFont(0);
  if (ret == 0){
    gEventLog.LogData("Font bitmap not found!", 0);
  }
  gGraphics.LoadFont();
  
  gEventLog.LogData("Font loaded",0);
  //SetupTexture();
   
  //datalog info
  gEventLog.LogData("**********************************",0);
  gEventLog.LogData("Start of Program.",0);
  gEventLog.LogData("**********************************",0);
  gEventLog.LogData("Version: " + prog.version_number, 0);
  gEventLog.LogData("Program: " + prog.program_name,0);
  gEventLog.LogData("Description: " + prog.program_description,0);
  gEventLog.LogData("Revision Date: " + prog.revision_date,0);
  gEventLog.LogData("Programmer: " + prog.programmer_name,0);
  gEventLog.LogData("Email: " + prog.programmer_email,0);
  gEventLog.LogData("URL: " + prog.programmer_URL,0);
  gEventLog.LogData("Help file: " + prog.help_file,0);

  //OpenGL/GLUT graphical outputs
  gEventLog.LogData("Initialize OpenGL/GLUT Graphics.",0);
  glutInit(&argc, argv);
  glutInitDisplayMode( GLUT_RGBA | GLUT_DOUBLE | GLUT_DEPTH );
  glutInitWindowSize(800,600);
  glutInitWindowPosition(0,0);
  glutCreateWindow(prog.program_name.c_str());
  glutSetCursor(GLUT_CURSOR_NONE);
  glutDisplayFunc(RenderScene);
  glutReshapeFunc(ResizeWindow);
  InitializeGraphics();
  srand(time(NULL));  //randomizer

  //OpenGL/GLUT inputs - keyboard and mouse
  gEventLog.LogData("Initialize OpenGL/GLUT Keyboard and Mouse.",0);
  glutKeyboardFunc(PressNormalKey);
  glutSpecialFunc(PressExtendedKey);
  glutSpecialUpFunc(ReleaseExtendedKey);     
  glutPassiveMotionFunc(MousePassiveMotion);
  glutMouseFunc(MouseFunction);
    
  //Miscellaneous stuff
  gEventLog.LogData("Initialize Misc. Stuff.",0);
  glutIdleFunc(RenderScene);
  glutTimerFunc(33,TimerUpdate,1);
  glutTimerFunc(50,TimerUpdate,2); //updating ships position

  //GLUT Main loop
  gEventLog.LogData("Begin glutMainLoop().",0);
  glutMainLoop();
}


