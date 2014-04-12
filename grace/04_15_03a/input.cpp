/*******************************************************************************
 INPUT.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:    All input: mouse, keyboard comes here for processing.
*******************************************************************************/

//Include files
#include "input.h"
#include "fperson.h"



void CInput::ProcessInput (PROGRAMINPUT input){
//if(gGame.GetRenderStatus()==true){
  switch(input.event){
    case INPUT_PASSIVE_MOUSE:
      PassiveMouse(input);
      break;
    case INPUT_ACTIVE_MOUSE:
      ActiveMouse(input);
      break;
    case INPUT_EXTENDED_KEY_DOWN:
      ExtendedKeyDown(input);
      break;
    case INPUT_EXTENDED_KEY_UP:
      ExtendedKeyUp(input);
      break;
    case INPUT_NORMAL_KEY_DOWN:
      NormalKeyDown(input);
      break;
    case INPUT_TIMER:
      break;
  }
//  }
// gGame.DisableRender();
}

void CInput::PassiveMouse(PROGRAMINPUT input){
  
  //configuration mode
  if (gGame.GetGameMode() == GAMEMODE_CONFIG){
    gInput.mouse_select = 0;
    if((input.y > 100) && (input.y < 125))
      gInput.mouse_select= 1;
    if((input.y > 175) && (input.y < 200))
      gInput.mouse_select = 2;
    if((input.y > 250) && (input.y < 275))
      gInput.mouse_select = 3;
    if((input.y > 325) && (input.y < 350))
      gInput.mouse_select = 4;
  }

  //controls game play
  if (gGame.GetGameMode() == GAMEMODE_PLAY){
    int cx = static_cast<int>(gGraphics.window_width/2);
    int cy = static_cast<int>(gGraphics.window_height/2);
    if (gFPerson.headinglock == false)
    {
      if (input.x < cx - 5){
        gFPerson.heading -= .05;
        if (gFPerson.heading < 0)
          gFPerson.heading = TWO_PI;
        if (gFPerson.heading > TWO_PI)
          gFPerson.heading = 0;
        glutWarpPointer(cx,cy);
        gFPerson.ChangeHeading(gFPerson.heading);
      }
      if (input.x > cx + 5){
        gFPerson.heading += .05;
        if (gFPerson.heading < 0)
          gFPerson.heading = TWO_PI;
        if (gFPerson.heading > TWO_PI)
          gFPerson.heading = 0;
        glutWarpPointer(cx,cy);
        gFPerson.ChangeHeading(gFPerson.heading);
      }
    }
    if (input.y < cy - 5){
      gFPerson.elevation -= .05;
      glutWarpPointer(cx,cy);
      if (gFPerson.elevation < -PI/2) gFPerson.elevation = -PI/2;
      gFPerson.ChangeElevation(gFPerson.elevation);
    }
    if (input.y > cy + 5){
      gFPerson.elevation += .05;
      glutWarpPointer(cx,cy);
      if (gFPerson.elevation> PI/2) gFPerson.elevation= PI/2;
      gFPerson.ChangeElevation(gFPerson.elevation);
    }
   
    glLoadIdentity();
    gluLookAt(gFPerson.x, gFPerson.y,gFPerson.z, gFPerson.x + gFPerson.lx,
            gFPerson.y + gFPerson.ly,gFPerson.z + gFPerson.lz,
			      gFPerson.up_x, gFPerson.up_y, gFPerson.up_z);
        
  }
}

void CInput::ActiveMouse(PROGRAMINPUT input){

  if (gGame.GetGameMode() == GAMEMODE_CONFIG){
    switch(input.button){
      case GLUT_LEFT_BUTTON:
        break;
      case GLUT_RIGHT_BUTTON:  
        break;
   }
  }

  if (gGame.GetGameMode() == GAMEMODE_PLAY){
    switch(input.button){
      case GLUT_LEFT_BUTTON:
        break;
      case GLUT_RIGHT_BUTTON:  
        gFPerson.headinglock = !gFPerson.headinglock;
        break;
    }
  }
}

//Cursor keys
void CInput::ExtendedKeyDown(PROGRAMINPUT input){
  //configuration mode
  if (gGame.GetGameMode() == GAMEMODE_CONFIG){


  }
  
  //game play mode
  if (gGame.GetGameMode() == GAMEMODE_PLAY){

    //chose key
	  switch (input.key) {
		/*case GLUT_KEY_LEFT : gFPerson.sidestep=-.2;break;
		  case GLUT_KEY_RIGHT : gFPerson.sidestep=.2;break;
		  case GLUT_KEY_UP : gFPerson.step = .2;break;
		  case GLUT_KEY_DOWN : gFPerson.step = -.2;break; */

		/*  case GLUT_KEY_LEFT : gFPerson.hdg -= 1;break;
		  case GLUT_KEY_RIGHT : gFPerson.hdg += 1;break;
		  case GLUT_KEY_UP : gFPerson.vel += .02;break;
		  case GLUT_KEY_DOWN : gFPerson.vel -= .02;break;*/
 		  case GLUT_KEY_LEFT : gFPerson.hdg_rate -= .02;break;
		  case GLUT_KEY_RIGHT : gFPerson.hdg_rate += .02;break;
		  case GLUT_KEY_UP : gFPerson.vel += .02;break;
		  case GLUT_KEY_DOWN : gFPerson.vel -= .02;break;


      case GLUT_KEY_F10: break;
      default:
        gFPerson.step=0;
        gFPerson.sidestep=0;
        break;
	  }
  }
}

void CInput::ExtendedKeyUp(PROGRAMINPUT input){
 
  if (gGame.GetGameMode() == GAMEMODE_CONFIG){
 
  }
 
  if (gGame.GetGameMode() == GAMEMODE_PLAY){

    switch(input.key){
	    case GLUT_KEY_LEFT : gFPerson.sidestep = 0;break;
		  case GLUT_KEY_RIGHT :gFPerson.sidestep = 0;break;
		  case GLUT_KEY_UP : gFPerson.step = 0;break;
		  case GLUT_KEY_DOWN : gFPerson.step = 0;break;
    }
  }
}

void CInput::Timer(PROGRAMINPUT input){

}

//Normal key is pressed down
void CInput::NormalKeyDown(PROGRAMINPUT input){

  //spacebar to toggle mode
  if (input.uckey == 32){   //toggle modes
    if (gGame.GetGameMode() == GAMEMODE_CONFIG)
        gGame.SetGameMode(GAMEMODE_PLAY);
    else
        gGame.SetGameMode(GAMEMODE_CONFIG);
  }

  //L key to toggle mouse tracking of ship heading
  if (gGame.GetGameMode() == GAMEMODE_PLAY)
  {
    if ((input.uckey == 96) || (input.uckey == 108))
    {
      gFPerson.headinglock = !gFPerson.headinglock;
    }
  }  

 
  
  if (input.uckey == 27){ //quit program
    exit(0);
  }

}


