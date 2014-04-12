/*************************************************************
  
*************************************************************/
#include <cstdlib>
#include <cstdio>
#include <iostream>
#include <conio.h>
#include "graphics.h"
#include "utility.h"

using namespace std;

int main(int argc, char *argv[])
{

  int key;
 
  //screen selection  
  posCursor(12,32);
  cout << "Fullscreen? (Y/N)";
  key = getch();
  if (key == 89 || key == 121)
    FullScr();
  else  
    SetConsoleTitle("Title goes here");

  //this loop monitors keystrokes
  while(key != 27)               //exits if ESC key pressed
  {

    key=getch();                 //capture keystroke

    if (key == 0||key == 224  )  //some require two keys
      key = 256 + getch();       //get second key in this case
        
    switch(key)                  //process keystrokes
    {
      case (328):                //up arrow
        break;
      case (336):                //down arrow
        break;
      case (331):                //left arrow
        break;
      case (333):                //right arrow
        break;
      case (329):                //page up
        break;
      case (337):                //page down
        break;    
      case (71):  //G or g - Generate Source Code
      case (103):
        break;
      case (78):  //N or n - Start Notepad.exe
      case (110):
        break;  
    }
  }            
  return EXIT_SUCCESS;
}
