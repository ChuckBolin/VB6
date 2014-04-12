/*************************************************************
  AutoCodeGen v0.1 - Written by Chuck Bolin, December 5, 2005
  Purpose: Automatically creates the code required to drive
  a robot in automous for up to 15 seconds.  All steps are 
  time based. Not position information is used. This can be
  added manually if required.
*************************************************************/
#include <cstdlib>
#include <cstdio>
#include <iostream>
#include <conio.h>
#include "graphics.h"
#include "firstrc.h"

using namespace std;

int main(int argc, char *argv[])
{
  int curr_row = 12, curr_col = 40; //place above main loop
  char choice;
  int key;
  
  posCursor(12,32);
  cout << "Fullscreen? (Y/N)";
  key = getch();
  
  if (key == 89 || key == 121)
    FullScr();
  else  
    SetConsoleTitle("FRC Team 342 Autonomous Code Generator v0.1");
    
  DrawTitle();
  DrawLabels();
  DrawInstructions();
  DrawValues();
  

                                 //this loop monitors keystrokes
  while(key != 27)               //exits if ESC key pressed
  {

    key=getch();                 //capture keystroke

    if (key == 0||key == 224  )  //some require two keys
      key = 256 + getch();       //get second key in this case
        
    switch(key)                  //process keystrokes
    {
      case (328):                //up arrow
        curr_row -= 1;
        posCursor(curr_row, curr_col);
        printf("X");
        break;
      case (336):                //down arrow
        curr_row += 1;
        posCursor(curr_row, curr_col);
        printf("X");
        break;
      case (331):                //left arrow
        curr_col -= 1;
        posCursor(curr_row, curr_col);
        printf("X");
        break;
      case (333):                //right arrow
        curr_col += 1;
        posCursor(curr_row, curr_col);
        printf("X");
        break;
      case (71):  //G or g - Generate Source Code
      case (103):
        PrintToFile();
        break;
      case (78):  //N or n - Start Notepad.exe
      case (110):
        system("c:\\winnt\\notepad.exe");
        break;  
        
    }
  }            

  //MessageBox(NULL, "That's cool!", "Good Bye", MB_OK);
  //system("PAUSE");
  return EXIT_SUCCESS;
}
