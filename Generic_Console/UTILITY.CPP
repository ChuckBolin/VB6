//first.cpp
#include "utility.h"

using namespace std;

void DrawTitle(void){
  //setColor(15,BackgroundColor);  //clears background to dark blue
  system("cls");  
  DrawBoxDouble(1,1,4,80);
  //setColor(14,BackgroundColor); //yellow
  posCursor(2, 14);
  cout << "<<< >>>";
  //setColor(11,BackgroundColor); //cyan
  posCursor(3,20);
  cout << "Written by Chuck Bolin, December 2005";
}


void PrintToFile(void){
	ofstream SaveFile("partial.cpp");

  SaveFile << " " << endl;
	SaveFile.close();
}

void PrintTimeDate(void){
  struct tm *newtime;
  time_t aclock;
  time(&aclock );   // Get time in seconds
  newtime = localtime( &aclock );   // Convert time to struct tm form 

  /* Print local time as a string */
  cout << asctime(newtime);  
  
}

void clearArea(void){
  setColor(1,1);
  for(int i = 0;i<10;i++){
    for(int j = 0; j < 60; j++){
      posCursor(36 + i, 5 + j);
      cout << " ";
    }   
  }    
  setColor(14,1);  
  DrawBoxDouble(36, 5 , 10, 60);
}

