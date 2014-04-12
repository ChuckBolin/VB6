//graphics.cpp
#include "graphics.h"

//This code was found at this site.
//http://forum.armkb.com/archive/index.php/t-34.html
void FullScr(void){
  typedef BOOL (WINAPI *PROCSETCONSOLEDISPLAYMODE)(HANDLE,DWORD,LPDWORD);
  typedef BOOL (WINAPI *PROCGETCONSOLEDISPLAYMODE)(LPDWORD);
  PROCSETCONSOLEDISPLAYMODE SetConsoleDisplayMode;
  PROCGETCONSOLEDISPLAYMODE GetConsoleDisplayMode;
  HMODULE hKernel32 = GetModuleHandle("kernel32");
  SetConsoleDisplayMode = (PROCSETCONSOLEDISPLAYMODE )
  GetProcAddress(hKernel32,"SetConsoleDisplayMode");
  GetConsoleDisplayMode = (PROCGETCONSOLEDISPLAYMODE)
  GetProcAddress(hKernel32,"GetConsoleDisplayMode");
  HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
  DWORD dwOldMode;
  SetConsoleDisplayMode(hOut,1,&dwOldMode);
}


void setColor(int foreground, int background)
{
  HANDLE hOut;
  hOut = GetStdHandle(STD_OUTPUT_HANDLE);
  SetConsoleTextAttribute(hOut, background * 16 + foreground); 
}  


void posCursor( int row, int column)
{
  HANDLE hOut;
  COORD cursorpos;
  hOut = GetStdHandle(STD_OUTPUT_HANDLE);
  cursorpos.X = column - 1;
  cursorpos.Y = row - 1;
  SetConsoleCursorPosition(hOut, cursorpos); 
}  

//r range 1 to 25, c range 1 to 80
//height range = 25 - r    width range = 80 - c
void DrawBoxDouble(int r, int c, int height, int width)
{
  int i;

  height -=1;
  width -=1;

  //draw corners  
  posCursor(r,c + width); printf ("%c",187);   //top right corner
  posCursor(r + height,c + width); printf("%c",188);//bottom right 
  posCursor(r,c); printf("%c",201);        //top left corner
  posCursor(r + height,c); printf("%c",200);   //bottom left corner
  
  //draw vertical sides
  for (i = 1;i<height;i++)
  {
    posCursor(r + i, c); printf("%c", 186);
    posCursor(r + i, c + width); printf("%c", 186);
  }  
  
  //draw horizontal sides
  for (i=1;i<width;i++)
  {
    posCursor(r,c + i); printf("%c", 205);
    posCursor(r + height, c + i); printf("%c",205);    
  }    
}
