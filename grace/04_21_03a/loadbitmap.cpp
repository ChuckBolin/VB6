#include <string>
#include <iostream>
#include <fstream>
#include "graphics.h"
#include <gl\gl.h>
#include <gl\glu.h>
#include "eventlog.h"

unsigned char *bitmap;     
int gnWidth, gnHeight;

//********************************
// R E A D  B I T M A P 
//********************************
unsigned char *ReadBitmap(string FileName, int width, int height){

  //stores bitmap data
  BITMAPFILEHEADER h;  
  BITMAPINFOHEADER info; 

  //opens file for reading
  ifstream in(FileName.c_str(),ios::binary | ios::in); 
  if(!in){
    gEventLog.LogData(FileName + " failed to open!", 0);
    exit(EXIT_FAILURE);
  }  
  in.read((char*)&h,sizeof(BITMAPFILEHEADER));    //read file header    14
  in.read((char*)&info,sizeof(BITMAPINFOHEADER)); //read info header    40
  gnWidth = info.width;     //width of image
  gnHeight = info.height;   //height of image
    
  //datalog file info
  gEventLog.LogData("Filename loaded: " + FileName, 0);
  //gEventLog.LogData("Image Width: " + cm.NumberToString(12), 0);
  // gEventLog.LogData("Image Height: " + cm.NumberToString(info.height), 0);
  // gEventLog.LogData("Image Size: " + cm.NumberToString(info.sizeimage), 0);
    
  unsigned char buffer[info.sizeimage];      //reserves memory for image
  unsigned char *pBitmapImage = &buffer[0];  //points to mem location of image
  in.read((char*)&buffer,info.sizeimage);    //reads image data into memory
  gEventLog.LogData("Data Loaded!",0);
    
  // swap the R and B values to get RGB since the bitmap color format is in BGR
  unsigned char	tempRGB;				// swap variable
  for(int imageIdx = 0; imageIdx < info.sizeimage; imageIdx+=3){
    tempRGB = buffer[imageIdx];
	buffer[imageIdx] = buffer[imageIdx + 2];
	buffer[imageIdx + 2] = tempRGB;
  }
  gEventLog.LogData("R and B values swapped!", 0);
  return pBitmapImage;
}

//********************************
// A D D  T E X T U R E
//********************************
void  AddTexture (void){
  gluBuild2DMipmaps(GL_TEXTURE_2D, 3, gnWidth, gnHeight,
                    GL_RGB, GL_UNSIGNED_BYTE, bitmap);
  glTexParameteri(GL_TEXTURE_2D,GL_TEXTURE_MIN_FILTER,GL_LINEAR_MIPMAP_NEAREST);
  glTexParameteri(GL_TEXTURE_2D,GL_TEXTURE_MAG_FILTER,GL_LINEAR_MIPMAP_LINEAR);
  glTexImage2D(GL_TEXTURE_2D,0,3, gnWidth, gnHeight,
               0,GL_RGB,GL_UNSIGNED_BYTE,bitmap);
}
void SetupTexture(void)
{
  glPixelStorei ( GL_UNPACK_ALIGNMENT, 1 );
  glGenTextures(gGraphics.GetMaxTextures(),&gGraphics.texture_objects[0]);
}
void LoadBMPTexture(int num, string filename)
{
  //glPixelStorei ( GL_UNPACK_ALIGNMENT, 1 );
  //glGenTextures(gGraphics.GetMaxTextures(),&gGraphics.texture_objects[0]);
  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[num]);
  bitmap=ReadBitmap( filename,0,0); 
  AddTexture();
}
//********************************
// L O A D T E X T U R E 
//********************************
void LoadTextures(void)
{
  /*
  glPixelStorei ( GL_UNPACK_ALIGNMENT, 1 );
  glGenTextures(gGraphics.GetMaxTextures(),&gGraphics.texture_objects[0]);
  //AddTexture();
  //free(bitmap);   
  
  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[0]);
  bitmap=ReadBitmap( "font2.bmp",0,0); 
  AddTexture();
  //free(bitmap);
  
  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[1]);
  bitmap=ReadBitmap("stone1.bmp",0,0);  
  AddTexture();

  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[2]);
  bitmap=ReadBitmap("extwall.bmp",0,0);
  AddTexture();

  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[3]);
  bitmap=ReadBitmap("stone2.bmp",0,0);
  AddTexture();

    */
  //glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[3]);
  //bitmap=ReadBitmap("stone2.bmp",0,0);
  //AddTexture();


  //free(bitmap);   //free buffer memory  
}

// DrawBitmap
// desc: draws the bitmap image data in bitmapImage at the location
//		 x,y in the window. x and y are the lower-left corner
//		 of the bitmap.
void DrawBitmap(int x, int y, long width, long height, unsigned char* bitmapImage)
{
	glPixelStorei(GL_UNPACK_ALIGNMENT, 4);
	glRasterPos2i(x,y);
	glDrawPixels(width, height, GL_RGB, GL_UNSIGNED_BYTE, bitmapImage);
}

