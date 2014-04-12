/*******************************************************************************
 LOADDATA class
 Programmer: Chuck Bolin, 2003
 Purpose:  Controls 3D Model file loading for .OBJ files
*******************************************************************************/
#ifndef _ENGINE_BITMAP_H
#define _ENGINE_BITMAP_H

#include <string>

unsigned char *ReadBitmap(string, int, int);
void  AddTexture (void);
void LoadTextures(void);
/*
//include files
#include <string>
#include <vector>
#include <sstream>
#include <fstream>
#include <strstream>
#include "mymath.h"

// Class Definition
class CLoadData
{
  public:

  static CLoadData& Instance()
  {
    static CLoadData instance;
    return instance;    
  }  
  
  //member functions
  void GetObjectFromFile(string, vector<VERTEX3D> *);
  void ExtractDataInput(string , vector<VERTEX3D> *);

  private:
	CLoadData()	{ }  //constructor
  ~CLoadData(){}   //destructor

};
#define gLoadData CLoadData::Instance()
*/
#endif _ENGINE_BITMAP_H

/***********************************************************************
  Sample usage:
  
  vector<VERTEX3D> gv;  
  int main()
  {
    gLoadData.GetObjectFromFile("test1.obj", &gv);
    vector<VERTEX3D>::const_iterator pos;
    for(pos=gv.begin();pos!=gv.end();++pos)
    {
      cout<< pos->x <<" " << pos->y << " " << pos->z <<endl;
    }
  system("pause");
  exit(0);
}
***********************************************************************/
