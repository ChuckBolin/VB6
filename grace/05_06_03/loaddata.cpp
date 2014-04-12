/*******************************************************************************
 LOADDATA.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:  Controls 3D Model file loading for .OBJ files
*******************************************************************************/

//Include files
#include "loaddata.h"

//this opens a .OBJ file and loads 3D data vectors
void CLoadData::GetObjectFromFile(string filename, vector<VERTEX3D> *pv)
{
  string line;
  streamsize count;
  
  ifstream inputfile(filename.c_str());
  char c[255];
  while(!inputfile.eof())
  {
    inputfile.getline(c,count);
    line = c;
    ExtractDataInput(line,pv);
  }  
}


//this extracts three floats for x,y,z from a space delimited string
void CLoadData::ExtractDataInput(string in, vector<VERTEX3D> *pv)
{
  int nSpace1, nSpace2, nSpace3;
  char buff1[10];
  strstream str1(buff1,10);
  char buff2[10];
  strstream str2(buff2,10);
  char buff3[10];
  strstream str3(buff3,10);

  VERTEX3D vtemp;

  if(in.substr(0,2)=="v "){  //line contains vertices

    //get space positions
    nSpace1=in.find(" ");
    nSpace2=in.find(" ",nSpace1 + 1);
    nSpace3=in.find(" ",nSpace2 + 1);
    
    if((nSpace2>nSpace1)&&(nSpace3>nSpace2)) //verify correct spacing found
    {
      //x vertex
      str1 << in.substr(nSpace1,nSpace2 - nSpace1)<<ends;
      str1 >> vtemp.x;
      str1.clear();
      
      //y vertex
      str1 << in.substr(nSpace2, nSpace3 - nSpace2-1)<<ends;
      str1 >> vtemp.y;
      str1.clear();

      //z vertex
      str1 << in. substr(nSpace3)<<ends;
      str1 >> vtemp.z;
      str1.clear();

      pv->push_back(vtemp);
    }
  }

}


