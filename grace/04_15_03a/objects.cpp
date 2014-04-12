#include <gl\gl.h>
#include <gl\glut.h>
#include <gl\glu.h>

#include <vector>
#include <string>
#include "mymath.h"
#include "3dstuff.h"
#include <iostream>
#include <sstream>
#include <ctime>
#include <cmath>



extern GLuint texture_objects[20];   //decrease number to reserve less mry
extern int gnWidth, gnHeight;
extern int gnWindowWidth, gnWindowHeight;
extern unsigned char *bitmap;
extern CGameObject cam;
extern float gfHeading;
extern float gfElevation;

//stores all 3D model data
extern int gnVertices; //total number of vertices in program
extern int gnTriangles; //total number of triangles in program
extern int gnNormals;
extern int gnObjects;
extern vector <VERTEX3D> ver;   //stores all vertices
extern vector <TRIANGLE3D> tri; //stores all triangles...listing their vertices
extern vector <VERTEX3D> norm;
extern vector <OBJECTTRACKER> obj;
extern vector <MYFONT> font;
extern int gnMaxTextures;

extern float gfZAngle;
extern float gfXAngle;
extern float gfYAngle;

//extern float gfGround[20][20];
extern float grid[100][100];
// L O A D F O N T
void LoadFont(void){
  //the font is stored in TEX_FONT1 texture...256x256 bitmap...70 fonts
  //this routine calculates the required values for use as texture fonts
  int count =0;
  for (int j=0;j<7;j++){
    for(int i=0;i<10;i++){
      ++count;
      font[count].a.x =.005 + i * .0508;
      font[count].a.y =.995 - j * .0702; 
      font[count].b.x = .005 + i * .0508;
      font[count].b.y = .93 - j * .0702;
      font[count].c.x = .05 + i * .0508;
      font[count].c.y = .93 - j * .0702;
      font[count].d.x = .05 + i * .0508;
      font[count].d.y = .995 - j * .0702;
   }
 }

 //manually assign ASCII values to each font character
 font[1].ascii = 65; //A
 font[2].ascii = 66; 
 font[3].ascii = 67; 
 font[4].ascii = 68; 
 font[5].ascii = 69; 
 font[6].ascii = 70; 
 font[7].ascii = 71; 
 font[8].ascii = 72; 
 font[9].ascii = 73; 
 font[10].ascii = 74; //J
 font[11].ascii = 75; //K
 font[12].ascii = 76; 
 font[13].ascii = 77; 
 font[14].ascii = 78; 
 font[15].ascii = 79; 
 font[16].ascii = 80; 
 font[17].ascii = 81; 
 font[18].ascii = 82; 
 font[19].ascii = 83; 
 font[20].ascii = 84; //T
 font[21].ascii = 85; //U
 font[22].ascii = 86; 
 font[23].ascii = 87; 
 font[24].ascii = 88; 
 font[25].ascii = 89; 
 font[26].ascii = 90; //Z
 
 font[27].ascii = 48; //0
 font[28].ascii = 49; 
 font[29].ascii = 50; 
 font[30].ascii = 51; 
 font[31].ascii = 52; 
 font[32].ascii = 53; 
 font[33].ascii = 54; 
 font[34].ascii = 55; 
 font[35].ascii = 56; 
 font[36].ascii = 57;//9 
 font[37].ascii = 46; //.  
 
 font[38].ascii = 44; 
 font[39].ascii = 60; 
 font[40].ascii = 62; 
 font[41].ascii = 91; 
 font[42].ascii = 93; 
 font[43].ascii = 63; 
 font[44].ascii = 34; 
 font[45].ascii = 34; 
 font[46].ascii = 59; 
 font[47].ascii = 58; 
 font[48].ascii = 126; 
 font[49].ascii = 47; 
 font[50].ascii = 92; 
 font[51].ascii = 124; 
 font[52].ascii = 33; 
 font[53].ascii = 45; 
 font[54].ascii = 43; 
 font[55].ascii = 42; 
 font[56].ascii = 38; 
 font[57].ascii = 37; 
 font[58].ascii = 36; 
 font[59].ascii = 35; 
 font[60].ascii = 61; 
 font[61].ascii = 40; 
 font[62].ascii = 41; 
 font[63].ascii = 32; //space
 font[64].ascii = 32;
}



//********************************
// D R A W T E X T
//********************************
void DrawText2D(int x, int y, int size, const string& text ){
  int step = 0;
  if (size<8) size=8;  //minimum size
  int num[text.length()+1];  //stores required numbers for font
  int count=0;
  bool bTest;

  //determine required mapped values for character in string
  for(string::const_iterator m = text.begin(); m != text.end(); ++m){
    bTest=true;  
    for (int z=1;z < 64; z++){
      if(font[z].ascii==(int)*m ){
        num[count] = z; //stores character from map
        count++;

      }
    }
  }
  
  //draw texture on a quad for each character in text
  glEnable(GL_TEXTURE_2D);
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_FONT1]);
  glBegin(GL_QUADS);
    int z;
    for(int k=0;k<count; k++){
      z= num[k];
      glTexCoord2f(font[z].a.x, font[z].a.y);
      glVertex2f(x + step * size, y );
      glTexCoord2f(font[z].b.x, font[z].b.y);
      glVertex2f(x + step * size, y + size );
      glTexCoord2f(font[z].c.x, font[z].c.y);
      glVertex2f(x + (step + 1) * size, y + size );
      glTexCoord2f(font[z].d.x, font[z].d.y);
      glVertex2f(x + (step + 1) * size, y );
      step++;
    }
  glEnd();     
  glDisable(GL_TEXTURE_2D);
}



//********************************
//D R A W 3 D   O B J E C T 
//********************************
void Draw3DObject(int object, int texture, float red, float green, float blue, float x, float y, float z, float anglex, float angley, float anglez, float trans)
{
  glPushMatrix();
    glTranslatef(x,y,z);
    glRotatef(anglex,1,0,0);
    glRotatef(angley,0,1,0);
    glRotatef(anglez,0,0,1);
    glEnable(GL_TEXTURE_2D);
    glBindTexture ( GL_TEXTURE_2D, texture);
    glColor4f(red,green,blue,trans);

    for (int i= obj[object].first;i < obj[object].last + 1;i++){
      glBegin(GL_TRIANGLES);
        if(tri[i].textureside==LEFT){
          glNormal3f(norm[tri[i].normal].x,norm[tri[i].normal].y, norm[tri[i].normal].z);
          glTexCoord2f(0,1);
          glVertex3f( ver[tri[i].a].x,ver[tri[i].a].y,ver[tri[i].a].z);
          glTexCoord2f(0,0); 
          glVertex3f( ver[tri[i].b].x,ver[tri[i].b].y,ver[tri[i].b].z);
          glTexCoord2f(1,1);
          glVertex3f( ver[tri[i].c].x,ver[tri[i].c].y,ver[tri[i].c].z);
        }
        if(tri[i].textureside==RIGHT){
          glNormal3f(norm[tri[i].normal].x,norm[tri[i].normal].y, norm[tri[i].normal].z);
          glTexCoord2f(1,1);
          glVertex3f( ver[tri[i].a].x,ver[tri[i].a].y,ver[tri[i].a].z);
          glTexCoord2f(0,0); 
          glVertex3f( ver[tri[i].b].x,ver[tri[i].b].y,ver[tri[i].b].z);
          glTexCoord2f(1,0);
          glVertex3f( ver[tri[i].c].x,ver[tri[i].c].y,ver[tri[i].c].z);
        }
      glEnd();
    glEnable(GL_TEXTURE_2D);  
    }
  glPopMatrix();
}


//********************************
//L O A D   3 D   O B J E C T S
//********************************
void Load3DObjects(void)
{

  VERTEX3D a,b,c,d;
  VERTEX3D va,vb,vc,vd;
  
  //northwest corner
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  a.x=-14;
  a.y=0;
  a.z=-14;
  LoadCylinder (a, 2,10,16,10,TEX_EXTWALL,1,.8,0);
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;
  
  //northeast corner
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  a.x=14;
  a.y=0;
  a.z=-14;
  LoadCylinder (a, 2,10,16,10,TEX_EXTWALL,1,.8,0);
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;

  //wall
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  for(int j= 0;j<6;j++){
    for(int i= -15;i<15;i ++){
      va.x=-15; //west wall
      va.y=j+1;
      va.z=i+1;
      
      vb.x=-15;
      vb.y=j;
      vb.z=i+1;

      vc.x=-15;
      vc.y=j;
      vc.z=i;

      vd.x=-15;
      vd.y=j+1;
      vd.z=i;
  
      LoadQuad(va,vb,vc,vd,TEX_EXTWALL,1,.8,0);
      
      va.x=i; //north wall
      va.y=j+1;
      va.z=-15;
      
      vb.x=i;
      vb.y=j;
      vb.z=-15;

      vc.x=i+1;
      vc.y=j;
      vc.z=-15;

      vd.x=i+1;
      vd.y=j+1;
      vd.z=-15;
  
      LoadQuad(va,vb,vc,vd,TEX_EXTWALL,1,.8,0);


      va.x=15; //east wall
      va.y=j+1;
      va.z=i;
      
      vb.x=15;
      vb.y=j;
      vb.z=i;

      vc.x=15;
      vc.y=j;
      vc.z=i+1;

      vd.x=15;
      vd.y=j+1;
      vd.z=i+1;
  
      LoadQuad(va,vb,vc,vd,TEX_EXTWALL,1,.8,0);
      
    }
  }
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;

  //snowman
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  a.x=0;
  a.y=.8;
  a.z=-2;
  LoadSphere(a,1,22,22,TEX_STONE2,1,1,1);
  a.x=0;
  a.y=2.4;
  a.z=-2;
  LoadSphere(a,.7,22,22,TEX_STONE2,1,1,1);
  a.x=0;
  a.y=3.3;
  a.z=-2;
  LoadSphere(a,.4,22,22,TEX_STONE2,1,1,1);
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;

  //ground
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  LoadGroundReference();
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;

  //draws cylinder for legs and arms
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  a.x=0;
  a.y=0;
  a.z=-1;
  LoadCylinder (a, .06,.5,8,5,0,0,.8,.8);
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;

  //draws cylinder for joints
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  a.x=0;
  a.y=0;
  a.z=-1;
  LoadSphere2 (a, .08, 5,5, 0,1,0,0);
  obj[gnObjects].render = false;
  obj[gnObjects].last = gnTriangles;


  //preload grid array for ground
  gnObjects += 1; //increment object counter
  obj[gnObjects].first = gnTriangles + 1;
  int s,t;
  float num;
  srand(time(NULL));

  //generates random vertical Y values
  for(int i=1;i<1000;i++){
    s=(int)((float)(rand()/524288) * 99);
    t=(int)((float)(rand()/524288) * 99);
    //LogData("Grid: " + IntegerToString(rand()) + FloatToString(((float)(rand()/32768) * 99)));
    num = -10 + ((float)rand()/524288) * 20; //vertical height

    if( (s>0)&&(s<100)&&(t>0)&&(t<100)){
      //grid[s][t]=num;
    }
  }
   for(int j=1;j<98;j++){
    for(int i=1;i<98;i++){
      num =(sin(i * PI/50) * sin(j * PI/22)* 20 +  (((float)rand()/32768) * 2));
      //num =(sin(i * PI/50) * sin(j * PI/98)* 10   +  (((float)rand()/32768) * 2/i));


      grid[i][j]=num ;
    }
   }
    //renders grid from array
  int thistexture=0;
  for(int j=1;j<99;j++){
    for(int i=1;i<99;i++){
      //num = ((float)rand()/524288) * 20;
      //grid[i][j]=num * .05;
       a.x=(i*1);
      a.y=grid[i][j];
      a.z=(j*1);
      
      b.x=(i*1);
      b.y=grid[i][j+1];
      b.z=(j*1)+1;
      
      c.x=(i*1)+1;
      c.y=grid[i+1][j+1];
      c.z=(j*1)+1;
      
      d.x=(i*1)+1;
      d.y=grid[i+1][j];
      d.z=(j*1);

      
      num=grid[i][j];  
      
      if (num > 15.0f){
        thistexture=TEX_SNOW1;
      }  
      else if(num > 0){
        float val = ((float)rand()/32768) * 10;
        if(val > 0.5){
          thistexture=TEX_GRASS2;
        }
        else{
          thistexture=TEX_STONE3;
        }  
      }  

      else if(num > -10){
        float val = ((float)rand()/32768) * 10;
        if(val > sin(i * PI/98) * 7){
          thistexture=TEX_STONE3;
        }
        else{
          thistexture=TEX_STONE;
        }  
      }  
      
      
      else{
        thistexture=TEX_STONE;
      }
      
      
      /*
      else(num == 10.0f){
        thistexture=TEX_STONE;
      }  
      else(num ==10.01f){
        thistexture=TEX_GRASS1;
      }
      */        
      LoadQuad(a, b, c, d, thistexture, 1,1,1);    
    }
  }
  obj[gnObjects].render = true;
  obj[gnObjects].last = gnTriangles;


}

//***************************************************
// L O A D  S I M P L E  C U B E
// One texture is used for each side
//***************************************************
void LoadSimpleCube(VERTEX3D a, float side, float scale, int texture_n, int texture_e,
                    int texture_s, int texture_w, int texture_t, int texture_b,
                    float red, float green, float blue)
{
  VERTEX3D va, vb, vc, vd;
  float off=side/2;//sqrt( side * side + side * side);
  
  //south side of cube
  va.x=a.x - off;
  va.y=a.y + off;
  va.z=a.z + off;
  vb.x=a.x - off;
  vb.y=a.y - off;
  vb.z=a.z + off;
  vc.x=a.x + off;
  vc.y=a.y - off;
  vc.z=a.z + off;
  vd.x=a.x + off;
  vd.y=a.y + off;
  vd.z=a.z + off;
  LoadQuad(va,vb,vc,vd,texture_s, red, green, blue);

  //east side of cube
  va.x=a.x + off;
  va.y=a.y + off;
  va.z=a.z + off;
  vb.x=a.x + off;
  vb.y=a.y - off;
  vb.z=a.z + off;
  vc.x=a.x + off;
  vc.y=a.y - off;
  vc.z=a.z - off;
  vd.x=a.x + off;
  vd.y=a.y + off;
  vd.z=a.z - off;
  LoadQuad(va,vb,vc,vd,texture_e, red, green, blue);

  //north side of cube
  va.x=a.x + off;
  va.y=a.y + off;
  va.z=a.z - off;
  vb.x=a.x + off;
  vb.y=a.y - off;
  vb.z=a.z - off;
  vc.x=a.x - off;
  vc.y=a.y - off;
  vc.z=a.z - off;
  vd.x=a.x - off;
  vd.y=a.y + off;
  vd.z=a.z - off;
  LoadQuad(va,vb,vc,vd,texture_n, red, green, blue);

  //west side of cube
  va.x=a.x - off;
  va.y=a.y + off;
  va.z=a.z - off;
  vb.x=a.x - off;
  vb.y=a.y - off;
  vb.z=a.z - off;
  vc.x=a.x - off;
  vc.y=a.y - off;
  vc.z=a.z + off;
  vd.x=a.x - off;
  vd.y=a.y + off;
  vd.z=a.z + off;
  LoadQuad(va,vb,vc,vd,texture_w, red, green, blue);

  //top side of cube
  va.x=a.x - off;
  va.y=a.y + off;
  va.z=a.z - off;
  vb.x=a.x - off;
  vb.y=a.y + off;
  vb.z=a.z + off;
  vc.x=a.x + off;
  vc.y=a.y + off;
  vc.z=a.z + off;
  vd.x=a.x + off;
  vd.y=a.y + off;
  vd.z=a.z - off;
  LoadQuad(va,vb,vc,vd,texture_t, red, green, blue);

  //bottom side of cube
  va.x=a.x - off;
  va.y=a.y - off;
  va.z=a.z + off;
  vb.x=a.x - off;
  vb.y=a.y - off;
  vb.z=a.z - off;
  vc.x=a.x + off;
  vc.y=a.y - off;
  vc.z=a.z - off;
  vd.x=a.x + off;
  vd.y=a.y - off;
  vd.z=a.z + off;
  LoadQuad(va,vb,vc,vd,texture_t, red, green, blue);

}                       

//***************************************************
// L O A D S P H E R E
// One texture is repeated for each quad in sphere
//***************************************************
void LoadSphere (VERTEX3D a, float radius, float sections, float disks, int texture, float red, float green, float blue)
{

  float dA, dH;
  VERTEX3D e,f,g,h; //stores newly calculated 
  dA = TWO_PI/sections;
  dH = PI / disks;
  float dHor1, dVer1;
  float dHor2, dVer2;
  
  for (int j=0;j<disks;j++){
    for (int i=0;i<sections;i++){
      dHor1= radius * sin(dH * j);
      dVer1 = radius * cos(dH * j);
      dHor2= radius * sin(dH * (j+1));
      dVer2 = radius * cos(dH * (j+1));

    
      e.x = a.x + dHor1 * cos(dA * i);
      e.y = a.y + dVer1;
      e.z = a.z + dHor1 * -sin(dA * i);
              
      f.x = a.x + dHor2 * cos(dA * i);
      f.y = a.y + dVer2;
      f.z = a.z + dHor2 * -sin(dA * i);
    
      g.x = a.x + dHor2 * cos(dA * (i + 1));
      g.y = a.y + dVer2;
      g.z = a.z + dHor2 * -sin(dA * (i + 1));
              
      h.x = a.x + dHor1 * cos(dA * (i + 1));
      h.y = a.y + dVer1;
      h.z = a.z + dHor1 * -sin(dA * (i + 1));

      if ((j>0)&&(j<disks - 1)){
        LoadQuad(e,f,g,h, texture, red, green, blue);
      }
      if (j==0){
        LoadTriangle(e,f,h, texture, LEFT, red, green, blue);
      }
      if (j==disks - 1){
        LoadTriangle(e,f,h, texture, LEFT, red, green, blue);
      }
      
      
    }
  }  
}

//***************************************************
// L O A D S P H E R E 2
// One texture is stretched across portion of sphere
//***************************************************
void LoadSphere2 (VERTEX3D a, float radius, float sections, float disks, int texture, float red, float green, float blue)
{

  float dA, dH;
  VERTEX3D e,f,g,h; //stores newly calculated 
  dA = TWO_PI/sections;
  dH = PI / disks;
  float dHor1, dVer1;
  float dHor2, dVer2;
  
  for (int j=0;j<disks;j++){
    for (int i=0;i<sections;i++){
      dHor1= radius * sin(dH * j);
      dVer1 = radius * cos(dH * j);
      dHor2= radius * sin(dH * (j+1));
      dVer2 = radius * cos(dH * (j+1));

    
      e.x = a.x + dHor1 * cos(dA * i);
      e.y = a.y + dVer1;
      e.z = a.z + dHor1 * -sin(dA * i);
              
      f.x = a.x + dHor2 * cos(dA * i);
      f.y = a.y + dVer2;
      f.z = a.z + dHor2 * -sin(dA * i);
    
      g.x = a.x + dHor2 * cos(dA * (i + 1));
      g.y = a.y + dVer2;
      g.z = a.z + dHor2 * -sin(dA * (i + 1));
              
      h.x = a.x + dHor1 * cos(dA * (i + 1));
      h.y = a.y + dVer1;
      h.z = a.z + dHor1 * -sin(dA * (i + 1));

      if ((j>0)&&(j<disks - 1)){
        LoadQuad(e,f,g,h, texture, red, green, blue);
      }
      if (j==0){
        LoadTriangle(e,f,h, texture, LEFT, red, green, blue);
      }
      if (j==disks - 1){
        LoadTriangle(e,f,h, texture, LEFT, red, green, blue);
      }
      
      
    }
  }  
}


void LoadCylinder (VERTEX3D a, float radius, float height, float sections, float disks, int texture, float red, float green, float blue)
{
  float dA, dH;
  VERTEX3D e,f,g,h; //stores newly calculated 
  dA = TWO_PI/sections;
  dH = height/disks;
  
  for (int j=0;j<disks;j++){
    for (int i=0;i<sections;i++){
      g.x = a.x + radius * cos(dA * (i + 1));
      g.y = a.y + (j * dH);
      g.z = a.z + radius * sin(dA * (i + 1));
              
      f.x = a.x + radius * cos(dA * i);
      f.y = a.y + (j * dH);
      f.z = a.z + radius * sin(dA * i);
    
      h.x = a.x + radius * cos(dA * (i + 1));
      h.y = a.y + ((j + 1) * dH);
      h.z = a.z + radius * sin(dA * (i + 1));
              
      e.x = a.x + radius * cos(dA * i);
      e.y = a.y + ((j + 1) * dH);
      e.z = a.z + radius * sin(dA * i);
    

    LoadQuad(e,h,g,f, texture, red, green, blue);
    }  
  }
  




}


void LoadQuadAngle(VERTEX3D a, float l, float h, float x,float y, float z, int texture, float red, float green, float blue)
{
  VERTEX3D b,c,d;
  d.x = a.x + l * cos(z)* sin(y);
  d.y = a.y + l * sin(z);
  d.z = a.z;

  b.x = a.x + h * sin(z);
  b.y = a.y - h * cos(z);
  b.z = a.z;

  c.x = a.x + (l * cos(z) * sin(y)) + (h * sin(z));
  c.y = a.y - (h * cos(z)) + (l * sin(z));
  c.z = a.z;

  LoadQuad(a,b,c,d,texture,red,green,blue);  


}



//********************************
// L O A D P L A N E
//
//   a-------c
//   |       |
//   |       |
//   b-------'
//
// This function takes a large plane such as a wall and reduces it into small quads
// which are then converted into small triangle and loaded into memory for rendering
//********************************
void LoadPlane(VERTEX3D a, VERTEX3D b, VERTEX3D c, int texture, int unit, float red, float green, float blue)
{
  VERTEX3D e,f,g,h;
  float tx,ty,tz; //deltas between a and c on top border
  float lx,ly,lz; //deltas between a and b on left border
  float dt,dl;    //distances between the top and left vertices 
  int stept, stepl; 
  
  tx = fabs(a.x - c.x); 
  ty = fabs(a.y - c.y);
  tz = fabs(a.z - c.z);
  dt = sqrt( tx*tx + ty*ty + tz*tz); //dist between a and c vertices
  tx = tx/dt;  
  ty = ty/dt;
  tz = tz/dt;
  
  lx = fabs(a.x - b.x);
  ly = fabs(a.y - b.y);
  lz = fabs(a.z - b.z);
  dl = sqrt (lx*lx + ly*ly + lz*lz); //dist between a and b vertices
  lx = lx/dl;
  ly = ly/dl;
  lz = lz/dl;
  
  stept = (int) (dt/unit);
  stepl = (int) (dl/unit);
  
  for (int j=0;j< stepl;j++){
    for (int i=0;i<stept;i++){
      e.x = a.x + (i * unit);
      f.x = a.x + (i * unit);
      g.x = a.x + ((i + 1) * unit);
      h.x = a.x + ((i + 1) * unit);
      
      e.y = a.y - (j * unit);
      f.y = a.y - ((j + 1) * unit);
      g.y = a.y - ((j + 1) * unit);
      h.y = a.y - (j * unit);
      
      e.z = a.z ;//+ (i * unit);
      f.z = a.z ;//+ (i * unit);
      g.z = a.z ;//+ (i * unit);
      h.z = a.z ;//+ (i * unit);

      LoadQuad(e,f,g,h,texture,red,green,blue);  
    }
  }  
}




//********************************
//L O A D Q U A D
//
//   a-------d
//   |       |
//   |       |
//   b-------c
//********************************
void LoadQuad(VERTEX3D a, VERTEX3D b, VERTEX3D c, VERTEX3D d, int texture, float red, float green, float blue)
{
  LoadTriangle(a,b,d,texture, LEFT, red,green, blue);
  LoadTriangle(d,b,c,texture, RIGHT, red,green,blue);
}

//********************************
//L O A D T R I A N G L E
//
//   LEFT Side     RIGHT Side
//   a----c              a
//   |   /              /|
//   |  /              / |
//   | /              /  |
//   b               b---c
//********************************
void LoadTriangle(VERTEX3D a, VERTEX3D b, VERTEX3D c, int texture, int side,float red, float green, float blue)
{
  //a,b,c= vertices
  //texture = ID of particular texture
  //side = NW corner (LEFT) or SE corner (RIGHT)
  float v[3][3];
  float out[3];

  //add vertices to ver vector
  gnVertices += 3;
  ver[gnVertices-2]=a;
  ver[gnVertices-1]=b;
  ver[gnVertices]=c;
  
  //add vertex index values to triangle
  gnTriangles += 1;           
  tri[gnTriangles].a=gnVertices-2;
  tri[gnTriangles].b=gnVertices-1;
  tri[gnTriangles].c=gnVertices;
  tri[gnTriangles].red=red;
  tri[gnTriangles].green=green;
  tri[gnTriangles].blue=blue;
  tri[gnTriangles].texture = texture;
  tri[gnTriangles].textureside = side;

  //calculate and retrieve normals
  v[0][0]=ver[gnVertices - 2].x;
  v[0][1]=ver[gnVertices - 2].y;
  v[0][2]=ver[gnVertices - 2].z;
  v[1][0]=ver[gnVertices - 1].x;
  v[1][1]=ver[gnVertices - 1].y;
  v[1][2]=ver[gnVertices - 1].z;
  v[2][0]=ver[gnVertices].x;
  v[2][1]=ver[gnVertices].y;
  v[2][2]=ver[gnVertices].z;
  CalcNormal(v,out);
  gnNormals += 1;
  norm[gnNormals].x = out[0]; //loads calculated normals 
  norm[gnNormals].y = out[1];
  norm[gnNormals].z = out[2];
  tri[gnTriangles].normal = gnNormals;
}

//*****************************************
// L O A D  G R O U N D  R E F E R E N C E
//*****************************************
void LoadGroundReference(void)
{
  VERTEX3D a,b,c,d,e,f,g,h;
  float fTemp[6][6];
  float dx,dz;
  float fFactor=.4;
  
  //height in units in a 20x20 world
  float fGround[20][20] = {
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,5,0,0,3,0,0,3,3,3,3,3,3,3,3,3,5,5,5},
  {1,0,5,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,5,0,4,3,6,6,6,2,2,0,0,1,2,3,4,4,5,5},
  {1,0,5,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,5,5,4,3,6,6,6,2,2,0,1,1,2,0,9,9,0,0},
  {1,0,5,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,0,5,4,3,6,6,5,1,1,0,1,1,2,0,8,9,6,0},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,0,5,4,3,6,6,5,1,1,0,1,1,2,3,8,9,6,0},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,0,0,5,4,4,5,0,0,0,0,3,3,1,2,3,8,9,6,0},
  {1,0,0,0,0,0,0,0,4,0,0,0,0,4,4,4,4,4,5,6},
  {0,0,0,5,4,0,0,0,0,0,0,0,3,1,2,3,8,5,6,0},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,7,6,5,5,4,0,3,3,3,3,0,0,1,2,3,4,5,6,0},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6},
  {0,7,6,5,0,5,0,0,0,0,0,0,0,1,0,3,4,5,6,6},
  {1,0,0,0,0,2,0,0,4,4,4,4,4,4,4,4,4,4,5,6}
  
    };


  //convert data into triangles
  //int l,k;
  //for (int k=0;k<3;k++){
    //for(int l=0;l<3;l++){
      for (int j=0;j<20;j++){
        for (int i=0;i<20;i++){
      
      //4 major vertices
      a.x=-20 + j*2;
      a.y=fFactor * fGround[j][i];
      a.z= -15 + i*2;
      
      b.x= -20 + j*2;
      b.y=fFactor * fGround[j][i+1];
      b.z=-15 + (i+1)*2;
      
      c.x= -20 + (j+1)*2;
      c.y=fFactor * fGround[j+1][i+1];
      c.z= -15 + (i+1)*2;

      d.x=-20 + (j+1)*2;
      d.y=fFactor * fGround[j+1][i];
      d.z=-15 + i*2;

      if (a.y< 0) a.y=0;
      if (b.y< 0) b.y=0;
      if (c.y< 0) c.y=0;
      if (d.y< 0) d.y=0;
      
      
      LoadQuad(a,b,c,d,TEX_STONE2,.8,.8,1);
      
      /*
      //extracts 25 smaller quads from one larger quad
      //uses vertical values
      for (int l=0;l<5;l++){
        dx=fabs(d.z - a.z);
        dz=fabs(b.z - a.z);
        for (int k=0;k<5;k++){
          fTemp[l][k]= a.y + dx;
          fTemp[k][l]= a.y + dz;       
        }
      }      
      
      for (int l=0;l<5;l++){
        for (int k=0;k<5;k++){
          e.x=a.x + k;
          e.y=a.y + fTemp[l][k];
          e.z=a.z + l;
          f.x=b.x + k;
          f.y=b.y + fTemp[l+1][k];
          f.z=b.z + l + 1;
          g.x=c.x + k + 1;
          g.y=c.y + fTemp[l+1][k+1];
          g.z=c.z + l + 1;
          h.x=d.x + k + 1;
          h.y=d.y + fTemp[l][k+1];
          h.z=d.z + l;
          LoadQuad(e,f,g,h,TEX_TREE,0,1,.3);

        }
      }      
      */
        }
      }  
   // }
  //}  
}

//********************************
// L O A D T E X T U R E 
//********************************
void LoadTextures(void)
{
  LogData("Commence loading textures...");
  glPixelStorei ( GL_UNPACK_ALIGNMENT, 1 );
  glGenTextures(gnMaxTextures,&texture_objects[0]); //<<<<<<<<<<<<
  AddTexture();
  
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_STONE]);
  bitmap=ReadBitmap("stone1.bmp",gnWidth,gnHeight); 
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_STONE2]);
  bitmap=ReadBitmap("stone2.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");
  
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_SNOW1]);
  bitmap=ReadBitmap("snow1.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");
  
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_SNOW3]);
  bitmap=ReadBitmap("snow3.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_METAL1]);
  bitmap=ReadBitmap("metal1.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_EXTWALL]);
  bitmap=ReadBitmap("extwall.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_FLOOR]);
  bitmap=ReadBitmap("floor.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");
  
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_GRASS1]);
  bitmap=ReadBitmap("grass1.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_IVY1]);
  bitmap=ReadBitmap("ivy1.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_FONT1]);
  bitmap=ReadBitmap("font1.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");
  
  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_STONE3]);
  bitmap=ReadBitmap("stone3.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  glBindTexture ( GL_TEXTURE_2D, texture_objects[TEX_GRASS2]);
  bitmap=ReadBitmap("grass2.bmp",gnWidth,gnHeight);  
  AddTexture();
  LogData("texture loaded");

  
  free(bitmap);   //free buffer memory  
  LogData("bitmap memory buffer freed.");

}



//********************************
//A D D T E X T U R E 
//********************************
void  AddTexture (void){
  gluBuild2DMipmaps(GL_TEXTURE_2D, 3, gnWidth, gnHeight, GL_RGB, GL_UNSIGNED_BYTE, bitmap);
  glTexParameteri(GL_TEXTURE_2D,GL_TEXTURE_MIN_FILTER,GL_LINEAR_MIPMAP_NEAREST);
  glTexParameteri(GL_TEXTURE_2D,GL_TEXTURE_MAG_FILTER,GL_LINEAR_MIPMAP_LINEAR);
  glTexImage2D(GL_TEXTURE_2D,0,3,gnWidth,gnHeight,
    0,GL_RGB,GL_UNSIGNED_BYTE,bitmap);
}

//********************************
//D R A W  H U D
//********************************
void DrawHUD(void){
    glDisable(GL_DEPTH_TEST);
    glDisable(GL_TEXTURE_2D);
    glDisable(GL_CULL_FACE);
    glDisable(GL_LIGHTING);
    glPushMatrix();
	  SetOrthographicProjection();
    glLoadIdentity();

    float cx,cy;
    cx = gnWindowWidth/2;
    cy = gnWindowHeight/2;

    glEnable(GL_BLEND);
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA);
    glEnable(GL_BLEND);
    glColor4f(0,.3,1,.4);

    
    glBegin(GL_LINES);
      glColor4f(.7,.7,0,.3);
      //glColor3f(0,0,0);
      for (int q=0;q<9;q++){             //add tick marks to crosshairs
        glVertex2f(cx - 10 ,cy - 80 + (q * 20)); //vertical tic marks
        glVertex2f(cx + 10 ,cy - 80 + (q * 20));
        glVertex2f(cx - 80 + (q * 20),cy - 10 ); //horizontal tic marks
        glVertex2f(cx - 80 + (q * 20),cy + 10 );
      }
      glVertex2f(cx -95,cy-95);        //top line
      glVertex2f(cx + 95, cy - 95);
      glVertex2f(cx -95,cy+95);        //bottom line
      glVertex2f(cx + 95, cy + 95);
      glVertex2f(cx - 95,cy -95);       //left line
      glVertex2f(cx - 95,cy + 95);
      glVertex2f(cx + 95,cy -95);       //right line
      glVertex2f(cx + 95,cy + 95);
      glVertex2f(cx , cy - 95);            //vertical line
      glVertex2f(cx , cy + 95);
      glVertex2f(cx - 95, cy);            //horizontal line
      glVertex2f(cx + 95, cy); 
      glColor4f(0,.7,.7,.3);
      glVertex2f(cx - 96, cy - 96);        //top line
      glVertex2f(cx + 96, cy - 96);
      glVertex2f(cx - 96, cy + 96);        //bottom line
      glVertex2f(cx + 96, cy + 96);
      glVertex2f(cx - 96, cy - 96);       //left line
      glVertex2f(cx - 96, cy + 96);
      glVertex2f(cx + 96, cy - 96);       //right line
      glVertex2f(cx + 96, cy + 96);
      glVertex2f(cx + 1 , cy - 95);            //vertical line
      glVertex2f(cx + 1 , cy + 95);
      glVertex2f(cx - 95, cy - 1);            //horizontal line
      glVertex2f(cx + 95, cy - 1); 
      for (int q=0;q<9;q++){             //add tick marks to crosshairs
        glVertex2f(cx - 10 ,cy - 81 + (q * 20)); //vertical tic marks
        glVertex2f(cx + 10 ,cy - 81 + (q * 20));
        glVertex2f(cx - 79 + (q * 20),cy - 10 ); //horizontal tic marks
        glVertex2f(cx - 79 + (q * 20),cy + 10 );
      }
    glEnd();

    glBegin(GL_TRIANGLES);
      glColor4f(0,.3,1,.4);
      glVertex2f(cx + cos(gfHeading - PI/2)*80, cy + sin(gfHeading-PI/2)*80);
      glVertex2f(cx + cos((gfHeading - 90) - PI/2)*5, cy + sin((gfHeading - 90) -PI/2)*5);
      glVertex2f(cx,cy);               
      glColor4f(0,.3,1,.4);
      glVertex2f(cx + cos(gfHeading - PI/2)*80, cy + sin(gfHeading-PI/2)*80);
      glVertex2f(cx,cy);               
      glVertex2f(cx + cos((gfHeading + 90) - PI/2)*5, cy + sin((gfHeading + 90) -PI/2)*5);
      glColor4f(0,.3,1,.4);
      glVertex2f(cx - 82, cy + 0 - 90 *(gfElevation/(PI/2)));
      glVertex2f(cx - 92, cy - 5 - 90 *(gfElevation/(PI/2)));
      glVertex2f(cx - 92, cy + 5 - 90 *(gfElevation/(PI/2)));
    glEnd();

    glDisable(GL_BLEND);   
    ResetPerspectiveProjection();  
    glEnable(GL_DEPTH_TEST);
    glEnable(GL_TEXTURE_2D);
    glEnable(GL_CULL_FACE);
    glEnable(GL_LIGHTING);
    glPopMatrix();
    Draw2DPanel((int)(cx - 100), (int)(cy - 100), 200,200, 0, 0, .2, 1, .15);


}




