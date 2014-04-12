/*******************************************************************************
 GRAPHICS class
 Programmer: Chuck Bolin, 2003
 Purpose:  Info pertaining to game graphics.
 Some function code contributed by www.lighthouse3d.com (thanks!)
*******************************************************************************/

//Include files
#include "graphics.h"

//*******************************
// D R A W 2 D  P A N E L
// status = 0, not transparent
// status = 1, transparent
//*******************************
void CGraphics::Draw2DPanel(int x,int y, int w, int h, float r, float g, float b,
   int status, float trans)
{
  glPushMatrix();

    glDisable(GL_DEPTH_TEST);
    glDisable(GL_TEXTURE_2D);
    glDisable(GL_CULL_FACE);
    glDisable(GL_LIGHTING);
    gGraphics.SetProjection(GRAPHICS_ORTHO);
    glLoadIdentity();

    if (status==0) { //no transparency
      glColor3f(r,g,b);
      glRectd(x,y,x + w, y + h);
    }
    else{
      glEnable(GL_BLEND);
      glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA);
      glEnable(GL_BLEND);
      glColor4f(r,g,b,trans);
      glRectd(x,y,x + w, y + h);
      glColor4f(r,g,b,trans + .5);
      glDisable(GL_BLEND);
    }

    gGraphics.ResetProjection(GRAPHICS_ORTHO);
    glEnable(GL_DEPTH_TEST);
    glEnable(GL_TEXTURE_2D);
    glEnable(GL_CULL_FACE);
    glEnable(GL_LIGHTING);

  glPopMatrix();
}


//**************************************
// D R A W 2 D  P A N E L T E X T U R E
// status = 0, not transparent
// status = 1, transparent
//**************************************
void CGraphics::Draw2DPanelTexture(int x,int y, int w, int h,
                        float r, float g, float b, int status)
{
  glPushMatrix();
    glDisable(GL_DEPTH_TEST);
    glDisable(GL_TEXTURE_2D);
    glDisable(GL_CULL_FACE);
    glDisable(GL_LIGHTING);
    gGraphics.SetProjection(GRAPHICS_ORTHO);

    glLoadIdentity();
    
    if (status==0) { //no transparency
      glColor3f(r,g,b);
      glRectd(x,y,x + w, y + h);
    }
    else{
      glEnable(GL_BLEND);
      glBlendFunc(GL_ONE,GL_ONE);
      glColor3f(r,g,b);
      glRectd(x,y,x + w, y + h);
      glDisable(GL_BLEND);
    }

    float th = 1 - (h/w);
    glBegin(GL_QUADS);
    
      glTexCoord2f(0,1);
      glVertex2f(x,y );
      glTexCoord2f(0,th); 
      glVertex2f(x,y + h );
      glTexCoord2f(1,th);
      glVertex2f(x + w, y + h);
      glTexCoord2f(1,1);
      glVertex2f(x + w, y);
    glEnd();

    gGraphics.ResetProjection(GRAPHICS_ORTHO);
    glEnable(GL_DEPTH_TEST);
    glEnable(GL_CULL_FACE);
    glEnable(GL_TEXTURE);
    glEnable(GL_LIGHTING);
  glPopMatrix();
}


void CGraphics::SetProjection(int mode) {

  if (mode == GRAPHICS_ORTHO){
	  glMatrixMode(GL_PROJECTION);	    
	  glPushMatrix();             	
	  glLoadIdentity();	
	  gluOrtho2D(0,gGraphics.window_width, 0, gGraphics.window_height);	
	  glScalef(1, -1, 1);
	  glTranslatef(0, -gGraphics.window_height, 0);
	  glMatrixMode(GL_MODELVIEW);
  }
}

void CGraphics::ResetProjection(int mode) {
  if (mode == GRAPHICS_ORTHO){
    glMatrixMode(GL_PROJECTION);
	  glPopMatrix();
	  glMatrixMode(GL_MODELVIEW);
  }
}

int CGraphics::GetMaxTextures(void){
  return max_textures;
}

void CGraphics::Print(float x, float y, void *font,char *string)
{
  
  char *c;
  glRasterPos2f(x, y);
  for (c=string; *c != '\0'; c++) {
    glutBitmapCharacter(font, *c);
  }
}

void CGraphics::PrintSpaced(float x,
        float y,int spacing, void *font,char *string) {
  char *c;
  int x1=(int)x;
  for (c=string; *c != '\0'; c++) {
	glRasterPos2f(x1,y);
    glutBitmapCharacter(font, *c);
	x1 = x1 + glutBitmapWidth(font,*c) + spacing;
  }
}

int CGraphics::CreateFont(int font_id){
  font_memory_size = 70;
  font_bmp = font_id;
  if (font_bmp < 0)
    return 0;
  if ( font_bmp > max_textures )
    return 1;
  font.reserve(font_memory_size);
  return 2; //good
}

void CGraphics::LoadFont(void){
  //the font is stored in TEX_FONT1 texture...256x256 bitmap...70 characters
  //this routine calculates the required values for use as texture fonts
  MYFONT m;
  font.push_back(m);
  int count =0;
  for (int j=0;j<7;j++){
    for(int i=0;i<10;i++){
      m.a.x =.005 + i * .0508;
      m.a.y =.995 - j * .0702;
      m.b.x = .005 + i * .0508;
      m.b.y = .93 - j * .0702;
      m.c.x = .05 + i * .0508;
      m.c.y = .93 - j * .0702;
      m.d.x = .05 + i * .0508;
      m.d.y = .995 - j * .0702;
      font.push_back(m);
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
void CGraphics::DrawText2D(int x, int y, int size, const string& text ){
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
  glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[font_bmp]);
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

void CGraphics::UpdateCamera(void){
 	glLoadIdentity();
  gluLookAt(gFPerson.x, gFPerson.y,gFPerson.z, gFPerson.x + gFPerson.lx,
            gFPerson.y + gFPerson.ly,gFPerson.z + gFPerson.lz,
			      gFPerson.up_x, gFPerson.up_y, gFPerson.up_z);
}
