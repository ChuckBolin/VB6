#include "drawing.h"


/***********************************************
 head = heading of head, ship = ships heading
***********************************************/
void DrawHUD(float head, float ship, float elev)
{
  int s = 185;
  glPushMatrix();

    glDisable(GL_DEPTH_TEST);
    glDisable(GL_TEXTURE_2D);
    glDisable(GL_CULL_FACE);
    glDisable(GL_LIGHTING);
    gGraphics.SetProjection(GRAPHICS_ORTHO);
    glLoadIdentity();

    //black panel display
    //glColor3f(0,.4,.4);
    //glRectd(158,438,622,502);
    //glColor3f(0,0,0);
    //glRectd(160,440,620,500);

    //icons
    glPushMatrix();
      glEnable(GL_TEXTURE_2D);
      glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[gGame.icon_bmp]);
      glBegin(GL_QUADS);
        glColor3f(1,1,0);
        glTexCoord2f(0,1);
        glVertex2f(s + 20 ,450);
        glTexCoord2f(0,.5);
        glVertex2f(s + 20,470);
        glTexCoord2f(.5, .5);
        glVertex2f(s + 40, 470);
        glTexCoord2f(.5,1);
        glVertex2f(s + 40, 450);
      glEnd();
      glDisable(GL_TEXTURE_2D);  
    glPopMatrix();
    glPushMatrix();
      glEnable(GL_TEXTURE_2D);
      glBindTexture ( GL_TEXTURE_2D, gGraphics.texture_objects[gGame.icon_bmp]);
      glBegin(GL_QUADS);
        glColor3f(1,1,1);
        glTexCoord2f(0,0.5);
        glVertex2f(s +20,480);
        glTexCoord2f(0,0);
        glVertex2f(s+20,500);
        glTexCoord2f(0.5, 0);
        glVertex2f(s + 40, 500);
        glTexCoord2f(.5,.5);
        glVertex2f(s + 40, 480);
      glEnd();
      glDisable(GL_TEXTURE_2D);  
    glPopMatrix();        
 
    //draws markers
    glBegin(GL_TRIANGLES);
    
      //ship heading
      glColor3f(1,1,1);
      if (ship < 181)
      {
        glVertex2f(s + 210 + ( 160 * ship / 180 ), 478);
        glVertex2f(s + 205 + ( 160 * ship / 180), 497);
        glVertex2f(s + 215 + ( 160 * ship / 180), 497);

      }
     else 
      {
        glVertex2f(s + 50 + ( 160 * (ship - 180) / 180 ), 478);
        glVertex2f(s + 45 + ( 160 * (ship - 180) / 180 ), 497);
        glVertex2f(s + 55 + ( 160 * (ship - 180) / 180 ), 497);
      }
 
      //head heading...look around
      glColor3f(1,1,0);
      if (head < 181)
      {
        glVertex2f(s + 210 + ( 160 * head / 180 ), 453);
        glVertex2f(s + 205 + ( 160 * head / 180), 472);
        glVertex2f(s + 215 + ( 160 * head / 180), 472);        
      }
     else 
      {
        glVertex2f(s + 50 + ( 160 * (head - 180) / 180 ), 453);
        glVertex2f(s + 45 + ( 160 * (head - 180) / 180 ), 472);
        glVertex2f(s + 55 + ( 160 * (head - 180) / 180 ), 472);
      }
    glEnd();

    //prints reference marks
    glColor3f(1,1,1);
    gGraphics.DrawText2D(s + 35,440,12, "180");
    gGraphics.DrawText2D(s + 115,440,12, "270");
    gGraphics.DrawText2D(s + 195,440,12, "000");
    gGraphics.DrawText2D(s + 275,440,12, "090");
    gGraphics.DrawText2D(s + 355,440,12, "180");

    //heading rate indicator
    
    int a,b;
    int mult = static_cast< int > (200 / gFPerson.hdg_rate_max);
    if (gFPerson.hdg_rate < 0)
    {
      b = 399;
      a = b + static_cast<int> (mult * gFPerson.hdg_rate); 
      glColor3f(.7,0,0);
    }
    else if (gFPerson.hdg_rate > 0)
    {
      a = 399;
      b = a + static_cast< int > (mult * gFPerson.hdg_rate); 
      glColor3f(0,.7,0);
    }
    else
    {
      a = 397;
      b = 403;
      glColor3f(0,0,0);
    }
   
    glBegin(GL_QUADS);
      glVertex2f(a, 418);
      glVertex2f(a, 437);
      glVertex2f(b, 437);
      glVertex2f(b, 418);
    glEnd();
    
    glColor3f(0,.4, .4);  // black lines in heading rate quad for effect
    glBegin(GL_LINES);
      for (int i = 198; i < 602; i += 10)
      {
        glVertex2f(i, 417);
        glVertex2f(i, 439);
        glVertex2f(i + 1, 417);
        glVertex2f(i + 1, 439);
      }
      glColor3f(1,1,1);
      glVertex2f(399,410);
      glVertex2f(399,439);
    
    glEnd();
  
    //velocity indicator
    mult = static_cast< int > (2000 / gFPerson.vel_max);
    if (gFPerson.vel < 0)
    {
      a = static_cast<int> (-mult * gFPerson.vel); 
    }
    else if (gFPerson.vel > 0)
    {
      a = static_cast< int > (mult * gFPerson.vel); 
    }
    else
    {
      a = 0;
    }
  
    glBegin(GL_QUADS);
      if (gFPerson.vel > 0)
      {
        glColor3f(0,0,1);
        glVertex2f(622,438 - a);
        glVertex2f(622,438);
        glVertex2f(642,438);
        glVertex2f(642,438 - a);
      }
      else if (gFPerson.vel < 0)
      {
        glColor3f(1,0,0);
        glVertex2f(622,438);
        glVertex2f(622,438 + a);
        glVertex2f(642,438 + a);
        glVertex2f(642,438);
      }
    glEnd();
    
    glColor3f(0,.4, .4);  // black lines in vel quad for effect
    glBegin(GL_LINES);
    for (int i = 238; i < 520; i += 10)
    {
      glVertex2f(622,i);
      glVertex2f(642,i);
      glVertex2f(622,i + 1);
      glVertex2f(642,i + 1);
    }
    glColor3f(1,1,1);
    glVertex2f(620,439);
    glVertex2f(644,439);
    glEnd();
    
    
    //gun sights
    s += 40;
    glColor3f(1,1,0);
    glBegin(GL_LINES);
      glVertex2f(s + 170,293); //crosshair - vert line
      glVertex2f(s + 170,307);
      glVertex2f(s + 163,300); //crosshair - horiz line
      glVertex2f(s + 177,300);   
/*      glVertex2f(s + 160,290); //inner box
      glVertex2f(s + 180,290);
      glVertex2f(s + 180,290);
      glVertex2f(s + 180,310);
      glVertex2f(s + 180,310);
      glVertex2f(s + 160,310);
      glVertex2f(s + 160,310);
      glVertex2f(s + 160,290); */
    glEnd();

    //elevation
    glBegin(GL_LINES);
      glColor3f(0, .5, .5);
      glVertex2f(s - 30,119);
      glVertex2f(s - 30,479);
      for (int i=0;i< 19;i++)
      {
        glVertex2f(s - 35, 120 + i * 20);
        glVertex2f(s - 25, 120 + i * 20);        
      }
      glColor3f(1,1,1);
      glVertex2f(s - 40, 300);
      glVertex2f(s - 20, 300);        
    glEnd();
    
    //elevation moving marker
    glBegin(GL_QUADS);
      glColor3f(.8,.8,0);
      if (elev < 91)
      {
        glVertex2f(s - 35, 297 + (2 *  elev));
        glVertex2f(s - 35, 302 + (2 *  elev));
        glVertex2f(s - 25, 302 + (2 *  elev));
        glVertex2f(s - 25, 297 + (2 *  elev));      
      }
      else
      {
        elev = 360 - elev;
        glVertex2f(s - 35, 297 - (2 *  elev));
        glVertex2f(s - 35, 302 - (2 *  elev));
        glVertex2f(s - 25, 302 - (2 *  elev));
        glVertex2f(s - 25, 297 - (2 *  elev));      
      }

      
    glEnd();

    gGraphics.ResetProjection(GRAPHICS_ORTHO);
    glEnable(GL_DEPTH_TEST);
    glEnable(GL_TEXTURE_2D);
    glEnable(GL_CULL_FACE);
    glEnable(GL_LIGHTING);

  glPopMatrix();





}

/***********************************
  D R A W A S T E R O I D S
***********************************/
void DrawAsteroids(void)
{
  glPushMatrix();
    glColor3f(.4,.4,.4);
    srand(time(0));
    for(int i=0; i < 200;i++)
    {   
      glPushMatrix();
        glTranslatef( -10 + ((float)rand()/32768) * 20 ,
                         1.25 +  ((float)rand()/64768) ,
                      -10 + ((float)rand()/32768) * 20 );
        glutSolidSphere((float)rand()/32768 * 0.1,10,5);  
      glPopMatrix();
  }  
  glPopMatrix();
}

/***********************************
  D R A W A S P H E R E
***********************************/
void DrawSphere(void)
{
  glPushMatrix();
    glColor3f(1,0,0);
    glutSolidSphere(1,20,20);  
  glPopMatrix();
}

/***********************************
  D R A W C L O U D
***********************************/
void DrawCloud(void)
{
  glPushMatrix();
    glDisable(GL_LIGHTING);  
    glPushMatrix();
      glEnable(GL_BLEND);
      glBlendFunc(GL_SRC_ALPHA,GL_ONE);
      glColor4f(1,1,0, 0.8);
      glTranslatef(0,1.7,-2);
      glutSolidSphere(0.05,10,10);
      glTranslatef(.05,0,0);
      glutSolidSphere(0.05,10,10);
      glTranslatef(.05,0,0);
      glutSolidSphere(0.05,10,10);
      glDisable(GL_BLEND);
    glPopMatrix();  
    glEnable(GL_LIGHTING);
  glPopMatrix();
  
  glPushMatrix();
    glEnable(GL_BLEND);
    glBlendFunc(GL_SRC_ALPHA,GL_ONE);
    glColor4f(1,0,0,0.7);
    glTranslatef(0,1.7,-2);
    glutSolidSphere(0.06,10,10);
    glTranslatef(.05,0,0);
    glutSolidSphere(0.06,10,10);
    glTranslatef(.05,0,0);
    glutSolidSphere(0.06,10,10);
    glDisable(GL_BLEND);
  glPushMatrix();  
  
  //glPopMatrix();
    
    //glEnable(GL_BLEND);
    //glBlendFunc(GL_SRC_ALPHA,GL_ONE);

    /*
    for(int i=0; i < 25;i++)
    {   
      glPushMatrix();
        //glColor3f(0,(float)rand()/32768,(float)rand()/32768);
        //if(i<13)
        //  glColor4f(1.0, 0.7,0, 0.4);
        //else
        //  glColor4f(1.0, 0.2,0, 0.4);
        //glTranslatef( ((float)rand()* .000003) ,1.7 + ((float)rand() * .000003) * 0.5, -1 + ((float)rand()* .000003) );
        //glutSolidSphere((float)rand()* .0000005 ,10,10);  
      glPopMatrix();
    } 
    */ 
    
    

}
  
    //glColor3f(0,(float)rand()/32768,(float)rand()/32768);
    //glVertex3f( ((float)rand()/62768) ,1 + ((float)rand()/62768),-3 + ((float)rand()/62768)); 
    //glVertex3f( i * .001 ,1 + sin(.00314 * i),-3 + ((float)rand()/62768)); 
    //glVertex3f( i * .001 ,1 + sin(.00314 * i),-3 + ((float)rand() * 0.0000005));
    //glVertex3f( ((float)rand()/32768) * 25 ,2 + ((float)rand()/32768),-10 + ((float)rand()/32768)); 
    //creates nice arc
    //glVertex3f( i * .001 ,1 + sin(.00314 * i),-3 + ((float)rand()/62768)); 
  
