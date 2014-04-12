#include "drawing.h"


/***********************************************
 head = heading of head, ship = ships heading
***********************************************/
void DrawHeadingIndicator(float head, float ship)
{
    int s = 190;
  glPushMatrix();

    glDisable(GL_DEPTH_TEST);
    glDisable(GL_TEXTURE_2D);
    glDisable(GL_CULL_FACE);
    glDisable(GL_LIGHTING);
    gGraphics.SetProjection(GRAPHICS_ORTHO);
    glLoadIdentity();

    //black panel display
    glRectd(160,440,620,500);

    //icons
    
    glBegin(GL_LINES);
      //eyeball
      glColor3f(1,1,0);
      glVertex2f(s,460);
      glVertex2f(s + 5,455);
      glVertex2f(s + 5,455);
      glVertex2f(s + 15,455);  
      glVertex2f(s + 15,455);
      glVertex2f(s + 20,460);
      glVertex2f(s + 20,460);
      glVertex2f(s + 15,465);
      glVertex2f(s + 15,465);
      glVertex2f(s + 5,465);
      glVertex2f(s + 5,465);
      glVertex2f(s,460);
      glVertex2f(s + 8,455);
      glVertex2f(s + 8,462);
      glVertex2f(s + 8,462);
      glVertex2f(s + 12,462);
      glVertex2f(s + 12,462);
      glVertex2f(s + 12,455);    
      
      //spaceship
      glColor3f(1,1,1);
      glVertex2f(s,480);
      glVertex2f(s,495);
      glVertex2f(s,495);
      glVertex2f(s + 20,495);
      glVertex2f(s + 20,495);
      glVertex2f(s + 12,487);
      glVertex2f(s + 12,487);
      glVertex2f(s + 5,487);
      glVertex2f(s + 5,487);
      glVertex2f(s,480);
    glEnd();

    //draws markers
    glBegin(GL_LINES);
    
      //ship heading
      glColor3f(1,1,1);
      if (ship < 181)
      {
        glVertex2f(s + 210 + ( 160 * ship / 180 ), 478);
        glVertex2f(s + 210 + ( 160 * ship / 180), 497);
      }
     else 
      {
        glVertex2f(s + 50 + ( 160 * (ship - 180) / 180 ), 478);
        glVertex2f(s + 50 + ( 160 * (ship - 180) / 180 ), 497);
      }
 
      //head heading...look around
      glColor3f(1,1,0);
      if (head < 181)
      {
        glVertex2f(s + 210 + ( 160 * head / 180 ), 453);
        glVertex2f(s + 210 + ( 160 * head / 180), 472);
      }
     else 
      {
        glVertex2f(s + 50 + ( 160 * (head - 180) / 180 ), 453);
        glVertex2f(s + 50 + ( 160 * (head - 180) / 180 ), 472);
      }
    glEnd();

    //prints reference marks
    glColor3f(1,1,1);
    gGraphics.DrawText2D(s + 35,440,10, "180");
    gGraphics.DrawText2D(s + 115,440,10, "270");
    gGraphics.DrawText2D(s + 195,440,10, "000");
    gGraphics.DrawText2D(s + 275,440,10, "090");
    gGraphics.DrawText2D(s + 355,440,10, "180");


    //gun sights
    s += 40;
    glColor3f(1,1,0);
    glBegin(GL_LINES);
      glVertex2f(s + 170,293); //crosshair - vert line
      glVertex2f(s + 170,307);
      glVertex2f(s + 163,300); //crosshair - horiz line
      glVertex2f(s + 177,300);   
      glVertex2f(s + 160,290); //inner box
      glVertex2f(s + 180,290);
      glVertex2f(s + 180,290);
      glVertex2f(s + 180,310);
      glVertex2f(s + 180,310);
      glVertex2f(s + 160,310);
      glVertex2f(s + 160,310);
      glVertex2f(s + 160,290); 
    
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
  
