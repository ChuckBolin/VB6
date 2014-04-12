/****************************************************************************
 Project: 3d_game.dev  Exe: 3d_game.exe
 Filename: 3dstuff.h Author: Chuck Bolin cbolin@dycon.com
 Date: August 2002  URL: http://www.clg-net.com
****************************************************************************/
//void DrawQuad (float , float , float , int , float , float );
//void DrawBoxEmpty(float , float , float , float , float , float );
//void DrawBoxCover(float x, float y, float z, float l,  float w);
//void DrawCube (float x, float y, float z, float l, float w, float h);
//void DrawFrame (float x, float y, float z, int orient, float l, float w, float h);
//void LoadTextures(void);

//const float PI=3.1415927;
//const int CGO_ALIVE=1;  //Class Game Object (CGO) constants for GetStatus();
//const int CGO_ACTIVE=2;
//const int CGO_VISIBLE=3;
//const int NORTH = 1;
//const int EAST = 2;
//const int WEST = 3;
//const int SOUTH = 4;
//const int TOP = 5;
//const int BOTTOM = 6;

/***********************************************************************
 CLASS GameObject - Used for robots and other objects that may or may
 not move about manually (by the human) or autonomously (by AI).
 September 2002
***********************************************************************/
/*class CGameObject 
{
  public:
    float x,y,z; //current 3d coordinates
    float vt;    //velocity total - units/second
    float angle_x, angle_y, angle_z; //angular position about 3 axis (degrees)
    float lx,ly,lz; //lookat position
 
  
    //constructor - initialize 3d coordinates of object
    CGameObject(float posx=0,float posy=0, float posz=0)
    {
      x=posx;
      y=posy;
      z=posz;
      alive=true;
      //active=true;
      //visible=true;
    }
    ~CGameObject(){}

    //Member functions for CLASS GameObject
    void Update(float fps=100)
    {
      if ((alive==true) &&(active==true)) //must be alive & active to update
      {
        if (vt==0)
        {
           vx=0;
           vy=0;
           vz=0;
           mx=0;
           my=0;
           mz=0;
        }
        else
        {
          old_x=x; //stores moves in case they are required for backtracking
          old_y=y;
          old_z=z;
          float h = angle_y * PI/180; // h is heading in radians
          vx= vt * cos(h);  //eastern velocity
          vz = -vt * sin(h); //northern velocity
          mx = vx/fps; //units per second divided by fps
          mz = vz/fps;                  
          x+=mx;
          z+=mz;
        }
      }  
    }
    
    bool GetStatus(int flag)
    {
      switch(flag)
      {
        case CGO_ACTIVE:
          if(active==true) return (true);
          if(active==false) return (false);
          break;
        
        case CGO_ALIVE:
          if(alive==true) return (true);
          if(alive==false) return (false);
          break;
        
        case CGO_VISIBLE:
          if(visible==true) return (true);
          if(visible==false) return (false);
          break;
        default:
          return (false);
          break;
      }
    }

    void SetColorRGB(float r, float g, float b)
    {
      colorR=r;
      colorG=g;
      colorB=b;    
    }

    float GetRed ()
    {
      return colorR;
    }

    float GetGreen ()
    {
      return colorG;
    }

    float GetBlue ()
    {
      return colorB;
    }

    void Activate()
    {
      if (alive==true)active=true;  //verify alive
    }

    void Deactivate()
    {
      active=false;
    }

    void Destroy()
    {
      active=false;
      alive=false;
      visible=false;
    }

    void Show()
    {
      if (alive==true) visible=true; //verify alive
    }

    void Hide(void)
    {
      visible=false;
    }

  private:
    float vx,vy,vz; //velocities along 3 axis
    float old_x, old_y,old_z; //stores previous position

    bool active, alive, visible; //active indicates it moves, operates
                                 //alive means it is still operational
                                 //visible...is updated during screen draw
    float colorR, colorG, colorB; //RGB colors   
    float mx,my,mz; //                             
};
  
*/

//this is old struct for objects seen in game...robots use class above
struct gameObject
{
  float x; //current position
  float y;
  float z;
  float lx; //lookat position
  float ly;
  float lz;
  float old_x; //last position
  float old_y;
  float old_z;
  float angle_x; //angular position
  float angle_y; //heading is about y axis
  float angle_z;
  bool alive; //true if alive
  float mx;  //incremental step to move
  float my;
  float mz;
  float color[2];
  
} ;

