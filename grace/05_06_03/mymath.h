/*******************************************************************************
 MYMATH class
 Programmer: Chuck Bolin, 2003
 Purpose:  Allows for various math functions
*******************************************************************************/
#ifndef _ENGINE_MYMATH_H
#define _ENGINE_MYMATH_H
#define GLUT_DISABLE_ATEXIT_HACK

//include files
#include <cmath>
#include <string>
#include <strstream>
#include <vector>

//define math constants
#define TWO_PI         6.283185307f  //360 deg
#define THREE_PI_HALF  4.71238898    //270 deg
#define PI             3.141592654f  //180 deg
#define PI_HALF		     1.570796327f  //90 deg
#define PI_FOURTH	     0.785398163f  //45 deg

//structures
typedef struct _vertex2d
{
  float x,y;
} VERTEX2D;

struct VERTEX3D
{
  float x,y,z;
};

// Class Definition
class mymath
{
  public:
  
  //Angular math systems.  In the real world, measurements are usually in 
  //degrees. North is 0, East is 90, South is 180 and West is 270. Computers
  //on the other hand operate in radians and share the measurement system of
  //mathematics.  East is 0, North is 1.57, West is 3.14, South is 4.71.
  //The following functions convert from the various computer (comp) angles
  //to the real world (true) angles.

  //converts degrees to radians
  inline float DegToRad(float degrees)
  {
    float result = (degrees * PI / 180);
    if (result > TWO_PI) result = result - TWO_PI;
    if (result < 0) result = result + TWO_PI;
    return result;
  }

  //converts radians to degrees
  inline float RadToDeg(float radians)
  {
    float result = (radians * 180/PI);
    if (result > 360)  result  = result - 360;
    if (result < 0)  result = result + 360;
    return result;
  }

  //Converts computer radians to real world radians (see above for description)
  inline float CompRadToTrueRad(float comp_angle)
  {
    float result;
    result = TWO_PI + PI_HALF  - comp_angle;
    if (result > TWO_PI) result = result - TWO_PI;
    if (result < 0 ) result = result + TWO_PI;
    return result;
  }

  //same as above but uses rads to degrees
  inline float CompRadToTrueDeg(float comp_angle)
  {
    float result;
    result = TWO_PI + PI_HALF  - comp_angle;
    if (result > TWO_PI) result = result - TWO_PI;
    if (result < 0 ) result = result + TWO_PI;
    return RadToDeg(result);
  }

  //Converts real world degrees to computer radians
  inline float TrueDegToCompRad(float degrees)
  {
    float result;
    result = DegToRad(degrees);   //degrees converted to radians
    result = TWO_PI + PI_HALF  - result;  //conversion
    if (result > TWO_PI) result = result - TWO_PI; //check limits
    if (result < 0 ) result = result + TWO_PI;
    return result;
  }
   /*
  //converts a number to a STL string - NOTE: function overloading
  inline string IntToString(int number)  //integer
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }

  inline string NumberToString(long number)  //long
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }

  inline string NumberToString(float number)  //float
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }

  inline string NumberToString(double number)  //double
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }
  
  inline string NumberToString(unsigned short number)  //
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }

  inline string NumberToString(unsigned long number)  //
  {
    ostringstream os;
    os << number << ends;
    return os.str();
  }
  */
  
  inline float StringToFloat(string s){
    char buffer[10];
    float num;
    strstream str(buffer,10);
    str << s << ends;
    str >> num;
    return num;  
  }

  inline int StringToInteger(string s){
    char buffer[10];
    int num;
    strstream str(buffer,10);
    str << s << ends;
    str >> num;
    return num;  
  }


  inline float Hyp2D(float a, float b)
  {
    return sqrt( a * a + b * b );
  }

  inline float Distance2D(float x1,float y1, float x2,float y2)
  {
    return Hyp2D((x1-x2),(y1-y2));
  }

  inline float Hyp3D(float a, float b, float c)
  {
    return sqrt( a * a + b * b + c * c);
  }

  inline float Distance3D(float x1,float y1, float z1, float x2,float y2, float z2)
  {
    return Hyp3D((x1-x2),(y1-y2),(z1-z2));
  }
 
   inline float Distance3DVector(const VERTEX3D &pv1, const VERTEX3D &pv2)
  {
    return Hyp3D( (pv1.x - pv2.x), (pv1.y - pv2.y),(pv1.z - pv2.z)  );
  }


  //constributed by Afrohorse.com - 01.15.03     Thanks Afro
  float SqDistance3DVector(const VERTEX3D &v1, const VERTEX3D &v2)
  {
    // Return the Squared distance between two vectors
    float x,y,z;
    x = (v1.x - v2.x);
    y = (v1.y - v2.y);
    z = (v1.z - v2.z);
    return ( (x*x)+(y*y)+(z*z) ); 
  }

  //constructor
  mymath(){}
  
  //destructor
  ~mymath(){}   
 
  
};
#endif _ENGINE_MYMATH_H


/***********************************************************************
  Sample usage:
  
  vector<VERTEX3D> v1;
  vector<VERTEX3D>::const_iterator i;
  VERTEX3D temp,ref;
  
  //reference vertex;
  ref.x=0;
  ref.y=0;
  ref.z=0;
  
  //load vertices
  for (int j=0;j<20;j++){
    temp.x = j * .1;
    temp.y = j * .2;
    temp.z = j * .3;
    v1.push_back(temp);  
  }
  float fn =3.14;
  mymath cm;
  //display both distances
  for(i=v1.begin();i!=v1.end();i++)
  {
    cout << cm.Distance3DVector(*i,ref)<<" " 
         << cm.SqDistance3DVector(*i,ref) << endl;

  }
  

***********************************************************************/

