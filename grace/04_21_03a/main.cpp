#include <iostream>
#include <string>
#include <vector>
#include <sstream>
#include <iomanip>
#include "loaddata.h"
#include "eventlog.h"
#include "mymath.h"

using namespace std;

//stores all vertices

vector<VERTEX3D> gv;  

//main program
int main()
{
  gEventLog.SetFilename("log.txt");
  gEventLog.LogData("Commence processing main function",gEventLog.LOG_TEXT);
     
  //display data loaded into vector
  vector<VERTEX3D>::const_iterator pos;
  for(pos=gv.begin();pos!=gv.end();++pos)
  {
    cout<< pos->x <<" " << pos->y << " " << pos->z <<endl;
  }

  //dev-c++ stuff
  system("pause");
  exit(0);
  
}

