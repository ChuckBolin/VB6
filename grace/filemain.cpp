#include <stdio.h>
#include <cstdlib>
#include <iostream>
#include <fstream>
#include <string>

int main(int argc, char *argv[])
{
  ifstream in;
  string line;

  string filename = "script1.gam";
  in.open(filename.c_str(),ios::in);
  if(!in){
   exit(0); 
  }
  while (getline(in,line)){  //get one line at a time from script file
   
    cout << line << endl;    

  }
  
  
  /*
  //ifstream in("data.txt",ios::in);
  if(!in){
   exit(0); 
  }

  while (!in.eof()){  //get one line at a time from script file
    linecount++;
    in.getline(c, count);
    line = c;
    cout << linecount << ": " << line << endl;
    system ("pause");
    if (linecount>5) exit(0);
  }
  */
  system("pause");
  return 0;
}

