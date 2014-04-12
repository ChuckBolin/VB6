#include <cstdlib>
#include <iostream>
#include <string>

using namespace std;

//extern void printText2d(int x, int y, string text);

void printText2d(int x, int y, string text){
  
  for(int i = 0; i< text.length() -1;i++){
    cout <<  static_cast<unsigned char>( text[i]) << "   " 
         << (static_cast<int>( text[i]) > - 1 ?
             static_cast<int>( text[i]) :
             129 + (128 + static_cast<int>( text[i])))
         << endl; 
  }
  
  
}


int main(int argc, char *argv[])
{
  //create string to hold 256 values
  string s;
  s.resize(256);

  //load array with characters 0 to 255 
  char key[256];
  for(int i = 32;i< 255;i++){
    key[i]=i; 
    s[i] = key[i];
  }

  //printText2d(100,100,s);
  printText2d(100,100,"Good morning!");
  
  //char n = s[33];
  //cout << static_cast<int>( n);
  
  //print out string...one char at a time  
  /*
  string::const_iterator iterator1 = s.begin();
  int i =0;
  while(iterator1 != s.end()){
   // cout << i << ": " << *iterator1 << endl;
    ++iterator1;  
    ++i;
  }
  */
  /*
  int len = s.length();
  cout << "Length: " << len << endl;
  */
    system("PAUSE");
    return EXIT_SUCCESS;
}
