/*******************************************************************************
 SCRIPT.CPP
 Programmer: Chuck Bolin, 2003
 Purpose:  Reads and Interprets Script files
*******************************************************************************/

//Include files
#include "script.h"
//#include "myfunctions.h"

/****************************************************
 CheckStatus( )
 Verifies correct number of { },( ) and ;
****************************************************/
int CScript::CheckSyntax(string linein){
  int status = SYNTAX_OK; //okay
  int open_par = 0;       //number of open parenthesis
  int closed_par = 0;      //number of closed parenthesis
  int semicolon = 0; 

  int quote = 0;
  int comment = 0;
  bool allcomment = false;
  bool commentfound = false;
  string code;

  //consider only lines with something in them
  if (linein.length() >0 ) {
    int ret = linein.find_first_of("//");  //first comment position

    if (ret == -1){   //not found
      ret = linein.length();
      
    }
    else{             //need to know if comment is at beginning or in middle
      if (linein.substr(0,2)=="//"){
        allcomment=true;
        }
        
    }

    for(int i=0;i< ret;i++){  //check parenthesis
      if (linein.substr(i,1)=="(") 
        open_par += 1;
      if (linein.substr(i,1)==")")
        closed_par += 1;
      if (linein.substr(i,1)==";")
        semicolon += 1;
      if (linein.substr(i,1)=="{")
        left_brace += 1;
      if (linein.substr(i,1)=="}")
        right_brace += 1;
      if (linein.substr(i,1)=="\"")
        quote += 1;  
      if (linein.substr(i,2)=="//")
        quote += 1;
    } 

    //create error codes based upon these rules
    if ((left_brace < 1) && (right_brace < 1)&& (semicolon < 1 )&&
         (allcomment==false))
      status = SYNTAX_SEMICOLON;
    if (((left_brace > 0 )||(right_brace > 0)) && (status == SYNTAX_SEMICOLON))
      status = SYNTAX_OK;  
    if (open_par != closed_par ){
      if (open_par > closed_par)
        status = SYNTAX_CLOSED_PARENTHESIS;
      else
        status = SYNTAX_OPEN_PARENTHESIS;
    }  
    if ((quote ==1 ) || (quote == 3) || (quote == 5) || (quote == 7))
      status = SYNTAX_QUOTE;
  }
  
  if (status==SYNTAX_OK){
    //gScript.Push(linein);
    //save string to code vector
  }
  
  return status;     
}

/****************************************************
 CleanUp( )
removes spacing and converts all to lower case
****************************************************/
string CScript::CleanUp(string linein){
  string lineout;
  int it;
  
  //clean up string that have something in them
  if (linein.length()>0){
    int ret = linein.find_first_of("//");//find //
    switch(ret){
      case -1:
        ret = linein.length();
        break;
      case 0:
        ret = 0;
        break;
    }
  
    //create replacement string without spaces
    for(it = 0; it < ret;it++){
      if(linein.substr(it,1)!= " ")
        lineout += linein.substr(it,1);
    }
    lineout = lineout + linein.substr(ret);
   
    //convert string to lower case (doesn't affect comment
    char *buf = new char[lineout.length()]; //create temp buffer
    lineout.copy(buf,lineout.length());
    for (int i=0; i<lineout.length();i++){
      buf[i] = tolower(buf[i]);             //convert to lowercase 1 at a time
    }  
    string lineoutlower(buf,lineout.length());
    lineout = lineoutlower;
    delete buf;
  }

  return lineout;
}

/****************************************************
  DisplayResults( )
  Display script file information
****************************************************/
void CScript::DisplayResults(string scriptfile){

  ifstream in;
  ofstream out("temp.fil",ios::out);
  string line;
  string lineout;
  streamsize count;
  int linecount = 0;
  int num_errors = 0;
  int status = 0;
  char c[255];
  bool validformat = true;
  
  //in.open("script.gam");
  in.open(scriptfile.c_str());    //filename passed as command line argument
  if(!in){
   exit(0); 
  }

  cout << "*********************************" << endl;
  cout << "Grace Engine v0.1 Sript Processor" << endl;
  cout << "*********************************" << endl;
    
  while (getline(in, line)){  //get one line at a time from script file
    lineout= CleanUp(line);   //clean up line
    linecount += 1;
    cout << linecount << ": " << lineout << endl;  
    if( lineout.find(";") != string::npos) Push(lineout);
    status = CheckSyntax(lineout);
    if (status > SYNTAX_OK) num_errors += 1;

    switch (status){  //check for syntax errors
      case SYNTAX_OK:
        out << lineout << endl;
        break;
      case SYNTAX_PARENTHESIS:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Parenthesis...open or closed...doesn't add up!" <<
                 endl<< endl;
        break;
      case SYNTAX_OPEN_PARENTHESIS:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Open parenthesis is missing!" << endl<< endl;
        break;
      case SYNTAX_CLOSED_PARENTHESIS:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Closed parenthesis is missing!" << endl<< endl;
        break;
      case SYNTAX_SEMICOLON:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Missing semicolon!" << endl<< endl;
        break;
      case SYNTAX_UNKNOWN:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Unknown error!" << endl<< endl;
        break;
      case SYNTAX_QUOTE:
        cout << "        ^ ^ ^ " << endl;
        cout << "       Odd number of quotation marks.!" << endl<< endl;
        break;
      
    } 
  }
  
  //check braces
  if(left_brace != right_brace){
    num_errors++;
    cout << "        ^ ^ ^ " << endl;
    cout << "       Number of left and right braces are unequal!" << endl<< endl;
  }
  
  cout << endl;
  cout << "*********************************" << endl;
  cout << num_errors << " errors was/were detected!" << endl;
  cout << "*********************************" << endl;
  cout << endl;
  number_errors = num_errors;
  //return num_errors;
}

/****************************************************
 GetNumErrors( )
 returns number of errors during parsing
****************************************************/
int CScript::GetNumErrors(void){
  return number_errors;
}

/****************************************************
 Push( )
 Adds string to .script
****************************************************/
void CScript::Push(string mystring){
  script.push_back(mystring);
}

/****************************************************
 GetSize( )
 Returns the number of commands on script stack
****************************************************/
int CScript::GetSize(void){
  return script.size();
}

/****************************************************
 Get( )
 Retrieves string a specified position
****************************************************/
string CScript::Get(int i){
  return script[i];
}

/****************************************************
 GetLength( )
 Returns length of a command inside script stack
****************************************************/
int CScript::GetLength(int i){
  return script[i].length();
}

/******************************************************
 ExtractInfo( )
 Returns function name, args, and argv[]...max=10 args
******************************************************/
PARAM CScript::ExtractInfo(int i){
  PARAM par;
  int lparen, rparen;
  int args =0;
  mymath mm;
  
  string line = Get(i);
  lparen = line.find_first_of("(");
  rparen = line.find_first_of(")");
  
  //proceed only if parentheses exists
  if((lparen > 0) && (rparen > lparen)){
    par.name=line.substr(0,lparen);

    //count commas and calc number of arguments/parameters
    for(int i=0;i < line.length();i++){
      if (line.substr(i,1)==",")  args++;  
    }
    args++;
    par.args= args;
    args=0;
    
    //extract arguments
    int begin = lparen;
    for (int i=lparen;i < rparen; i++){
      if(line.substr(i,1)==","){
        par.argv[args] = line.substr(begin + 1,i-begin-1);
        args++;
        begin = i;  
      }
    }
    par.argv[args]= line.substr(begin + 1, rparen - begin - 1);
  }  
  return par;
}

/******************************************************
  ParseFile( )
  Extracts all info from script lines 
******************************************************/
void CScript::ParseFile(void){

  PARAM par;
 
  for (int i = 0;i < GetSize();i++){
    if (GetLength(i) > 0){

      //simple function line
      if( Get(i).find_first_of(";") != string::npos){
        par = ExtractInfo(i);
        parameters.push_back(par);
      }

      //line has left brace
      else if(gScript.Get(i).find_first_of("{") != string::npos){
      }

      //line has right brace
      else if(gScript.Get(i).find_first_of("}") != string::npos){
      }

      //comment located
      else if(gScript.Get(i).find_first_of("//") != string::npos){
      }
    }  
  }
  
  for (int i=0;i< parameters.size();i++){
    //cout << parameters[i].name << "  " 
    //     << parameters[i].args << "  ";
    for (int j=0;j< parameters[i].args;j++){
      //cout << parameters[i].argv[j] << "  ";
    }
    //cout << endl;         
  }
}

PARAM CScript::GetScriptLine(int i){
  return parameters[i];
}

/***********************************************************************
 Sample usage:

 //saving a string to gScript.script
 Push(lineout);

 //reading strings from gScript.script
  for (int i = 0;i < gScript.GetSize();i++){
    if (gScript.GetLength(i) > 0) cout << gScript.Get(i)<< endl;
  }

  OR....
  
  PARAM par;
  gScript.DisplayResults(scriptfile);  //checks for errors, loads .script
  gScript.ParseFile();
  for(int i=0;i< gScript.GetSize();i++){
    par = gScript.GetScriptLine(i);
    cout << par.name << " " << par.args << " " << endl; 
    for (int j=0;j< par.args;j++){
      cout << "  " <<  par.argv[j];
    }
    cout << endl << endl;
  }

***********************************************************************/

