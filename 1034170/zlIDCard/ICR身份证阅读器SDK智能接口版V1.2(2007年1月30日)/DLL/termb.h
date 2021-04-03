#ifndef _TERMB_H_
#define _TERMB_H_

int PASCAL InitComm(int port);
int PASCAL CloseComm();
int PASCAL Authenticate();
int PASCAL Read_Content(int active);
int PASCAL Read_Content_Path(char* cPath,int active);
BSTR PASCAL GetSAMID();

#endif