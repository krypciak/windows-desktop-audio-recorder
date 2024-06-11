// attribution: http://stackoverflow.com/a/15281070/32453
// you can compile this with gcc [mingw or mingw-w64] and probably other compilers as well

#include <stdlib.h> // NULL
#include <stdbool.h> // false
#include <windows.h> // AttachConsole, CTRL_C_EVENT, etc.
#include <stdio.h> // printf

#include <strsafe.h>
void logLastError() 
{ 
    LPTSTR errorText = NULL;
    
    FormatMessage(
       // use system message tables to retrieve error text
       FORMAT_MESSAGE_FROM_SYSTEM
       // allocate buffer on local heap for error text
       |FORMAT_MESSAGE_ALLOCATE_BUFFER
       // Important! will fail otherwise, since we're not 
       // (and CANNOT) pass insertion parameters
       |FORMAT_MESSAGE_IGNORE_INSERTS,  
       NULL,    // unused with FORMAT_MESSAGE_FROM_SYSTEM
       GetLastError(),
       MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
       (LPTSTR)&errorText,  // output 
       0, // minimum size for output buffer
       NULL);   // arguments - see note 
       
    if ( NULL != errorText )
    {
       printf("failure: %s", errorText);
       LocalFree(errorText);
       errorText = NULL;
    }
}

void sendCtrlC(int pid)
{
  //This does not require the console window to be visible.
  printf("sending ctrl+c to pid %d", pid); // if we want to log anything at all, have to log here or we'll be sending it to the other console with printf [!] XXX could maybe duplicate descriptors first or some odd...oh well.
  FreeConsole(); // free ourselves from our current console, in case we're being run from a cmd.exe otherwise can't attach to another (at most one).  And we probably are...if it fails because we don't have one, that's OK too.
  if (AttachConsole(pid))
  {
    // Disable Ctrl-C handling for our own program, so we don't "kill" ourselves on accident...though that might not hurt currently...
    SetConsoleCtrlHandler(NULL, true);
    GenerateConsoleCtrlEvent(CTRL_C_EVENT , 0); // these get sent to every process attached to this console, FWIW...
    // TODO could wait here for process to exit, hard kill it after a sufficient wait, etc. see http://stackoverflow.com/a/31020562/32453
    // could also send it *another* ctrl_c_event, for instance if you're trying to interrupt something in a batch file and it has that ugly "Batch file interrupted, continue [y/n]" [?]
  }
  else {
      logLastError(); // failure, PID might no longer exist, show a GUI error window
  }
}

int main( int argc, const char* argv[] ) {
    // assume they know the pid...
    if (argc != 2) {
        printf("syntax: pid");
        return 1;
    }
    int pid = atoi(argv[1]);
    sendCtrlC(pid);
}
