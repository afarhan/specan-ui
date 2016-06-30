#define _WIN32_IE 0x0900
#include <windows.h>
#include <commctrl.h>
#include <math.h>
#include "resource.h"
#include <stdio.h>

#define IDC_MAIN_STATUS 1000
#define IDC_FREQ 1001
#define IDC_PWR 1002
#define IDC_SPAN 1003
#define IDC_STEPS 1004
#define IDC_SWEEP 1005
#define IDC_1KHZ 1006
#define IDC_300KHZ 1007
#define IDC_MARK 1008
#define IDC_READING 1009
#define IDC_DIFF 1010
#define MARK_ME 100 //in the popup - menu

#define TEXT_HEIGHT (17)


/*  Make the main window class name into a global variable  */
char szClassName[ ] = "WindowsApp";
HWND mainWnd;
HINSTANCE instance;
int currentPort = 8;
int	currentFrequency = 0;
int	currentPower = -1000;

/* the status window has two parts .. read out, status*/
int statusWidths[] = {200, -1};
HWND statusWnd;
HWND toolWnd;
HANDLE serialPort = INVALID_HANDLE_VALUE;

/* these three are used for reading the serial port input from the Sweeperino,
   one byte at a time, until the line is complete
*/
#define BUFF_MAX 100
char inbuff[BUFF_MAX];
int buff_count=0;
int line_count = 0;

/* arrays to hold readings */
#define MAX_READINGS 100000
struct Reading {
  int frequency; /* in Hz */
  int power; /* in 1/10th of a db. value of '27' corresponds to 2.7db */
};

struct Reading readings[MAX_READINGS];
struct Reading refReadings[MAX_READINGS];
int nextReading=0;
int nrefReadings=0;

/* these are last selection from the sweep dialog */
int selectedSpan=9;
int selectedSteps=3;
int centerFreq = 30000000L;
int startFreq, endFreq;
int stepSize = 2000;
int	markFrequency = 0; //in Hz
int markPower = -1000; //in /10th of dbm

/* flag that tells you that sweeperino is busy doing a sweep */
int sweeperIsBusy = 0;
/* flag that triggers continous sweeps */
int continuous = 0;

/* ranges of sweeps offered by the sweep menu dialog */
struct Range{
    int width;
    int divisions;
    char *name;
};
#define RANGE_COUNT 11
struct Range spans[RANGE_COUNT] = {
  {1000, 200, "1 KHz"},
  {3000, 500, "3 KHz"},
  {10000, 2000, "10 KHz"},
  {40000, 5000, "30 KHz"},
  {100000, 20000, "100 KHz"},
  {400000, 50000, "300 kHz"},
  {1000000, 200000, "1 MHz"},
  {3000000, 500000, "3 MHz"},
  {10000000, 500000, "10MHz"},
  {30000000, 5000000, "30 MHz"},
  {50000000, 10000000, "50 Mhz"}
};

/* resolution of steps offered by the sweep menu dialog */
struct Step {
  int nsteps;
  char  *name;
};

#define STEPS_COUNT 5
struct Step steps[STEPS_COUNT] = {
  {10, "10 steps"},
  {30, "300 steps"},
  {100, "100 steps"},
  {300, "300 steps"},
  {600, "600 steps"}
};

/* GUI objects and dimensions */
HPEN penBlue, penBlack, penGray, penRed;
HFONT fontFixed, fontText;
HBRUSH background;
int	doRepaint = 0;
#define DISPLAY_WIDTH 600
#define DISPLAY_HEIGHT 500
#define X_OFF 20
#define Y_OFF 20

void logger(char *string){
  FILE *pf = fopen("sweeplog.txt", "a+t");
  if (!pf)
    return;
  fwrite(string, strlen(string), 1, pf);
  fclose(pf);
}

void setStatus(char *msg){
  SendMessage(statusWnd, SB_SETTEXT, 0, (LPARAM)msg);
}

void setStatus2(char *msg){
  SendMessage(statusWnd, SB_SETTEXT, 1, (LPARAM)msg);
}

void setReadOut(int power, int frequency){
  char readOut[100];
  
  sprintf(readOut, "Power : %d.%02dbm, Freq : %d.%03d.%03d", power/10, abs(power % 10), 
    frequency/1000000, (frequency % 1000000)/1000, frequency % 1000);
  SendMessage(statusWnd, SB_SETTEXT, 1, (LPARAM)readOut);  
}

/*  Declare Windows procedure  */
LRESULT CALLBACK WindowProcedure (HWND, UINT, WPARAM, LPARAM);
void plotReadings(HDC hdc);

int openSerialPort(int port){
  char c[200];
  DCB dcbSerialParams = {0};  
  COMMTIMEOUTS timeouts ={0};  
  
  sprintf(c, "\\\\.\\COM%d", port);
  serialPort = CreateFile(c, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  if (serialPort == INVALID_HANDLE_VALUE && GetLastError() != ERROR_SUCCESS){
    sprintf(c, "Can't open COM%d", port);
    setStatus(c);
    return -1;
  }
    
  /* we assume that we have successfully opened the port, we proceed to setting the port correctly */
  dcbSerialParams.DCBlength=sizeof(dcbSerialParams);
  if (!GetCommState(serialPort, &dcbSerialParams)) {
    CloseHandle(serialPort);
    sprintf(c, "Unable to read parameters for COM%d:", port);
    MessageBox(NULL, c, "Com port error", MB_OK);
    return -1;
  }
  
  dcbSerialParams.BaudRate=CBR_9600;
  dcbSerialParams.ByteSize=8;
  dcbSerialParams.StopBits=ONESTOPBIT;
  dcbSerialParams.Parity=NOPARITY;
  if(!SetCommState(serialPort, &dcbSerialParams)){
    sprintf(c, "Unable to set parameters for COM%d:", port);
    setStatus(c);    
    CloseHandle(serialPort);
    return -1;
  }

  timeouts.ReadIntervalTimeout=50;
  timeouts.ReadTotalTimeoutConstant=50;
  timeouts.ReadTotalTimeoutMultiplier=10;
  timeouts.WriteTotalTimeoutConstant=50;
  timeouts.WriteTotalTimeoutMultiplier=10;  

  if(!SetCommTimeouts(serialPort, &timeouts)){
    sprintf(c, "Unable to set timeouts for COM%d:", port);
    setStatus(c);    
    CloseHandle(serialPort);
    return -1;
  }

  sprintf(c, "COM%d initialised", port);
  setStatus(c);
  
  SetupComm(serialPort, 2000, 50);
  currentPort = port;
  return 0;
}

void closeSerialPort(){
	if (serialPort != INVALID_HANDLE_VALUE){
		CloseHandle(serialPort);
		serialPort = INVALID_HANDLE_VALUE;
	}
}

void saveCaliberation(){
  FILE *pf;
  int i;
    
  pf = fopen("sweeperino.caliberation", "wt");
  if (!pf)
    return;  
	fprintf(pf, "port:%d\n", currentPort);
	fclose(pf);
}

void loadCaliberation(){
    FILE *pf;
    char buff[100], key[20], c[100], *p, *q;
    int value;
        
    pf = fopen("sweeperino.caliberation", "rt");
    if (!pf)
      return;
    
    nrefReadings = 0;
    
    while (fgets(buff, sizeof(buff), pf)){

      p = strtok(buff, ":=");
 
      if (!p)
        continue;
        
      strcpy(key, p);
      p = strtok(NULL, " \n");

      value = atoi(p);

			if (!strcmp(key, "port")){
				closeSerialPort();
		    if (openSerialPort(value))
		      setStatus("Serial port error");
	    }
   }

	fclose(pf);
  InvalidateRect(mainWnd, NULL, TRUE);
  UpdateWindow(mainWnd);
}

int  serialWrite(char *c){
  DWORD len, i;
  
  len = strlen(c);
  if (!WriteFile(serialPort, c, len, &i, NULL))
    return 0;
  return 1;
}

/* each line of the readings comes as:
   "r[freq]:[power]\n"
  freq is in hz
  power is relative and each unit corresponds to 0.2db
  that is, a reading of 13 means 2.6db
*/
void enterReading(char *string){
  int freq, rfpower, i;
  char c[100], s[100];
  HDC dc;
  HANDLE oldpen;
  
  if (*string++ != 'r')
    return;
  
  for (i = 0; i < sizeof(c)-1; i++)
    if (isdigit(*string))
      c[i] = *string++;
  c[i] = 0;
  freq = atoi(c);
  
  if (*string != ':' || freq <= 0){
    logger("Bad format for reading\n");
    return;
  }
  string++; /* skip the colon */
  for (i = 0; i < sizeof(c)-1 && (isdigit(*string) || *string == '-') && *string; i++)
      c[i] = *string++;
  c[i] = 0;
	rfpower = atoi(c);
  
	readings[nextReading].frequency = freq;
  readings[nextReading].power = rfpower;
  sprintf(c, "%d, %d, %d, %d\n", nextReading, readings[nextReading].frequency , readings[nextReading].power, atoi(c));  
  logger(c);
  sprintf(c, "Reading %d/%d", nextReading, (endFreq-startFreq)/stepSize);
  setStatus(c);
  nextReading++;
  
	InvalidateRect(mainWnd, NULL, FALSE);
	UpdateWindow(mainWnd);
}

void serialReceived (){
  char c[BUFF_MAX], string[BUFF_MAX];
  
  sprintf(c, "Rx [%s]\n", inbuff);
  logger(c);
 
  switch(*inbuff){
    case 'k':
      break;
    case 'b':
      nextReading = 0;
      sweeperIsBusy = 1;
      setStatus("Busy...");
      break;
    case 'e':
      sweeperIsBusy = 0;
      setStatus("Ready");
      InvalidateRect(mainWnd, NULL, TRUE);
      UpdateWindow(mainWnd);      
      break;
    case 'r':
      enterReading(inbuff);
			break;
  }
}

void setupSweep(){
  char c[100];
  int f1, f2;

	closeSerialPort();
	openSerialPort(currentPort);
	
  endFreq = centerFreq + spans[selectedSpan].width;
  startFreq = centerFreq - spans[selectedSpan].width;
  
	f2 = endFreq;
	f1 = startFreq;
	stepSize = (f2 - f1) / steps[selectedSteps].nsteps;
  
  sprintf(c, "f%d\n", f1);
  serialWrite(c);
  Sleep(100);
  sprintf(c, "t%d\n", f2);
  serialWrite(c);
  Sleep(100);
  sprintf(c, "s%d\n", stepSize);
  serialWrite(c);
  Sleep(100);
	
	if (IsDlgButtonChecked(mainWnd, IDC_1KHZ))
		serialWrite("n\n");
	else
		serialWrite("w\n");
	
}

void startSweep(){
  if (sweeperIsBusy)
    return;
  if (serialWrite("g\n")){
    nextReading = 0;
    sweeperIsBusy = 1;
  }
  SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
}

void serialPoll(){
  char c;
  DWORD n = 0;
  char string[100];

  if (serialPort == NULL)
    return;
    
  while (ReadFile(serialPort, &c, 1, &n, NULL) > 0){
    if (n != 1)
      break;
        
    if (c == '\n'){
      inbuff[buff_count] = 0;
      serialReceived();
      buff_count = 0;
    }
    else if (buff_count < BUFF_MAX-1){
        inbuff[buff_count++] = c;
    }
  } /* end of the loop */
}

/* get a save as file name */
int getFilename(char *path)
{
  OPENFILENAME ofn;

  ZeroMemory(&ofn, sizeof(ofn));
  *path = 0;

  ofn.lStructSize = sizeof(ofn); /* SEE NOTE BELOW */
  ofn.hwndOwner = mainWnd;
  ofn.lpstrFilter = "Bitmap Files (*.bmp)\0*.bmp\0";
  ofn.lpstrFile = path;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_EXPLORER | OFN_SHOWHELP | OFN_OVERWRITEPROMPT;
  ofn.lpstrDefExt = "bmp";

  if(GetSaveFileName(&ofn))
    return 1;
  else
    return 0;
}
/* to save the screen as a bmp */
int captureDisplay(HWND hWnd, char *path)
{
    HDC hdcWindow;
    HDC hdcMemDC = NULL;
    HBITMAP hbmScreen = NULL;
    BITMAP bmpScreen;

    /* Retrieve the handle to a display device context for the client 
     area of the window. */
    hdcWindow = GetDC(hWnd);
    hdcMemDC = CreateCompatibleDC(hdcWindow); 

    if(!hdcMemDC)
    {
        MessageBox(hWnd, "CreateCompatibleDC has failed","Failed", MB_OK);
        goto done;
    }

    /* Get the client area for size calculation */
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    /* This is the best stretch mode */
    SetStretchBltMode(hdcWindow,HALFTONE);
    
    /* Create a compatible bitmap from the Window DC */
    hbmScreen = CreateCompatibleBitmap(hdcWindow, rcClient.right-rcClient.left, rcClient.bottom-rcClient.top);
    
    if(!hbmScreen)
    {
        MessageBox(hWnd, "CreateCompatibleBitmap Failed", "Failed", MB_OK);
        goto done;
    }

    /* Select the compatible bitmap into the compatible memory DC. */
    SelectObject(hdcMemDC,hbmScreen);
    
    /* Bit block transfer into our compatible memory DC. */
    if(!BitBlt(hdcMemDC, 
               0,0, 
               rcClient.right-rcClient.left, rcClient.bottom-rcClient.top, 
               hdcWindow, 
               0,0,
               SRCCOPY))
    {
        MessageBox(hWnd, "BitBlt has failed", "Failed", MB_OK);
        goto done;
    }

    /* Get the BITMAP from the HBITMAP */
    GetObject(hbmScreen,sizeof(BITMAP),&bmpScreen);
     
    BITMAPFILEHEADER   bmfHeader;    
    BITMAPINFOHEADER   bi;
     
    bi.biSize = sizeof(BITMAPINFOHEADER);    
    bi.biWidth = bmpScreen.bmWidth;    
    bi.biHeight = bmpScreen.bmHeight;  
    bi.biPlanes = 1;    
    bi.biBitCount = 32;    
    bi.biCompression = BI_RGB;    
    bi.biSizeImage = 0;  
    bi.biXPelsPerMeter = 0;    
    bi.biYPelsPerMeter = 0;    
    bi.biClrUsed = 0;    
    bi.biClrImportant = 0;

    DWORD dwBmpSize = ((bmpScreen.bmWidth * bi.biBitCount + 31) / 32) * 4 * bmpScreen.bmHeight;

    /* Starting with 32-bit Windows, GlobalAlloc and LocalAlloc are implemented as wrapper functions that 
       call HeapAlloc using a handle to the process's default heap. Therefore, GlobalAlloc and LocalAlloc 
       have greater overhead than HeapAlloc. */
    HANDLE hDIB = GlobalAlloc(GHND,dwBmpSize); 
    char *lpbitmap = (char *)GlobalLock(hDIB);    

    /* Gets the "bits" from the bitmap and copies them into a buffer 
       which is pointed to by lpbitmap. */
    GetDIBits(hdcWindow, hbmScreen, 0,
        (UINT)bmpScreen.bmHeight,
        lpbitmap,
        (BITMAPINFO *)&bi, DIB_RGB_COLORS);

    /* A file is created, this is where we will save the screen capture. */
    HANDLE hFile = CreateFile(path,
        GENERIC_WRITE,
        0,
        NULL,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL, NULL);   
    
    /* Add the size of the headers to the size of the bitmap to get the total file size */
    DWORD dwSizeofDIB = dwBmpSize + sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
 
    /* Offset to where the actual bitmap bits start. */
    bmfHeader.bfOffBits = (DWORD)sizeof(BITMAPFILEHEADER) + (DWORD)sizeof(BITMAPINFOHEADER); 
    
    /* Size of the file */
    bmfHeader.bfSize = dwSizeofDIB; 
    
    /* bfType must always be BM for Bitmaps */
    bmfHeader.bfType = 0x4D42; /* BM identifier */  
 
    DWORD dwBytesWritten = 0;
    WriteFile(hFile, (LPSTR)&bmfHeader, sizeof(BITMAPFILEHEADER), &dwBytesWritten, NULL);
    WriteFile(hFile, (LPSTR)&bi, sizeof(BITMAPINFOHEADER), &dwBytesWritten, NULL);
    WriteFile(hFile, (LPSTR)lpbitmap, dwBmpSize, &dwBytesWritten, NULL);
    
    /* Unlock and Free the DIB from the heap */
    GlobalUnlock(hDIB);    
    GlobalFree(hDIB);

    /* Close the handle for the file that was created */
    CloseHandle(hFile);
       
    /* Clean up */
done:
    DeleteObject(hbmScreen);
    DeleteObject(hdcMemDC);
    ReleaseDC(hWnd,hdcWindow);

    return 0;
}

void onSaveAs(){
  char path[MAX_PATH];
  
  if (getFilename(path))
    captureDisplay(mainWnd, path);
}

BOOL CALLBACK dlgSweep(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam) {
    char c[10];
    int i, center, newspan, newcenter, newsteps, f1, f2;
    
  if (Msg == WM_INITDIALOG){
    
      SendDlgItemMessage(hWnd, IDC_CENTER, EM_LIMITTEXT, 6, 0); 
     
      /* populate the list boxes */
      for (i = 0; i < RANGE_COUNT; i++)
        SendDlgItemMessage(hWnd, IDC_RANGE, CB_ADDSTRING, 0, (LPARAM)(spans[i].name)); 
        
      for (i = 0; i < STEPS_COUNT; i++)     
        SendDlgItemMessage(hWnd, IDC_NSTEPS, CB_ADDSTRING, 0, (LPARAM)(steps[i].name)); 
      
      /* set the current settings */
      SendDlgItemMessage(hWnd, IDC_RANGE, CB_SETCURSEL, (WPARAM)selectedSpan, 0);
      SendDlgItemMessage(hWnd, IDC_NSTEPS, CB_SETCURSEL, (WPARAM)selectedSteps, 0);
      sprintf(c, "%d", centerFreq/1000);
      SendDlgItemMessage(hWnd, IDC_CENTER, WM_SETTEXT, 0, (LPARAM)c);
      
      /* stop the power meter polling */
      KillTimer(mainWnd, 2);
  }
    
  if (Msg == WM_COMMAND) {
    switch(LOWORD(wParam)){
      case IDOK:
        newspan = SendDlgItemMessage(hWnd, IDC_RANGE, CB_GETCURSEL, 0, 0);
        if (newspan < 0 || newspan >= RANGE_COUNT){
          MessageBox(hWnd, "Choose a range to be swept", "Range Not Selected", MB_OK);          
          return 0;
        }
        SendDlgItemMessage(hWnd, IDC_CENTER, WM_GETTEXT, 7, (LPARAM)c);
        newcenter = atoi(c);
        if (newcenter > 200000){
          MessageBox(hWnd, "The center frequency should be below 200,000KHz (200 MHz)", "Invalid Frequency", MB_OK);
          return;
        }
        f1 = (newcenter * 1000) - spans[newspan].width;
        f2 = (newcenter * 1000) + spans[newspan].width;
        if (f1 < 0){
          MessageBox(hWnd, "The range is too wide for the central frequency\r\nEither reduce the centeral freq or the range", 
          "Wrong Range", MB_OK);
          return;          
        }
        centerFreq = newcenter * 1000;
        selectedSpan = newspan;
        selectedSteps = SendDlgItemMessage(hWnd, IDC_NSTEPS, CB_GETCURSEL, 0, 0);
        setupSweep();
        startSweep(); 
        InvalidateRect(mainWnd, NULL, TRUE);
        UpdateWindow(mainWnd);
      case IDCANCEL:
        EndDialog(hWnd, wParam);
        return 0;
    }
  }
  else
    return FALSE;

}

BOOL CALLBACK dlgPortSetting(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam) {
    char c[10];
    int i;
    
    /* fill the ports list and set it to the current selection */
    if (Msg == WM_INITDIALOG){
      for (i = 1; i < 20; i++){
        sprintf(c, "COM%d", i);
        SendDlgItemMessage(hWnd, IDC_PORT, CB_ADDSTRING, 0, (LPARAM)c);
      }
      if (currentPort != -1)
        SendDlgItemMessage(hWnd, IDC_PORT, CB_SETCURSEL, (WPARAM)currentPort, 0L);
    }
    if (Msg == WM_COMMAND) {
      switch(LOWORD(wParam)){
        case IDOK:
          i = SendDlgItemMessage(hWnd, IDC_PORT, CB_GETCURSEL, 0, 0);
          /* the list starts with zero index, and the com ports start with com1, so we add 1 to the index */
          CloseHandle(serialPort);
          if(!openSerialPort(i+1)){ /*if the port opened, then save the caliberation file */
						currentPort = i + 1;
				    saveCaliberation();
		      }
		      else {
            MessageBox(hWnd, "Port is not available. Check the serial cable!", "Serial Port Error", MB_OK);
          }
          EndDialog(hWnd, wParam);
          break;
        case IDCANCEL:
          EndDialog(hWnd, wParam);
          return 0;
      }
    }
    else
      return FALSE;

}

int freqToScreenx(int freq){
  int f1, f2, x;
  
  /* f1 = centerFreq - spans[selectedSpan].width;
  f2 = centerFreq + spans[selectedSpan].width; */
  return MulDiv(freq - startFreq, DISPLAY_WIDTH, endFreq - startFreq) + X_OFF;
}

int screenToFreq(int x){
  int f1, f2;
  
  x = x - X_OFF;
/*  f1 = centerFreq - spans[selectedSpan].width;
  f2 = centerFreq + spans[selectedSpan].width; */
  return MulDiv(x, endFreq - startFreq, DISPLAY_WIDTH) + startFreq;
}

/* in 1/10th of db, starting from -100dbm to 0 dbm */
int screenToPower(int y){
  /* y = DISPLAY_HEIGHT + Y_OFF - y; /* invert the y axis */
  //return  (1000 * y)/DISPLAY_HEIGHT -1000; *
	return ((Y_OFF-y) * 1000)/DISPLAY_HEIGHT;
	
}

int powerToScreen(int power){
  /* return DISPLAY_HEIGHT + Y_OFF - ((power +1000)* DISPLAY_HEIGHT)/1000; */
	return Y_OFF - ((power * DISPLAY_HEIGHT)/1000);
}

int frequencyToPower(int freq){
	int power = -1000; //in 1/10th of a dbm
	int match = 1000000000;
	int i;
  for (i = 0; i < nextReading; i++){
		if (abs(readings[i].frequency - freq) < match){
			match = abs(readings[i].frequency - freq);
			power = readings[i].power;
		}
  }
	return power;
}

/* this searches the caliberation readings and gets an approximate reference for the reading */
int getReference(int frequency){
  char buff[100];
  int i, bestmatch = 0;
  for (i = 0; i < MAX_READINGS; i++)
    if (abs(frequency - refReadings[i].frequency) <= abs(frequency - refReadings[bestmatch].frequency))
        bestmatch = i;
        
/* sprintf(buff, "ref @ %d is %d at inx %d", frequency, refReadings[bestmatch].frequency, bestmatch);
 setStatus(buff);*/
  return refReadings[bestmatch].power;
}

void plotReadings(HDC hdc){
  int i=0;
  HANDLE oldpen;
  
  oldpen = SelectObject(hdc, penBlack);  
  MoveToEx(hdc, freqToScreenx(readings[i].frequency), powerToScreen(readings[i].power), NULL);
  for (i = 0; i < nextReading; i++){
    LineTo(hdc, freqToScreenx(readings[i].frequency), powerToScreen(readings[i].power));
  }
  SelectObject(hdc, oldpen);
}

void plotReference(HDC hdc){
  int i=0;
  HANDLE oldpen;
  
  oldpen = SelectObject(hdc, penRed);  
  MoveToEx(hdc, freqToScreenx(readings[i].frequency), powerToScreen(getReference(readings[i].frequency)), NULL);
  for (i = 0; i < nextReading-1; i++){
    LineTo(hdc, freqToScreenx(readings[i].frequency), powerToScreen(getReference(readings[i].frequency)));
  }
  SelectObject(hdc, oldpen);
}

void loadControls(HWND hwnd){
	char c[100];
  sprintf(c, "%d", centerFreq/1000);
  SendDlgItemMessage(hwnd, IDC_FREQ, WM_SETTEXT, 0, (LPARAM)c);
  SendDlgItemMessage(hwnd, IDC_SPAN, CB_SETCURSEL, (WPARAM)selectedSpan, 0);
  SendDlgItemMessage(hwnd, IDC_STEPS, CB_SETCURSEL, (WPARAM)selectedSteps, 0);
}

void onSweep(HWND hWnd){
	int newspan, newcenter, f1, f2;
	char c[100];
	
  newspan = SendDlgItemMessage(hWnd, IDC_SPAN, CB_GETCURSEL, 0, 0);
  if (newspan < 0 || newspan >= RANGE_COUNT){
		MessageBox(hWnd, "Choose a range to be swept", "Range Not Selected", MB_OK);          
    return;
  }
  
	SendDlgItemMessage(hWnd, IDC_FREQ, WM_GETTEXT, 7, (LPARAM)c);
  newcenter = atoi(c);
  if (newcenter > 200000){
    MessageBox(hWnd, "The center frequency should be below 200,000KHz (200 MHz)", "Invalid Frequency", MB_OK);
    return;
  }
  f1 = (newcenter * 1000) - spans[newspan].width;
  f2 = (newcenter * 1000) + spans[newspan].width;
  if (f1 < 0){
		MessageBox(hWnd, "The range is too wide for the central frequency\r\nEither reduce the centeral freq or the range", 
      "Wrong Range", MB_OK);
    return;          
  }
  centerFreq = newcenter * 1000;
  selectedSpan = newspan;
  selectedSteps = SendDlgItemMessage(hWnd, IDC_STEPS, CB_GETCURSEL, 0, 0);
  setupSweep();
  startSweep(); 
  InvalidateRect(hWnd, NULL, TRUE);
  UpdateWindow(hWnd);
}
void onPaint(HWND wnd){
  PAINTSTRUCT ps;
  HDC hdc;
  int markers, f1, f2, f, i, x, y;
  char c[100];
  HANDLE oldpen;
  
  hdc = BeginPaint(wnd, &ps);
  SelectObject(hdc, GetStockObject(BLACK_PEN));
	SelectObject(hdc, fontText);
	SetBkMode(hdc, TRANSPARENT);	
	
  SelectObject(hdc, fontFixed);
  /* draw the grid,
  each division represents a frequency */
	
  SelectObject(hdc, GetStockObject(WHITE_BRUSH));
  Rectangle(hdc, X_OFF, Y_OFF, DISPLAY_WIDTH + X_OFF, DISPLAY_HEIGHT + Y_OFF);
  f2 = centerFreq + spans[selectedSpan].width;
  f1 = centerFreq - spans[selectedSpan].width;
  
  for (f = centerFreq; f < f2; f = f+spans[selectedSpan].divisions){
    x = freqToScreenx(f);
    SelectObject(hdc, penBlue);  
    MoveToEx(hdc, x, Y_OFF, NULL);
    LineTo(hdc, x, DISPLAY_HEIGHT + Y_OFF);
  }
  sprintf(c, "%gMhz", (double)(f2/1000000.0));
  TextOut(hdc, freqToScreenx(f2)-10, DISPLAY_HEIGHT+Y_OFF+4, c, strlen(c));

  sprintf(c, "%gMHz", (double)(centerFreq/1000000.0));
  TextOut(hdc, freqToScreenx(centerFreq)-10, DISPLAY_HEIGHT+Y_OFF+4, c, strlen(c));

  sprintf(c, "%gMhz", (double)(f1/1000000.0));
  TextOut(hdc, X_OFF-10, DISPLAY_HEIGHT+Y_OFF+4, c, strlen(c));
  
  for (f = centerFreq; f1 < f; f = f-spans[selectedSpan].divisions){
    x = freqToScreenx(f);
    MoveToEx(hdc, x, 20, NULL);
    LineTo(hdc, x, DISPLAY_HEIGHT + Y_OFF);
  }
  
  /* plot from -100dbm to 0dbm */
  for (i = -1000; i <= 0; i += 100){
    y = powerToScreen(i);
    SelectObject(hdc, penBlue);  
    MoveToEx(hdc, X_OFF, y, NULL);
    LineTo(hdc, DISPLAY_WIDTH + X_OFF, y);
    sprintf(c, "%d dbm", i/10);
    TextOut(hdc, DISPLAY_WIDTH + X_OFF + 5, y-7, c, strlen(c));        
  }  

  /* plot the reference level */
 /*  plotReference(hdc); */
  plotReadings(hdc); 
  EndPaint(wnd, &ps);
}

void setMark(int x, int y){
  char c[100];
  
  markFrequency = currentFrequency;
	markPower = currentPower;
  sprintf(c, "MARK: %d.%03d.%03d %d.%ddbm", 
			markFrequency/1000000L, (markFrequency % 1000000L)/1000, markFrequency % 1000, 
			markPower/10, abs(markPower%10));
	SendDlgItemMessage(mainWnd, IDC_MARK, WM_SETTEXT, 0, (LPARAM)c);

}

void onMouseMove(int x, int y){
  char c[100];
	int diffPower, diffFrequency;
  
  currentFrequency = screenToFreq(x);
	currentPower = frequencyToPower(currentFrequency);
  sprintf(c, "Freq: %d.%03d.%03d Hz Power:%d.%d db", 
			currentFrequency/1000000L, (currentFrequency % 1000000L)/1000, currentFrequency % 1000, 
			currentPower/10, abs(currentPower%10));
  setStatus2(c);
  sprintf(c, "MHz:%d.%03d.%03d  Power:%d.%ddbm", 
			currentFrequency/1000000L, (currentFrequency % 1000000L)/1000, currentFrequency % 1000, 
			currentPower/10, abs(currentPower%10));
	SendDlgItemMessage(mainWnd, IDC_READING, WM_SETTEXT, 0, (LPARAM)c);
	
	diffPower = markPower - currentPower;
	diffFrequency = markFrequency - currentFrequency;
  sprintf(c, "DIFF: %d.%03d.%03d %d.%ddbm", 
			diffFrequency/1000000L, abs(diffFrequency % 1000000L)/1000, abs(diffFrequency % 1000), 
			diffPower/10, abs(diffPower%10));
	SendDlgItemMessage(mainWnd, IDC_DIFF, WM_SETTEXT, 0, (LPARAM)c);	
}


void setupControls(HWND hwnd){
	int i, y;
	HWND w;
	char c[100];
	
	y = Y_OFF;
	
	w = CreateWindow("static", "READING:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 300, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_READING,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		

	y += (TEXT_HEIGHT * 3)/2;
	
	w = CreateWindow("static", "MARK:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 200, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_MARK,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		
	
	y += (TEXT_HEIGHT * 3)/2;
	w = CreateWindow("static", "DIFF:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 200, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_DIFF,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		

	y += (TEXT_HEIGHT * 2);

	CreateWindow("BUTTON","SWEEP",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 100, 30, hwnd,
		(HMENU)IDC_SWEEP,
		GetModuleHandle(NULL),
		NULL);
	
	y += 40;
	
	w = CreateWindow("static", "Center Freq(KHz):",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 150, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_STATIC,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);
	y+= TEXT_HEIGHT;
	
	CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT","",
		WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL|ES_NUMBER,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, 4 + y, 100, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_FREQ,
		GetModuleHandle(NULL),
		NULL);
	SendDlgItemMessage(hwnd, IDC_FREQ, WM_SETFONT, (WPARAM)fontText, 0);
	SendDlgItemMessage(hwnd, IDC_FREQ, EM_SETLIMITTEXT, (WPARAM)5, 0);

  sprintf(c, "%d", centerFreq/1000);
  SendDlgItemMessage(hwnd, IDC_FREQ, WM_SETTEXT, 0, (LPARAM)c);

	
	y += TEXT_HEIGHT * 2;
	
	w = CreateWindow("static", "Freq Span:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 150, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_STATIC,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);
	
	y += TEXT_HEIGHT;
	
	CreateWindow(WC_COMBOBOX,"",
		WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|CBS_HASSTRINGS|CBS_DISABLENOSCROLL|WS_VSCROLL,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 100, TEXT_HEIGHT * 8, hwnd,
		(HMENU)IDC_SPAN,
		GetModuleHandle(NULL),
		NULL);	
	SendDlgItemMessage(hwnd, IDC_SPAN, WM_SETFONT, (WPARAM)fontText, 0);	
	for (i = 0; i < RANGE_COUNT; i++)
			SendDlgItemMessage(hwnd, IDC_SPAN, CB_INSERTSTRING, i, (LPARAM)spans[i].name);
			
	y += TEXT_HEIGHT * 2;		
	w = CreateWindow("static", "Steps:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 150, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_STATIC,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		
	
	y += TEXT_HEIGHT;
	
	w = CreateWindow(WC_COMBOBOX,"",
		WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|CBS_HASSTRINGS|CBS_DISABLENOSCROLL|WS_VSCROLL,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, 4 + y, 100, TEXT_HEIGHT * 8, hwnd,
		(HMENU)IDC_STEPS,
		GetModuleHandle(NULL),
		NULL);	
		
	SendDlgItemMessage(hwnd, IDC_STEPS, WM_SETFONT, (WPARAM)fontText, 0);	
	for (i = 0; i < STEPS_COUNT; i++)
			SendDlgItemMessage(hwnd, IDC_STEPS, CB_INSERTSTRING, i, (LPARAM)steps[i].name);
	
	y += TEXT_HEIGHT * 3;	
	
	w = CreateWindow("static", "Resolution Bandwidth:",
		WS_CHILD|WS_VISIBLE,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 200, (TEXT_HEIGHT * 3)/2, hwnd,
		(HMENU)IDC_STATIC,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		
	y += (TEXT_HEIGHT * 3)/2;
	
	w = CreateWindow("BUTTON","300 KHz",
		WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON|WS_GROUP,
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 200, TEXT_HEIGHT, hwnd,
		(HMENU)IDC_300KHZ,
		GetModuleHandle(NULL),
		NULL);		
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		
	
	y += (TEXT_HEIGHT * 3)/2;
	
	w = CreateWindow("BUTTON","1 KHz RBW",
		WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, 
		(X_OFF * 2) + DISPLAY_WIDTH + 60, y, 200, TEXT_HEIGHT, hwnd,
		(HMENU)IDC_1KHZ,
		GetModuleHandle(NULL),
		NULL);
	SendMessage(w, WM_SETFONT, (WPARAM)fontText, TRUE);		


  SendDlgItemMessage(hwnd, IDC_SPAN, CB_SETCURSEL, (WPARAM)selectedSpan, 0);
  SendDlgItemMessage(hwnd, IDC_STEPS, CB_SETCURSEL, (WPARAM)selectedSteps, 0);
  sprintf(c, "%d", centerFreq/1000);
  SendDlgItemMessage(hwnd, IDC_FREQ, WM_SETTEXT, 0, (LPARAM)c);
	CheckDlgButton(hwnd, IDC_300KHZ, BST_CHECKED);
}

/*  This function is called by the Windows function DispatchMessage()  */

LRESULT CALLBACK WindowProcedure (HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)                  /* handle the messages */
    {
        case WM_CREATE:
          setStatus("Initializing...");
          SetTimer(hwnd, 1, 10, NULL); /* timer 1 polls for more data from sweeperino's serial port */
          SetTimer(hwnd, 2, 200, NULL); /* timer 2 polls for current power meter reading from the sweeperino */
          logger("Started log\n");
					setupControls(hwnd);
          break;
				case WM_CTLCOLORSTATIC:{
					HDC hdc;
					hdc = (HDC) wParam; 
					SetTextColor(hdc, RGB(0,0,0));    
					SetBkMode (hdc, TRANSPARENT);
					return (LRESULT)background;
					}
					break;
        case WM_MOUSEMOVE:
          onMouseMove(LOWORD(lParam), HIWORD(lParam));
          break;
        case WM_TIMER:
          if (wParam == 1)
            serialPoll();
          break;
        case WM_COMMAND:
          switch(LOWORD(wParam)){
            case IDM_SET_PORT:
              DialogBox(instance, MAKEINTRESOURCE(IDD_SET_PORT), hwnd, (DLGPROC)dlgPortSetting);
              break;              
            case IDM_SWEEP:
              if (sweeperIsBusy)
                MessageBox(hwnd, "Sweeper is busy. Try later", "Sweeper Busy", MB_OK);
              else if (serialPort == INVALID_HANDLE_VALUE)
                MessageBox(hwnd, "Connect to Sweeperino from Action>Connect.. menu", "Sweeperino disconnected", MB_OK);
              else
                DialogBox(instance, MAKEINTRESOURCE(IDD_SWEEP), hwnd, (DLGPROC)dlgSweep);
              break;
							
            case IDM_QUIT:
              DestroyWindow(hwnd);
              break;
            case IDM_SAVE_AS:
              onSaveAs();
              break;
						case IDC_SWEEP:
							if (HIWORD(wParam) == BN_CLICKED)
								onSweep(hwnd);
							break;
          }
          break;
        case WM_PAINT:
          onPaint(hwnd);
          break;
        case WM_LBUTTONDOWN:
					{
						HMENU hmenu;
						POINT pt;
						RECT	r;
						int selected, i;
						pt.x = LOWORD(lParam); 
						pt.y = HIWORD(lParam); 
						r.left = X_OFF;
						r.top = Y_OFF;
						r.bottom = Y_OFF + DISPLAY_HEIGHT;
						r.right = X_OFF + DISPLAY_WIDTH;
						if (PtInRect(&r, pt)){
							hmenu = CreatePopupMenu();
							GetCursorPos(&pt);
							
							for (i = 0; i < RANGE_COUNT; i++)
								AppendMenu(hmenu, MF_ENABLED, i +1, spans[i].name); 
							AppendMenu(hmenu, MF_SEPARATOR, i +1, ""); 	
							AppendMenu(hmenu, MF_ENABLED, MARK_ME, "Mark"); 
							selected = TrackPopupMenu(hmenu, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, 0, mainWnd, 0);
							DestroyMenu(hmenu);
							if (selected == MARK_ME){
								setMark(pt.x, pt.y);
							}
							else if (selected > 0){
								selected--;
								int f1, f2;
								f1 = currentFrequency - spans[selected].width;
								f2 = currentFrequency + spans[selected].width;
								if (f1 < 0){
									f1 = 0;
									f2 = spans[selected].width * 2;
								}
								if (f2 > 100000000){
									f2 = 100000000;
									f1 = f2 - (spans[selected].width * 2);
								}
								centerFreq = (f1+f2)/2;
								selectedSpan = selected;
								loadControls(hwnd);
								setupSweep();
								startSweep(); 
							}
						}
					}		
          InvalidateRect(mainWnd, NULL, TRUE);
          UpdateWindow(mainWnd);
          break;
        case WM_RBUTTONDOWN:
          /* if (continuous)
            continuous  = 0;
          else
            continuous = 1; */
          setupSweep();
          startSweep();
          break;
        case WM_DESTROY:
        case WM_CLOSE:
          PostQuitMessage (0);       /* send a WM_QUIT to the message queue */
          break;
        default:                      /* for messages that we don't deal with */
          return DefWindowProc (hwnd, message, wParam, lParam);
    }

    return 0;
}

int startEverything(){
    WNDCLASSEX wincl;        /* Data structure for the windowclass */
    INITCOMMONCONTROLSEX initex;
    LOGFONT lf;
    int i;

    background = (HBRUSH) CreateSolidBrush(RGB(192,192,255));//GetStockObject(WHITE_BRUSH);
    
    /* initialize some cool controls */
    initex.dwSize = sizeof(initex);
    initex.dwICC =  ICC_BAR_CLASSES | ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&initex); 

    /* The Window structure */
    wincl.hInstance = instance;
    wincl.lpszClassName = szClassName;
    wincl.lpfnWndProc = WindowProcedure;      /* This function is called by windows */
    wincl.style = CS_DBLCLKS;                 /* Catch double-clicks */
    wincl.cbSize = sizeof (WNDCLASSEX);

    /* Use default icon and mouse-pointer */
    wincl.hIcon = LoadIcon (NULL, IDI_APPLICATION);
    wincl.hIconSm = LoadIcon (NULL, IDI_APPLICATION);
    wincl.hCursor = LoadCursor (NULL, IDC_ARROW);
    wincl.lpszMenuName = MAKEINTRESOURCE(IDR_SWEEPER_MENU);                 /* No menu */
    wincl.cbClsExtra = 0;                      /* No extra bytes after the window class */
    wincl.cbWndExtra = 0;                      /* structure or the window instance */
    /* Use Windows's default color as the background of the window */
    wincl.hbrBackground = background;//GetStockObject(WHITE_BRUSH);;

    
    /* Register the window class, and if it fails quit the program */
    if (!RegisterClassEx (&wincl))
        return 0;

    memset(&lf, 0, sizeof(lf));
    lf.lfHeight = 10;
    strcpy(lf.lfFaceName, "Courier");
    fontFixed = CreateFontIndirect(&lf);
		
		lf.lfHeight = TEXT_HEIGHT;
		strcpy(lf.lfFaceName, "Sans Serif");
		fontText = CreateFontIndirect(&lf);

    penBlue = CreatePen(PS_SOLID, 1, RGB(160,160,255));
    penGray = CreatePen(PS_SOLID, 1, RGB(160,160,160));
    penRed  = CreatePen(PS_SOLID, 1, RGB(255,0,0));
		penBlack= CreatePen(PS_SOLID, 2, RGB(0,0,0));


    /* The class is registered, let's create the program*/
    mainWnd = CreateWindowEx (
           0,                   /* Extended possibilites for variation */
           szClassName,         /* Classname */
           "SpecAn",       /* Title Text */
					 WS_OVERLAPPEDWINDOW,
           50,       /* Windows decides the position */
           50,       /* where the window ends up on the screen */
           1000,                 /* The programs width */
           630,                 /* and height in pixels */
           (HWND)HWND_DESKTOP,        /* The window is a child-window to desktop */
           (HMENU)NULL,                /* No menu */
           (HINSTANCE)instance,       /* Program Instance handler */
           NULL                 /* No Window Creation data */
           );

    statusWnd = CreateWindowEx(0, STATUSCLASSNAME, NULL,
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0,
        mainWnd, (HMENU)IDC_MAIN_STATUS, instance, NULL);
    SendMessage(statusWnd, SB_SETPARTS, sizeof(statusWidths)/sizeof(int), (LPARAM)statusWidths);        
    
    setStatus("Choose 'Connect' from Action menu...");  
    
    /* set the refreadings to some nominal value  from 0 to 200 MHz in 100 Khz steps*/
    for(i=0; i < MAX_READINGS; i++){
      refReadings[i].frequency = i * 100000;
      refReadings[i].power = 390;
    }

  loadCaliberation();    
  return 1;
}



int WINAPI WinMain (HINSTANCE hThisInstance,
                    HINSTANCE hPrevInstance,
                    LPSTR lpszArgument,
                    int nFunsterStil)

{
    HWND hwnd;               /* This is the handle for our window */
    MSG messages;            /* Here messages to the application are saved */
    WNDCLASSEX wincl;        /* Data structure for the windowclass */
    INITCOMMONCONTROLSEX initex;
    LOGFONT lf;

    instance = hThisInstance;
  
    startEverything();    

    /* Make the window visible on the screen */
    ShowWindow (mainWnd, nFunsterStil);
        
    /* Run the message loop. It will run until GetMessage() returns 0 */
    while (GetMessage (&messages, NULL, 0, 0)){
        /* Translate virtual-key messages into character messages */
        TranslateMessage(&messages);
        /* Send message to WindowProcedure */
        DispatchMessage(&messages);
    }

    CloseHandle(serialPort);
    /* The program return-value is 0 - The value that PostQuitMessage() gave */
    return messages.wParam;
}



