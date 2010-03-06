/*
	File: MainWnd.cpp

	ListeningNowTracker application.
	
	This app was created to link Spotify and Skype together. The app shows the title
	of the currently playing music track in your Skype profile.

	Spotify sends a "now playing" events for MSN-Messenger. Latest Spotify client 
	supports also Last.FM scrobbling, which is nice. However, it leaves applications
	like Skype without this support.

	This application sits in a background and listens the "Listening Now" message meant 
	to go to MSN and re-routes the message to Skype ("mood message" of Skype profile).

	At the moment supports following players
	- Spotify
	- MS MediaPlayer
	- plus various other apps supporting MSN Messenger "MsnMsgrUIManager" window event

	Supported target apps to show the "listening now" text
	- Skype (mood text of the currently logged in profile)
	- Well, nothing else at the moment but it would be easy to add command line/external applicaiton
	  linking to do whatever custom app wants to with the track events.

	Note! 
	
	This app requires that Skype was installed with "extension manager" plugin,
	which is the case in default Skype installation. However, user can leave out
	the plugin in custom installation of Skype. This app doesn't really need
	the "Extension Manager" of Skype, but it provides "Skype4OLE" OLE library which
	is required.
	
	If user did not install that extension with Skype then one must download 
	Skype4OLE SDK file from Skype web site and register the missing OLE object.
	Following cmdline command registers the library.

		regsvr32 Skype4COM.dll
	
	Copyright (c) 2010 Mika Nieminen. All rights reserved.
	
	See LICENSE.TXT for more information about the license (opensource, free to use license).

 3rd party libraries used in this application (why re-invent the wheel when someone has already done it better?):

 NOTE! There is no need to download these libraries because the source code ZIP file of this
       application already contains these files. However, if you find bugs in any of those libs
	   then feel free to download up-to-date version of the library from the original web site.

    VOLE/COMSTL:  http://vole.sourceforge.net/
		VOLE/COMSTL is Windows OLE API library. The source code ZIP package contains these 
		libraries already, but feel free to visit the web page to download more recent version 
		of those libraries if needed. VOLE is great library for pure C++ applications to 
		use OLE automation objects. No need to use MFC/ATL libraries.

	Other notes:
		See also http://code.google.com/p/scrobblify/ application. Part of the source code
		is derived from that application (modified, but credits from usage of MSN WM_COPYDATA event
		goes there).

*/

#include "stdafx.h"						// VC++ trick to speed up the process of using common header files

#include <vole/vole.hpp>				// VOLE+STLSoft OLE libraries. Absolutely fantastic libraries to 
#include <comstl/util/initialisers.hpp> // to utilize OLE objects from pure C++ apps. Used to communicate with Skype OLE objects.

#include "resource.h"					// Windows API resource definitions (tray icon etc)

#include "CThread.h"					// Thread wrapper
#include "CIniFile.h"				    // INI file handler


const LPTSTR g_szAppName = _T("ListeningNowTracker"); 

// Window class name used to receive Spotify NowPlaying messages (meant to go to MSN)
// and the ID number of that special event
const LPTSTR g_szMsn_WindowClassName   = _T("MsnMsgrUIManager"); 
const ULONG  g_iMsn_NowPlayingEventNum = 0x547; 


//
// Global variables
//

HINSTANCE g_hMainAppInstance;	// Main app instance handle
HWND      g_hMainWnd;			// Main wnd handle

std::wstring g_strListeningNowText; // Format mask for "Listening now" text shown in Skype profile (INI file parameter)


// 
// Global "shared resources" for the process (all threads share these values)
//
CThread          g_objThreadWatchDog;  // WatchDog thread to clear Skype MoodText in case MusicPlayer has crashed
CCriticalSection g_objProcessCS;	   // CriticalSection object to control the usage of shared resources

NOTIFYICONDATA   g_ToolbarTrayIcon;			    // Toolbar tray icon object

BOOL			 g_bProcessRunning;	             // TRUE=Process is valid, FALSE=Process is closing. Do nothing in child threads except closing immediately
DWORD			 g_dwLastTrackChangeTimeStampMS; // The timestamp of the last received "track changed" event


//--------------------------------------------------------
// Convert CHAR string to WCHAR string (brute-force-method)
//
// Found from http://www.daniweb.com/forums/thread249262.html post by vijayan121.
// Brute-force and probably not the most "official" way to do it
// but does the trick in this case.
//
std::wstring str2wstr(const LPCSTR szText)
{
	std::string  strText (szText);
	std::wstring wstrText(strText.size(), ' ');
	std::copy(strText.begin(), strText.end(), wstrText.begin());
	return wstrText;
}


//----------------------------------------------------------
// Initialize or update tray icon and tooltip text
//   If hWND = NULL then change the tooltip text of existing tray icon
//   If hWND != NULL then create a new tray icon (done once during initialization of the main wnd)
//
void InitTray(HWND hWnd, const std::wstring& strTrayIconText) 
{ 
	// Max length of tip text (leave room for null char just in case)
	int iMaxTipTextSize = (sizeof(g_ToolbarTrayIcon.szTip) / sizeof(TCHAR)) - sizeof(TCHAR);

	if(hWnd != NULL) 
	{
		// Create a new tray icon
		g_ToolbarTrayIcon.cbSize = sizeof(g_ToolbarTrayIcon); 
		g_ToolbarTrayIcon.uID    = IDM_TRAYICON; 

		g_ToolbarTrayIcon.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP; 
		g_ToolbarTrayIcon.hIcon  = LoadIcon(g_hMainAppInstance, MAKEINTRESOURCE(IDI_TRAYICON)); 
		g_ToolbarTrayIcon.hWnd   = hWnd; 
		g_ToolbarTrayIcon.uCallbackMessage = WM_USER; 
	}
	else 
	{
		// Update tooltip text of the existing tray icon
		g_ToolbarTrayIcon.uFlags = NIF_TIP;
	}

	// Substr text to make sure the tooltip text doesn't overflow the max size of szTip array
	wcscpy_s(g_ToolbarTrayIcon.szTip, strTrayIconText.substr(0,iMaxTipTextSize).c_str()); 

	if(hWnd != NULL) Shell_NotifyIcon(NIM_ADD, &g_ToolbarTrayIcon); 
	else if (g_ToolbarTrayIcon.uID != 0) Shell_NotifyIcon(NIM_MODIFY, &g_ToolbarTrayIcon);
} 

void UpdateTrayText(const std::wstring& strTrayIconText) 
{
	InitTray(NULL, strTrayIconText);
}


//--------------------------------------------------------
// Update Skype mood text using Skype4OLE interface (comes with Skype Windows client).
//
void UpdateSkypeMoodText(const std::wstring& strMoodText)
{
	using vole::object;

	g_objProcessCS.Enter();

  try
  {
	// If there is no "active song title" and new text is also empty then no need to do anything
	if (g_dwLastTrackChangeTimeStampMS == 0 && strMoodText.empty())
	{
		// Do nothing
	}
	else
	{
		// Update the timestamp of the last change of Skype MoodText.
		// If the new text is empty string then set timestamp to zero to indicate that
		// there is no "active song title" in the Skype MoodText property.
		//
		if (strMoodText.empty()) g_dwLastTrackChangeTimeStampMS = 0;
		else g_dwLastTrackChangeTimeStampMS = GetTickCount();
		  
		object objSkype = object::create(L"Skype4COM.Skype");
		object objClient = objSkype.get_property<object>(L"Client");

		// TODO: Should we start Skype automatically if it's not running?
		//       OLE method "Start(true,true)" would start Skype minimized and no splash screen.
		//if ( objClient.get_property<bool>(L"IsRunning") == false)
		//{
		//	objClient.invoke_method_v(L"Start", true, true);
		//	Sleep(1000 * 5);
		//}
		//

		if ( objClient.get_property<bool>(L"IsRunning") )
		{
			object objProfile = objSkype.get_property<object>(L"CurrentUserProfile");
			//std::wstring strMood = objProfile.get_property<std::wstring>(L"MoodText");
			objProfile.put_property(L"MoodText", strMoodText.c_str());
		}
		else
			UpdateTrayText(std::wstring(L"WARNING: Skype is not running. Cannot update profile text"));
	}

  } catch (/*...*/ std::exception &x ) { 
	UpdateTrayText( std::wstring(L"ERROR: ").append( str2wstr(x.what()) ) );
  }

  	g_objProcessCS.Leave();
}


//---------------------------------------------------------
// Application is closing. Cleanup everything.
//
void CleanupApplication(void)
{
	// App is already "cleaned up". No need to do it twice (app cannot revert its
	// status back from FALSE to TRUE state)
	if (g_bProcessRunning == FALSE) return;

	g_objProcessCS.Enter();

  try
  {
	if (g_bProcessRunning)
	{
		// Flag this application for closing. Nothing can be done anymore with shared resources
		g_bProcessRunning = FALSE;

		//
		// Signal thread to stop. Note! Doesn't use forceKill because the thread probably stops
		// by the time this process is ready to close. Anyway, there is "Stop(TRUE) forceKill"
		// command at the end of this process just-in-case.
		//
		g_objThreadWatchDog.Stop();

		// Set empty "Skype mood text" because this app no longer monitors
		// the Spotify "Playing" events (otherwise Skype would show the last text permanently)
		if (g_ToolbarTrayIcon.uID != 0)	
		{
			UpdateSkypeMoodText(std::wstring());
			Shell_NotifyIcon(NIM_DELETE, &g_ToolbarTrayIcon); 
		}
		g_ToolbarTrayIcon.uID = 0;
	}
  }
  catch (...)
  {
	  // Do nothing special
  }

	g_objProcessCS.Leave();
}

//-----------------------------------------
// Something went wrong while initializing the app. Relase mutex object
// if it already exist and return "false" as "do not continue running the app"
//
bool AbnormalAppClosing(void)
{
	CleanupApplication();
	return false;
}


//-------------------------------------------------------
// Forwards a message to MSN Live Messenger. 
// Helps to workaround the problem with LiveMessenger (blocks WM_COPYDATA messages from this tool?) 
/**
void NotifyMsnMessenger(PCOPYDATASTRUCT data) 
{ 
	for (HWND w = NULL; (w = FindWindowEx(NULL, w, g_szMsn_WindowClassName, NULL)) != NULL; ) 
	{ 
		// Do not re-send the message to ourself
		if (w != g_hMainWnd) 
		{ 
			// SendNotifyMessage doesn't block like SendMessage 
			SendNotifyMessage(w, WM_COPYDATA, (WPARAM) kMsnMagicNumber, (LPARAM) data); 
		} 
	} 
} 
***/


//-------------------------------------------------- 
// Process "Listening song" WM_COPYDATA event. 
// Data is expected to be in "\0Music\0<status>\0<format>\0<song>\0<artist>\0<album>\0" format
//
// Update Skype mood text based on the song title and artist texts.
//
// Parsing of lpData data derived from http://code.google.com/p/scrobblify/ application (with modifications).
//
LRESULT CALLBACK ProcessWMCopyDataEvent(HWND /*hWnd*/, WPARAM /*wParam*/, LPARAM lParam) 
{ 
	// Max text of "Listening" text is 200 chars in this app. Feel free to increase if necessary
	WCHAR szBuffer[200];
	size_t iPos;

	PCOPYDATASTRUCT cds = (PCOPYDATASTRUCT) lParam; 
	std::wstring data   = (TCHAR *) cds->lpData;

	// TODO: uncomment when this works 
	// NotifyMsnMessenger(cds); 

	// Field delimiter
	std::wstring delimiter = _T("\\0");			
	size_t delimiter_size  = delimiter.size(); 

	// The first tag must be "\0Music\0", otherwise this msg is something we don't know about
	std::wstring sTagMusic = delimiter + _T("Music") + delimiter; 

	iPos = sTagMusic.size(); 
	if (data.compare(0, iPos, sTagMusic) != 0) 
	{
		return 0;  // Hmmm.. Unknown prefix in the data. Do nothing.
	}

	// Status: 1=Playing, 0=Stopped or Paused
	std::wstring status = data.substr(iPos, data.find(delimiter, iPos) - iPos);  
	iPos += status.size() + delimiter_size; 

	// Format: ???
	std::wstring format = data.substr(iPos, data.find(delimiter, iPos) - iPos);  
	iPos += format.size() + delimiter_size; 

	// Title of the song
	std::wstring title = data.substr(iPos, data.find(delimiter, iPos) - iPos); 
	iPos += title.size() + delimiter_size; 

	// Artist of the song
	std::wstring artist = data.substr(iPos, data.find(delimiter, iPos) - iPos); 
	iPos += artist.size(); 

	// Format "Listening" text string (too bad std:wstring doesn't have built-in printf 
	// formatter, so we have to do it through old-fashioned temp buffer.
	_snwprintf_s(szBuffer, (sizeof(szBuffer) / sizeof(WCHAR)) - sizeof(WCHAR), _TRUNCATE, 
		g_strListeningNowText.c_str(),
		title.c_str(), 
		artist.c_str()
	);
	std::wstring strListeningText = szBuffer;

	// Update tray tooltip
	UpdateTrayText(strListeningText);
	
	// Update Skype mood text or clear it if song is stopped/paused/arist-title text is empty
	if ( status.compare(_T("0")) == 0 || ( title.empty() && artist.empty() ) ) 
		UpdateSkypeMoodText(std::wstring());
	else	
		UpdateSkypeMoodText(strListeningText);

	return 0; 
} 


//---------------------------------------------------------
// Show trayicon popup menu. Menu events are sent to the message
// loop of the main window (WM_COMMAND events).
//
void ShowContextMenu(HWND hWnd) 
{ 
	POINT pt; 
	GetCursorPos(&pt); 
	HMENU hTrayMenu = CreatePopupMenu(); 

	if (hTrayMenu) 
	{ 
		AppendMenu(hTrayMenu, MF_STRING, WM_DESTROY, _T("Exit")); 
		//AppendMenu(hTrayMenu, MF_STRING, WM_APP+1, _T("About")); 

		SetForegroundWindow(hWnd); 
		TrackPopupMenu(hTrayMenu, TPM_BOTTOMALIGN, pt.x, pt.y, 0, hWnd, NULL); 
		DestroyMenu(hTrayMenu); 
	} 
} 

//---------------------------------------------------------
// Message handler of the main window
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) 
{ 
	int wm_id;
	//int wm_event; 

	switch (message) 
	{ 
		case WM_CREATE: 
			// Create a new tray icon (default tooltip is the application name)
			InitTray(hWnd, std::wstring(g_szAppName)); 
			break; 

		case WM_DESTROY: 
			CleanupApplication();
			PostQuitMessage(0); 
			break; 

		case WM_USER: 
			switch (lParam) 
			{ 
				case WM_RBUTTONDOWN: 
				case WM_CONTEXTMENU: 
					ShowContextMenu(hWnd); 
					break; 
			} 
			break; 

		case WM_COMMAND: 
			wm_id    = LOWORD(wParam); 
			//wm_event = HIWORD(wParam); 

			switch (wm_id) 
			{ 
				case WM_DESTROY: 				
					// "Exit" command from trayicon popup menu
					DestroyWindow(hWnd); 
					break; 

				// case WM_APP+1:
				//	DoTest();
				//	break;
			} 
			break; 

		case WM_COPYDATA: 
			// Is this "Listening" event from Spotify?
			if (((PCOPYDATASTRUCT) lParam)->dwData == g_iMsn_NowPlayingEventNum) 
			{ 
				ProcessWMCopyDataEvent(hWnd, wParam, lParam); 
			} 
			break; 
	} 

	return DefWindowProc(hWnd, message, wParam, lParam); 
} 


//---------------------------------------------------------
// Register window class
//
bool RegisterAppWndClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex; 

	// Register window class using the class name used by Spotify to notify MSN about "listening" events
	wcex.cbSize      = sizeof(wcex); 
	wcex.style       = 0; 
	wcex.lpfnWndProc = WndProc; 
	wcex.cbClsExtra  = 0; 
	wcex.cbWndExtra  = 0; 
	
	wcex.hInstance = hInstance; 
	wcex.hIcon     = LoadIcon(hInstance, (LPCTSTR) IDI_TRAYICON); 
	wcex.hCursor   = LoadCursor(NULL, IDC_ARROW); 
	
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW+1); 
	wcex.lpszMenuName  = NULL; 
	wcex.lpszClassName = g_szMsn_WindowClassName; 
	wcex.hIconSm       = NULL; 
	
	if (RegisterClassEx(&wcex) == 0) return false;
	else return true;
}

//---------------------------------------------------------
// Initialize the instance of the application
//
bool InitInstance(HINSTANCE hInstance, int nCmdShow) 
{ 
	// Check if MSN is running 
//	if (FindWindowEx(NULL, NULL, g_szMsn_WindowClassName, NULL) != NULL) 
//	{ 
//		MessageBox(NULL, 
//            _T("MSN Live Messenger running. Quite MSN if you want to trap the NowPlaying events"), 
//			_T("ERROR: MSN Live Messenger is running"), 
//            MB_ICONWARNING | MB_TOPMOST | MB_OK); 
//	} 

	g_hMainAppInstance = hInstance; 
	g_hMainWnd = CreateWindow(g_szMsn_WindowClassName, g_szAppName, 
                   WS_MINIMIZE, 
                   CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 
                   HWND_MESSAGE, 
                   NULL, 
                   hInstance, 
                   NULL); 

	if (!g_hMainWnd)
	{
		// Hmmm.. Failed to create the main window. Can't do much without it, so quit
		return false; 
	}

	// ShowWindow(g_hMainWnd, nCmdShow); 
	UpdateWindow(g_hMainWnd); 
	return true; 
} 


//----------------------------------------------------
// WatchDog handler thread to reset Skype mood text back to empty string if
// the same song title has been as "ListeningNow" text more than X minutes
// (maybe MusicPlayer crashed and doesn't send anymore change events?)
//
// Note! This function is executed in a separate thread (see CThread)
//
unsigned __stdcall ThreadWatchDogHandler(void *pArg)
{
	DWORD dwCurrentTimeStampMS;
	DWORD dwResetPeriodInMS;

	CThreadContext *objThreadCtx = (CThreadContext*) pArg;	

	// The main thread sends the "X minutes" parameter in a userdata pointer of the thread context object.
	// However, WaitForSingleObject timeout uses milliseconds, so multiply the "minute" value to MS value
	dwResetPeriodInMS = *((DWORD*)objThreadCtx->m_pUserData);
	dwResetPeriodInMS = 1000 * 60 * dwResetPeriodInMS;

	// Thread needs to do its own OLE initialization or it fails to use Skype OLE object
	comstl::com_initialiser coinit;

	while (g_bProcessRunning) 
	{ 
		// Sleep X minutes and check whether the song title is still the same or
		// until the thread is signaled to stop (ie. hStopEvent is something else than TIMEOUT signaled)
		if (::WaitForSingleObject(objThreadCtx->m_hStopEvent, dwResetPeriodInMS) != WAIT_TIMEOUT) break;
		if (g_bProcessRunning == FALSE) break;

		g_objProcessCS.Enter();
	   try
	   {
	    // Do nothing if Skype MoodText is not set (ie. empty text because no song playing)
		if (g_dwLastTrackChangeTimeStampMS != 0) 
		{
			dwCurrentTimeStampMS = ::GetTickCount();
			if (dwCurrentTimeStampMS < g_dwLastTrackChangeTimeStampMS)
			{
				// GetTickCount wraps around back to zero if Windows is not rebooted within 49 days.
				// In that case do not reset the Skype mood text, just wait for the next timeout and 
				// this watchdog is back in normal cycle
				g_dwLastTrackChangeTimeStampMS = dwCurrentTimeStampMS;
			}
			else if ( (dwCurrentTimeStampMS - g_dwLastTrackChangeTimeStampMS) >= dwResetPeriodInMS)
			{
				// It's been more than 10 minutes since the song title was changed. Reset Skype mood text to empty string.
				UpdateSkypeMoodText(std::wstring(L""));				
			}
		}
	   }
	   catch(...)
	   {
	   }
		g_objProcessCS.Leave();
	}

	// CRT _beginthreadex requires _endthreadex within the thread to signal and cleanup the thread
	_endthreadex(0);
	return 0;
}


//----------------------------------------------------
// MAIN procedure. Everything starts from here
//
int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpCmdLine*/, int nCmdShow) 
{ 
	MSG	  msg; 
	DWORD dwSongTitleResetPeriodInMins;

	CIniFile objAppINIFile(CIniFile::GetApplicationPath().append(L"\\ListeningNowTracker.ini").c_str());
	CMutex   objProcessMutex(std::wstring(L"mutex_").append(g_szAppName).c_str());

	// ::MessageBox(NULL, L"test", L"title", MB_ICONWARNING | MB_TOPMOST | MB_OK); 

	// Initialize shared resources
	ZeroMemory(&g_ToolbarTrayIcon, sizeof(g_ToolbarTrayIcon)); 
	g_bProcessRunning = TRUE;
	g_dwLastTrackChangeTimeStampMS = 0;

	// Proceed to initialize the application

	// Makes sure this is the only instance of this application. Quit if app is already running
	if (objProcessMutex.GetShareCount() != 1)
		return AbnormalAppClosing();

	// Register window class using the class name used by Spotify to notify MSN about "listening" events
	if (!RegisterAppWndClass(hInstance))
		return AbnormalAppClosing();

	// Initialize the application
	if (!InitInstance(hInstance, nCmdShow)) 
		return AbnormalAppClosing();

	// Initialize OLE APIs (used to communicate with Skype API)
	comstl::com_initialiser coinit;

	// Start a watchdog thread (resets Skype MoodText back to empty string if song 
	// title haven't changed in X minutes. It is assumed that MusicPlayer has crashed or quit)

	g_strListeningNowText        = objAppINIFile.ReadString (L"CONFIG", L"ListeningNowText", L"Listening '%1s' by %2s");
	dwSongTitleResetPeriodInMins = objAppINIFile.ReadInteger(L"CONFIG", L"WatchDogTimerInMins", 10);

	g_objThreadWatchDog.Attach(ThreadWatchDogHandler);
	g_objThreadWatchDog.Start(&dwSongTitleResetPeriodInMins);
 
	// Start the main message loop
	while(GetMessage(&msg, NULL, 0, 0)) 
	{ 
		TranslateMessage(&msg); 
		DispatchMessage(&msg); 
	} 

	// Terminate the watchdog thread brutally just in case it is not yet terminated.
	// Also, call CleanupApplication just in case (should have been called already as WM_DESTROY message handling)
	g_objThreadWatchDog.Stop(TRUE);
	CleanupApplication();

	return ((int) msg.wParam); 
} 

