#ifndef __CTHREAD_H__
#define __CTHREAD_H__

#include <process.h>

/*
   Simple Win32 thread library (better alternative would have been Boost thread component (www.boost.org),
   but it would come with lots of unnecessary bits-and-pieces for this application. 

   Basis for this code was derived from http://www.codeguru.com/cpp/misc/misc/threadsprocesses/article.php/c3793
   web site. The version here is modifed from that original code (fixed the code to use CRT compatible
   _beginthreadex call and added Mutex/CriticalSection/StopEvent support).
   
   The original code was without any license and copyright noticies, so I assume it was released in
   public domain as such.
*/

// CRT _begintthreadex compatible function pointer (thread handler func)
typedef unsigned (__stdcall *LPTHREAD_START_ROUTINE_CRT) (void *);


//------------------------------------------------------------------
// Ûber simple critical section object. 
//
class CCriticalSection
{
  protected:
	CRITICAL_SECTION m_hCriticalSection;

  public:
	CCriticalSection()
	{
		::InitializeCriticalSection(&m_hCriticalSection);
	}

	~CCriticalSection()
	{
		::DeleteCriticalSection(&m_hCriticalSection);
	}

	// Enter and block execution while CS acquired 
	void Enter()
	{
		::EnterCriticalSection(&m_hCriticalSection);
	}

	// Leave CS and let other threads to do whatever they want within the CS
	void Leave()
	{
		::LeaveCriticalSection(&m_hCriticalSection);
	}
};


//------------------------------------------------------------------
// Über simple mutex object. This app uses mutex obj to make sure 
// there is only one instance of this process running.
//
class CMutex
{
  protected:
	  HANDLE m_hMutex;		// Handler of the mutex obj
	  int    m_iShareCount; // 1 = This is the only process using this mutex, 2 = More than one process using the same mutex

  public:
	CMutex(const TCHAR* sMutexName)
	{
		m_hMutex = ::CreateMutex(NULL, FALSE, sMutexName); 
		
		if (GetLastError() == ERROR_ALREADY_EXISTS) m_iShareCount = 2; // Shared mutex
		else if (m_hMutex == NULL)  m_iShareCount = 0; // Error. Failed to create mutex
		else m_iShareCount = 1; // Single process using the mutex
	}

	~CMutex()
	{
		if(m_hMutex != NULL) ::CloseHandle(m_hMutex);
		m_hMutex = NULL;
	}

	int GetShareCount() const { return m_iShareCount; } // 1 = Single process, >1 = Shared mutex
};


//--------------------------------------------------------------------------------
// Simple class to manage multithreaded processing (thread starting and stopping logic)
// 

// Every thread object uses certain "thread context" values (thread handles, user data pointer, etc)
class CThreadContext
{
  public:
	UINT   m_dwTID;				// Thread ID
	HANDLE m_hThread;			// Thread Handle
	HANDLE m_hStopEvent;		// Event used to stop the thread gracefully (Stop method posts this event)
	LPVOID m_pUserData;			// User data pointer
	DWORD  m_dwExitCode;		// Exit Code of the thread

 public:
	CThreadContext()
	{
		// Quick and dirty way to "zero out" all member variables
		memset(this, 0, sizeof(CThreadContext));
	}
};

//
// CThread - Thread class itself (no need to use CThreadContext class directly, because CThrread has it as member object)
//
class CThread
{
	protected:
		CThreadContext		m_ThreadCtx;	//	Thread Context member
		LPTHREAD_START_ROUTINE_CRT	m_pThreadFunc;	//	Worker Thread Function Pointer

	public:
		CThread()
		{ 
			m_pThreadFunc = NULL;
		}

		CThread(LPTHREAD_START_ROUTINE_CRT lpExternalRoutine)
		{
			// Attach thread handler to function pointer provided through constructor 
			Attach(lpExternalRoutine);
		}

		~CThread()
		{
			// Stop thread if it is still running and close "stop event" handle object
			if ( m_ThreadCtx.m_hThread ) Stop(true);
			if ( m_ThreadCtx.m_hStopEvent ) ::CloseHandle(m_ThreadCtx.m_hStopEvent);
		}

		/*
		 *	This function starts the thread pointed by m_pThreadFunc
		 */
		DWORD Start( void* arg = NULL )
		{
			// If thread handler is missing then do nothing (no Attach method called)
			if (m_pThreadFunc == NULL) return -1;

			m_ThreadCtx.m_pUserData = arg;

			// Begin a new thread and call handler function with threadCtx object pointer as parameter.
			// If thread creation succeeded then create a "stop thread" event handle (themain process can signal the thread to quit)
			m_ThreadCtx.m_hThread = (HANDLE) _beginthreadex(NULL, 0, m_pThreadFunc, &m_ThreadCtx, 0, &m_ThreadCtx.m_dwTID);
			if (m_ThreadCtx.m_hThread != 0) m_ThreadCtx.m_hStopEvent = ::CreateEvent(NULL, FALSE, FALSE, NULL);

			m_ThreadCtx.m_dwExitCode = (DWORD)-1;

			return GetLastError();
		}

		/*
		 *	This function stops the current thread. 
		 *	We can force kill a thread which results in a TerminateThread.
		 */
		DWORD Stop ( bool bForceKill = false )
		{
			BOOL bSucceeded;

			if ( m_ThreadCtx.m_hThread )
			{
				// Signal the thread to stop its processing loop
				::SetEvent(m_ThreadCtx.m_hStopEvent);
			
				//
				// If ForceKill=TRUE then make sure the thread is closed before proceeding.
				// Give couple secs for the thread to close and use brute-force TerminateThread 
				// if the thread didn't wanna quit after all (hanging maybe?)
				//
				bSucceeded = ::GetExitCodeThread(m_ThreadCtx.m_hThread, &m_ThreadCtx.m_dwExitCode);

				if (bForceKill && bSucceeded)
				{
					for (int idx = 0; idx < 20; idx++)
					{
						// Thread is no longer running, so exit the test loop
						if ( bSucceeded && m_ThreadCtx.m_dwExitCode != STILL_ACTIVE) break;

						// Sleep half a sec and test again whether the thread has terminated already
						::Sleep(500);
						bSucceeded = ::GetExitCodeThread(m_ThreadCtx.m_hThread, &m_ThreadCtx.m_dwExitCode);
					}

					if ( bSucceeded && m_ThreadCtx.m_dwExitCode == STILL_ACTIVE)
					{
						// Thread still running. Oh well, let's terminate it brutally even when it is a bit dangerous
						::TerminateThread(m_ThreadCtx.m_hThread, DWORD(-1));
					}
				}

				if (bForceKill)
				{
					// Thread was created with _beginthreadex, so the main process must call 
					// CloseHandle after it is no longer interested in the thread handle.
					::CloseHandle(m_ThreadCtx.m_hThread);
					m_ThreadCtx.m_hThread = NULL;
				}
			}

			return m_ThreadCtx.m_dwExitCode;
		}


		/*
		 *	Attaches a Thread handler function pointer to the thread object.
		 *	Start method calls this handler function when thread is started.
		 */
		void Attach( LPTHREAD_START_ROUTINE_CRT lpThreadFunc )
		{
			m_pThreadFunc = lpThreadFunc;
		}

		/*
		 *	Detaches the Attached Thread Function
		 */
		void  Detach( void )
		{
			m_pThreadFunc = NULL;
		}
};

#endif //__CTHREAD_H__
