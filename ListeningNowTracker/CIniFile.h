#ifndef __CINIFILE_H__
#define __CINIFILE_H__

#include <string> 

/*
 * Simple class to handle Windows INI files (read support only here)
*/

class CIniFile
{
  protected:
	  std::wstring m_strFileName;

  public:
		CIniFile(const TCHAR* szFileName)
		{
			m_strFileName = szFileName;
		}

		// Class function to return the path of the executable file
		static std::wstring GetApplicationPath()
		{
			TCHAR szAppFileName[_MAX_PATH + 1];
			std::wstring strResult;

			::GetModuleFileName(NULL, szAppFileName, _MAX_PATH); 
			
			strResult = szAppFileName;
			return strResult.substr(0, strResult.rfind('\\'));
		}

		int ReadInteger(const TCHAR* szSection, const TCHAR* szKey, int iDefaultValue = 0)
		{
			return ::GetPrivateProfileInt(szSection, szKey, iDefaultValue, m_strFileName.c_str()); 
		}

		std::wstring ReadString(const TCHAR* szSection, const TCHAR* szKey, const TCHAR* szDefaultValue)
		{
			TCHAR szResult[255];
			DWORD dwResultLen;
			
			dwResultLen = ::GetPrivateProfileString(szSection, szKey, szDefaultValue, szResult, 255-1, m_strFileName.c_str()); 
			if (dwResultLen > 0 && dwResultLen < 255)
			{
				// Make sure the text value is null-terminated (usually it is, but this doesn't hurt anything)
				szResult[dwResultLen] = '\0';
				return std::wstring(szResult);
			}
			else
			{
				return std::wstring(szDefaultValue);
			}
		}
};

#endif //__CINIFILE_H__
