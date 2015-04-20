#include "stdafx.h"
#include "Windows.h"
#include "Oleacc.h"
#include "Psapi.h"
#include <iostream>

//#define _WIN32_WINNT 0x0602

#ifdef _UNICODE
#define tcout std::wcout
#define tcerr std::wcerr
#else
#define tcout std::cout
#define tcerr std::cerr
#endif

#define NUMCHARS(a) (sizeof(a)/sizeof(*a))

void CALLBACK HandleWinEvent(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
	LONG idObject, LONG idChild,
	DWORD dwEventThread, DWORD dwmsEventTime);
void InitializeMSAA();
void ShutdownMSAA();

HWINEVENTHOOK g_hook;

void InitializeMSAA()
{
	CoInitialize(NULL);
	g_hook = SetWinEventHook(
		EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
		NULL,
		HandleWinEvent,
		0, 0,
		WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
}

void ShutdownMSAA()
{
	UnhookWinEvent(g_hook);
	CoUninitialize();
}

static HRESULT NormalizeNTPath(wchar_t* pszPath, size_t nMax)
{
	wchar_t* pszSlash = wcschr(&pszPath[1], '\\');
	if (pszSlash) pszSlash = wcschr(pszSlash + 1, '\\');
	if (!pszSlash)
		return E_FAIL;
	wchar_t cSave = *pszSlash;
	*pszSlash = 0;

	wchar_t szNTPath[_MAX_PATH];
	wchar_t szDrive[_MAX_PATH] = L"A:";
	for (wchar_t cDrive = 'A'; cDrive < 'Z'; ++cDrive)
	{
		szDrive[0] = cDrive;
		szNTPath[0] = 0;
		if (0 != QueryDosDevice(szDrive, szNTPath, NUMCHARS(szNTPath)) &&
			0 == _wcsicmp(szNTPath, pszPath))
		{
			wcscat_s(szDrive, NUMCHARS(szDrive), L"\\");
			wcscat_s(szDrive, NUMCHARS(szDrive), pszSlash + 1);
			wcscpy_s(pszPath, nMax, szDrive);
			return S_OK;
		}
	}
	*pszSlash = cSave;
	return E_FAIL;
}

void CALLBACK HandleWinEvent(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
	LONG idObject, LONG idChild,
	DWORD dwEventThread, DWORD dwmsEventTime)
{
	IAccessible* pAcc = NULL;
	VARIANT varChild;
	HRESULT hr = AccessibleObjectFromEvent(hwnd, idObject, idChild, &pAcc, &varChild);
	if ((hr == S_OK) && (pAcc != NULL))
	{
		HWND ActiveWindow = GetForegroundWindow();
		HWND me = GetConsoleWindow();
		DWORD ProcessId = 0;
		DWORD capacity = 65535;
		LPTSTR ProcessName = new TCHAR[MAX_PATH];
		GetWindowThreadProcessId(ActiveWindow, &ProcessId);
		
		//QueryFullProcessImageName(ActiveWindow, PROCESS_NAME_NATIVE, ProcessName, &capacity);
		
		//GetProcessImageFileName(ActiveWindow, ProcessName, sizeof(ProcessName));

		HANDLE processHandle = NULL;
		wchar_t path[_MAX_PATH];
		
		processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, ProcessId);
		if (processHandle != NULL) {
			if (
				//GetModuleFileNameEx(processHandle, NULL, ProcessName, MAX_PATH)
				QueryFullProcessImageName(processHandle, PROCESS_NAME_NATIVE, ProcessName, &capacity)
				//GetProcessImageFileName(processHandle, ProcessName, sizeof(ProcessName))
				== 0) {
				tcerr << "Failed to get module filename." << std::endl;
			}
			else {
				tcout << std::endl << "FOCUS CHANGE:" << std::endl << std::endl;
				wcscpy_s(path, _MAX_PATH, ProcessName);
				HRESULT hr = NormalizeNTPath(path, _MAX_PATH);
				if (SUCCEEDED(hr)){
					//SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), FOREGROUND_RED | BACKGROUND_BLUE);
					std::wcout << path << std::endl;
				}
				tcout << "Module filename is: " << ProcessName << std::endl;
			}
			CloseHandle(processHandle);
		}
		else {
			tcerr << "Failed to open process." << std::endl;
		}

		//GetModuleFileNameEx(ActiveWindow, NULL, ProcessName, capacity);
		
		BSTR bstrName;
		
		pAcc->get_accName(varChild, &bstrName);
		
		if (event == EVENT_SYSTEM_MENUSTART)
		{
			printf("Begin: ");
		}
		else if (event == EVENT_SYSTEM_MENUEND)
		{
			printf("End:   ");
		}
		else if (event == EVENT_SYSTEM_FOREGROUND)
		{
			printf("Focus change.   ");
		}
		printf("%S\n", bstrName);
		std::cout << ProcessId << ":" << ProcessName << std::endl;
		
		SysFreeString(bstrName);
		
		pAcc->Release();

		tcout << std::endl << "_______________________" << std::endl;
		
	}
}

int _tmain(int argc, _TCHAR* argv[])
{
	InitializeMSAA();
	MSG msg;
	
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	
	ShutdownMSAA();
	return 0;
}

