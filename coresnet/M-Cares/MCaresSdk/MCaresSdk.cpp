// MCaresSdk.cpp : DLL 응용 프로그램을 위해 내보낸 함수를 정의합니다.
//

#include "stdafx.h"


//	정의
#define CSY_PARAM_NAME_UPDATE_AUTO										L"-UpdateParam:Auto"

#define CSY_PARAM_NAME_UPDATE_HIDE										L"-UpdateParam:Hide"

#define CSY_PARAM_NAME_UPDATE_SHOW										L"-UpdateParam:Show"


//	외부 사용 함수
extern "C" __declspec(dllexport) BOOL BeingMCares();
extern "C" __declspec(dllexport) BOOL Init();
extern "C" __declspec(dllexport) BOOL StartFolder();
extern "C" __declspec(dllexport) BOOL StartMCares();
extern "C" __declspec(dllexport) BOOL StartWhitelisting();
extern "C" __declspec(dllexport) BOOL StopFolder();
extern "C" __declspec(dllexport) BOOL StopMCares();
extern "C" __declspec(dllexport) BOOL StopWhitelisting();

extern "C" __declspec(dllexport) INT GetFolderStatus();
extern "C" __declspec(dllexport) INT GetStatus();
extern "C" __declspec(dllexport) INT GetWhitelistingStatus();

extern "C" __declspec(dllexport) VOID ExecuteMain();
extern "C" __declspec(dllexport) VOID ExecuteUpdate(INT nUpdateType);
extern "C" __declspec(dllexport) VOID Uninit();


//	내부 사용 함수
BOOL IsVistaOrHigher();
BOOL ReadSettings(CSYSETTINGS &Settings);
BOOL RegistryQueryValue(HKEY hKey, PWCHAR szKey, PWCHAR szValue, PBYTE pData, DWORD &dwType, DWORD &dwSize);
BOOL WriteMailslot(PWCHAR szMailslotName, PVOID pData, ULONG lLength);

VOID ReadDefaultSettings(CSYSETTINGS &Settings);


//	외부 사용 함수
//	BOOL
BOOL BeingMCares()
{
	HANDLE hMutex = NULL;

	hMutex = OpenMutex(MUTEX_ALL_ACCESS, 0, CSY_MUTEX_NAME_SERVICE);
	if (hMutex == NULL)
		return FALSE;
	else
	{
		CloseHandle(hMutex);
		return TRUE;
	}
}

BOOL Init()
{
	DWORD dwSize = 0;
	DWORD dwType = 0;

	WCHAR szFileName[MAX_PATH];

	dwSize = MAX_PATH * sizeof(WCHAR);
	dwType = REG_SZ;
	if (RegistryQueryValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\M-Cares", L"Path", (PBYTE)szFileName, dwType, dwSize) != TRUE)
		wcscpy_s(szFileName, MAX_PATH, L"C:\\Sense\\M-Cares");
	PathAddBackslash(szFileName);
	wcscat_s(szFileName, MAX_PATH, L"EWBlock.exe");

	ShellExecute(NULL, L"open", szFileName, NULL, NULL, SW_SHOW);

	return TRUE;
}

BOOL StartFolder()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_CC_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_START;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}

BOOL StartMCares()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_START;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}

BOOL StartWhitelisting()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_WL_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_START;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}

BOOL StopFolder()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_CC_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_STOP;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}

BOOL StopMCares()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_STOP;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}

BOOL StopWhitelisting()
{
	BOOL bReturn = FALSE;

	CSYMAILSLOTDATA MailslotData;

	MailslotData.Header = CSY_MAILSLOT_HEADER_WL_MODE_TRAY;
	MailslotData.Mode = CSY_MODE_STOP;
	bReturn = WriteMailslot(CSY_MAILSLOT_NAME_EW_SERVICE, &MailslotData, CSYMAILSLOTDATA_SIZE);
	WriteMailslot(CSY_MAILSLOT_NAME_EW_MAIN, &MailslotData, CSYMAILSLOTDATA_SIZE);
	return bReturn;
}


//	INT
INT GetFolderStatus()
{
	CSYSETTINGS Settings;

	if (BeingMCares() != TRUE)
		return -2;

	if (ReadSettings(Settings) != TRUE)
		return -1;

	if (Settings.CCMode == CSY_MODE_START)
		return 1;
	else
		return 0;
}

INT GetStatus()
{
	CSYSETTINGS Settings;

	if (BeingMCares() != TRUE)
		return -2;

	if (ReadSettings(Settings) != TRUE)
		return -1;

	if (Settings.Mode == CSY_MODE_START)
		return 1;
	else
		return 0;
}

INT GetWhitelistingStatus()
{
	CSYSETTINGS Settings;

	if (BeingMCares() != TRUE)
		return -2;

	if (ReadSettings(Settings) != TRUE)
		return -1;

	if (Settings.WLMode == CSY_MODE_START)
		return 1;
	else
		return 0;
}


//	VOID
VOID ExecuteMain()
{
	DWORD dwSize = 0;
	DWORD dwType = REG_SZ;

	HANDLE hEvent = NULL;
	HANDLE hMapping = NULL;

	LPVOID lpViewOf = NULL;

	SECURITY_ATTRIBUTES SecAttr;

	ULONG lCompare[2] = { 0, 0 };

	WCHAR szInstallFolder[MAX_PATH];
	WCHAR szMainFileName[MAX_PATH];
	WCHAR szParam[MAX_PATH];

	if (IsVistaOrHigher() == TRUE)
	{
		BOOL fSaclPresent = FALSE;
		BOOL fSaclDefaulted = FALSE;

		PACL pSacl = NULL;

		PSECURITY_DESCRIPTOR pSD;

		WCHAR szSecDesc[SECURITY_DESCRIPTOR_MIN_LENGTH];

		SecAttr.nLength = sizeof(SecAttr);
		SecAttr.bInheritHandle = FALSE;
		SecAttr.lpSecurityDescriptor = &szSecDesc;

		InitializeSecurityDescriptor(SecAttr.lpSecurityDescriptor, SECURITY_DESCRIPTOR_REVISION);
		SetSecurityDescriptorDacl(SecAttr.lpSecurityDescriptor, TRUE, 0, FALSE);

		ConvertStringSecurityDescriptorToSecurityDescriptorW(L"S:(ML;;NW;;;LW)", SDDL_REVISION_1, &pSD, NULL);
		GetSecurityDescriptorSacl(pSD, &fSaclPresent, &pSacl, &fSaclDefaulted);
		SetSecurityDescriptorSacl(SecAttr.lpSecurityDescriptor, TRUE, pSacl, FALSE);
	}
	else
	{
		SecAttr.nLength = sizeof(SecAttr);
		SecAttr.bInheritHandle = FALSE;
		SecAttr.lpSecurityDescriptor = NULL;

		::ConvertStringSecurityDescriptorToSecurityDescriptorW(L"O:SYG:AUD:(A;OICI;GA;;;WD)(A;OICI;GA;;;AU)", SDDL_REVISION_1, &(SecAttr.lpSecurityDescriptor), NULL);
	}

	dwSize = MAX_PATH * sizeof(WCHAR);
	dwType = REG_SZ;
	if (RegistryQueryValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\M-Cares", L"Path", (PBYTE)szInstallFolder, dwType, dwSize) != TRUE)
		wcscpy_s(szInstallFolder, MAX_PATH, L"C:\\Sense\\M-Cares");
	PathAddBackslash(szInstallFolder);

	lCompare[0] = 19741009;
	lCompare[1] = 19870914;

	hEvent = CreateEvent(&SecAttr, FALSE, FALSE, L"Global\\QUICKEVENT");
	hMapping = CreateFileMapping((HANDLE)INVALID_HANDLE_VALUE, &SecAttr, PAGE_READWRITE, 0, 13, L"Local\\QUICKLAUNCH");
	lpViewOf = MapViewOfFile(hMapping, FILE_MAP_READ | FILE_MAP_WRITE, 0, 0, 13);

	CopyMemory(lpViewOf, lCompare, sizeof(ULONG) * 2);

	swprintf_s(szMainFileName, MAX_PATH, L"%s%s", szInstallFolder, CSY_FILE_NAME_EW_MAIN_EXE);
	wcscpy_s(szParam, MAX_PATH, L"/quick");
	ShellExecute(NULL, L"open", szMainFileName, szParam, NULL, SW_SHOW);
	if (hEvent != NULL)
	{
		WaitForSingleObject(hEvent, 3000);
		CloseHandle(hEvent);
	}
	hEvent = NULL;

	if (lpViewOf != NULL)
		UnmapViewOfFile(lpViewOf);

	if (hMapping != NULL)
		CloseHandle(hMapping);
	hMapping = NULL;
}

VOID ExecuteUpdate(INT nUpdateType)
{
	DWORD dwSize = 0;
	DWORD dwType = REG_SZ;

	WCHAR szParam[MAX_PATH];
	WCHAR szUpdateFileName[MAX_PATH];

	dwSize = MAX_PATH * sizeof(WCHAR);
	dwType = REG_SZ;
	ZeroMemory(szUpdateFileName, MAX_PATH * sizeof(WCHAR));
	if (RegistryQueryValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\M-Cares", L"Path", (PBYTE)szUpdateFileName, dwType, dwSize) != TRUE)
		wcscpy_s(szUpdateFileName, MAX_PATH, L"C:\\Sense\\M-Cares");
	PathAddBackslash(szUpdateFileName);
	wcscat_s(szUpdateFileName, MAX_PATH, L"EWUpdate.exe");
	if (nUpdateType == 1)
		wcscpy_s(szParam, MAX_PATH, CSY_PARAM_NAME_UPDATE_AUTO);
	else if (nUpdateType == 2)
		wcscpy_s(szParam, MAX_PATH, CSY_PARAM_NAME_UPDATE_HIDE);
	else
		wcscpy_s(szParam, MAX_PATH, CSY_PARAM_NAME_UPDATE_SHOW);
	ShellExecute(NULL, L"open", szUpdateFileName, szParam, NULL, SW_SHOW);
}

VOID Uninit()
{
}


//	내부사용 함수
//	BOOL
BOOL IsVistaOrHigher()
{
	OSVERSIONINFO osvi;

	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&osvi);
	if (osvi.dwMajorVersion >= 6)
		return TRUE;
	else
		return FALSE;
}

BOOL ReadSettings(CSYSETTINGS &Settings)
{
	errno_t nError = -1;

	size_t nRead = 0;

	BYTE bBuffer[CSYSETTINGS_SIZE * 4];
	BYTE bDecryption[CSYSETTINGS_SIZE * 4];

	DWORD dwSize = 0;
	DWORD dwType = REG_SZ;

	FILE *pFile = NULL;

	LEA_KEY LeaKey;

	WCHAR szSettingsFileName[MAX_PATH];

	dwSize = MAX_PATH * sizeof(WCHAR);
	dwType = REG_SZ;
	if (RegistryQueryValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\M-Cares", L"Path", (PBYTE)szSettingsFileName, dwType, dwSize) != TRUE)
		wcscpy_s(szSettingsFileName, MAX_PATH, L"C:\\Sense\\M-Cares");
	PathAddBackslash(szSettingsFileName);
	wcscat_s(szSettingsFileName, MAX_PATH, CSY_FILE_NAME_EW_SETTINGS_DAT);

	ReadDefaultSettings(Settings);
	nError = _wfopen_s(&pFile, szSettingsFileName, L"rb");
	if (nError != 0)
		return FALSE;

	nRead = fread(bBuffer, 1, CSYSETTINGS_SIZE * 4, pFile);
	fclose(pFile);

	lea_set_key(&LeaKey, EW_KEY, 32);
	lea_cbc_dec(bDecryption, bBuffer, (INT)nRead, EW_IV, &LeaKey);

	CopyMemory(&Settings, bDecryption, CSYSETTINGS_SIZE);

	return TRUE;
}

BOOL RegistryQueryValue(HKEY hKey, PWCHAR szKey, PWCHAR szValue, PBYTE pData, DWORD &dwType, DWORD &dwSize)
{
	HKEY hKey2;

	if (RegOpenKeyEx(hKey, szKey, 0, KEY_ALL_ACCESS, &hKey2) != ERROR_SUCCESS)
		return FALSE;

	if (RegQueryValueEx(hKey2, szValue, 0, &dwType, pData, &dwSize) != ERROR_SUCCESS)
	{
		RegCloseKey(hKey2);
		return FALSE;
	}

	RegCloseKey(hKey2);
	return TRUE;
}

BOOL WriteMailslot(PWCHAR szMailslotName, PVOID pData, ULONG lLength)
{
	BOOL bRet = FALSE;

	DWORD dwReturn = 0;

	HANDLE hMailslot = INVALID_HANDLE_VALUE;

	hMailslot = CreateFile(szMailslotName, GENERIC_WRITE, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hMailslot == INVALID_HANDLE_VALUE)
		return bRet;
	bRet = WriteFile(hMailslot, pData, lLength, &dwReturn, NULL);
	CloseHandle(hMailslot);

	return bRet;
}


//	VOID
VOID ReadDefaultSettings(CSYSETTINGS &Settings)
{
	Settings.Mode = CSY_MODE_START;
	Settings.WLMode = CSY_MODE_START;
	Settings.CCMode = CSY_MODE_START;
	Settings.Blacklist = CSY_BLACKLIST_SCAN;
	Settings.OddUsb = CSY_ODDUSB_ALLOW;
	Settings.Mbr = CSY_MBR_BLOCK;
}
