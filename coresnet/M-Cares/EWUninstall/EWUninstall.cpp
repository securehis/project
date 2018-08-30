// EWUninstall.cpp : 응용 프로그램에 대한 진입점을 정의합니다.
//

#include "stdafx.h"
#include "EWUninstall.h"


int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// TODO: 여기에 코드를 입력합니다.
	WCHAR szUninstallFileName[MAX_PATH];

	GetModuleFileName(NULL, szUninstallFileName, MAX_PATH);
	PathRemoveFileSpec(szUninstallFileName);
	PathAddBackslash(szUninstallFileName);
	wcscat_s(szUninstallFileName, MAX_PATH, L"Uninstall.exe");

	ShellExecute(NULL, L"open", szUninstallFileName, NULL, NULL, SW_SHOW);

	return 0;
}
