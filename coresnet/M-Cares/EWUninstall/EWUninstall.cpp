// EWUninstall.cpp : ���� ���α׷��� ���� �������� �����մϴ�.
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

	// TODO: ���⿡ �ڵ带 �Է��մϴ�.
	WCHAR szUninstallFileName[MAX_PATH];

	GetModuleFileName(NULL, szUninstallFileName, MAX_PATH);
	PathRemoveFileSpec(szUninstallFileName);
	PathAddBackslash(szUninstallFileName);
	wcscat_s(szUninstallFileName, MAX_PATH, L"Uninstall.exe");

	ShellExecute(NULL, L"open", szUninstallFileName, NULL, NULL, SW_SHOW);

	return 0;
}
