//============================================================================
// Name        : wctlcmd.cpp
// Author      : lovemch
// Version     :
// Copyright   : Everyzone Inc.
// Description : Whitelist control command
//============================================================================

#include <dirent.h>
#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/types.h>
#include <unistd.h>


enum
{
	CSY_SEND_ON = 19870914,
	CSY_SEND_OFF,
	CSY_SEND_CONTINUE,
	CSY_SEND_PAUSE,
};


bool SendSignal(int nSignalNumber);

pid_t GetDaemonPid();
pid_t GetPidNumber(char *szPid);

void ReloadWhitelistDatabase();
void WhitelistContinue();
void WhitelistOff();
void WhitelistOn();
void WhitelistPause();


int main(int nArgc, char **szArgv)
{
	if (nArgc != 2)
		return 0;

	if (strcasecmp(szArgv[1], "/wloff") == 0)
		WhitelistOff();
	else if (strcasecmp(szArgv[1], "/wlon") == 0)
		WhitelistOn();
	else if (strcasecmp(szArgv[1], "/wlpause") == 0)
		WhitelistPause();
	else if (strcasecmp(szArgv[1], "/wlcontinue") == 0)
		WhitelistContinue();
	else if (strcasecmp(szArgv[1], "/reloaddb") == 0)
		ReloadWhitelistDatabase();

	return 0;
}


//	bool
bool SendSignal(int nSignalNumber)
{
	pid_t nProcessId = 0;

	union sigval svWhitelistOff;

	nProcessId = GetDaemonPid();
	if (nProcessId == -1)
		return false;

	svWhitelistOff.sival_int = nSignalNumber;
	if (sigqueue(nProcessId, SIGRTMIN, svWhitelistOff) == -1)
		return false;
	else
		return true;
}


//	pid_t
pid_t GetDaemonPid()
{
	pid_t pid;

	struct dirent *pDirent = NULL;

	char szProcPath[260];
	char szPidFile[260];
	char szLine[1024];
	char szTag[100];
	char szProcessName[100];

	DIR *pDir = NULL;

	FILE *pFile = NULL;

	strcpy(szProcPath, "/proc");
	pDir = opendir(szProcPath);
	if (pDir == NULL)
		return -1;

	while ((pDirent = readdir(pDir)))
	{
		pid = GetPidNumber(pDirent->d_name);
		if (pid == -1)
			continue;

		snprintf(szPidFile, 260, "/proc/%d/status", pid);
		pFile = fopen(szPidFile, "r");
		if (pFile == NULL)
			continue;

		fgets(szLine, 1024, pFile);
		sscanf(szLine, "%s %s", szTag, szProcessName);
		if (strcasecmp(szProcessName, "wctld") == 0)
		{
			closedir(pDir);
			return pid;
		}
	}

	closedir(pDir);

	return -1;
}

pid_t GetPidNumber(char *szPid)
{
	size_t nLength, nCount;

	nLength = strlen(szPid);
	for (nCount = 0; nCount < nLength; nCount ++)
	{
		if ((szPid[nCount] < '0') || (szPid[nCount] > '9'))
			return -1;
	}

	return atoi(szPid);
}


//	void
void ReloadWhitelistDatabase()
{
	char szDaemonFile[260];
	char szPath[260];
	char szPid[260];

	char *pBuffer = NULL;

	pid_t pid = -1;

	snprintf(szPid, 260, "/proc/%d/exe", getpid());
	readlink((const char *)szPid, szPath, 260);
	pBuffer = strrchr(szPath, '/');
	if (pBuffer != NULL)
		*(pBuffer + 1) = '\0';
	snprintf(szDaemonFile, 260, "%sMakeAwlDB", szPath);

	pid = fork();
	switch (pid)
	{
		case -1:
			break;

		/* Child falls through... */
		case 0:
			execl(szDaemonFile, szDaemonFile, NULL);
			exit(0);
			break;

		/* while parent terminates */
		default:
			//return false;
			break;
    }
}

void WhitelistContinue()
{
	SendSignal(CSY_SEND_CONTINUE);
}

void WhitelistOff()
{
	SendSignal(CSY_SEND_OFF);
}

void WhitelistOn()
{
	SendSignal(CSY_SEND_ON);
}

void WhitelistPause()
{
	SendSignal(CSY_SEND_PAUSE);
}
