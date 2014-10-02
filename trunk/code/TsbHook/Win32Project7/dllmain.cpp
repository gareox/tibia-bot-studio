#define _CRT_SECURE_NO_DEPRECATE
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <winsock2.h>
#include <windows.h>
#include <iostream>
#include <fstream>
#include <string>
#include <shlobj.h>
#include <sstream>

using namespace std;
#pragma comment(lib,"ws2_32.lib")

typedef std::basic_string<TCHAR> tstring;
typedef std::basic_ofstream<TCHAR> tofstream;

typedef INT(WINAPI *realConnect)(SOCKET s, const struct sockaddr* name, int namelen);
typedef INT(WINAPI *SendPtr)(SOCKET sock, CONST CHAR* buf, INT len, INT flags);
typedef INT(WINAPI *RecvPtr)(SOCKET sock,  CHAR* buf, INT len, INT flags);

INT WINAPI OurSend(SOCKET sock, CONST CHAR* buf, INT len, INT flags);
INT WINAPI OurRecv(SOCKET sock,  CHAR* buf, INT len, INT flags);

VOID *Detour(BYTE *source, CONST BYTE *destination, CONST INT length);


SendPtr RealSend;
RecvPtr RealRecv;

HANDLE Console;
realConnect pConnect;
SOCKET Bot;
SOCKADDR_IN addr;

char *IpServer;
u_short Port;

void _stdcall InitConnect()
{
	//IpServer = ip;
	//Port = pt;

	
	//Socketpart
	WSADATA wsa;
	WSAStartup(MAKEWORD(2, 2), &wsa);

	Bot = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
	memset(&addr, 0, sizeof(SOCKADDR_IN)); // zuerst alles auf 0 setzten 
	addr.sin_family = AF_INET;
	addr.sin_port = htons(Port);
	addr.sin_addr.s_addr = inet_addr(IpServer);

	short status;
	status = connect(Bot, (SOCKADDR*)(&addr), sizeof(addr));
	if (status == SOCKET_ERROR)
	{
		MessageBox(NULL, TEXT("No se puede Conectar a TibiaStudio"), TEXT("Error"), MB_OK);
		exit(0);
	}

}

void _stdcall ReadFile()
{
	fstream fIn;
	

	fIn.open("c:/ip.txt", ios::in);

	if (fIn.is_open())
	{
		string s;
		char *tmpip = new char[s.length() + 1];
		char *tmp = new char[s.length() + 1];
		unsigned short tmpport;
		int i;
		i = 0;
		while (getline(fIn, s))
		{
			cout << s << endl;
			i += 1;

			if (i == 1)
			{
				std::strcpy(tmpip, s.c_str());
				IpServer = tmpip;
			}

			if (i == 2)
			{
				std::stringstream ss(s.c_str());
					ss >> tmpport;
					Port = tmpport;
					std::strcpy(tmp, s.c_str());

			}

			// Tokenize s here into columns (probably on spaces)
		}
		fIn.close();
		Sleep(100);
		InitConnect();

	}
	else
		MessageBox(NULL, TEXT("No se puede Conectar a TibiaStudio"), TEXT("Error"), MB_OK);
		CloseHandle(Console);

}


BOOL WINAPI my_connect(SOCKET s, const struct sockaddr* name, int namelen)
{
	WORD port = ntohs((*(WORD*)name->sa_data));
	sockaddr_in *sockaddr = (sockaddr_in*)name;
	sockaddr->sin_port = htons(Port);
	if (port != 80)
	{
		sockaddr->sin_addr.S_un.S_addr = inet_addr(IpServer);
	}
	return pConnect(s, name, namelen);
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved)
{

	if (reason == DLL_PROCESS_ATTACH)
	{
		 
		// Dla okienkowych aplikacji i osób nie lubiących OutputDebugString(), można jeszcze dodać logowanie pakietów do pliku
		/*
		AllocConsole();
		Console = GetStdHandle(STD_OUTPUT_HANDLE);
		CHAR buffer[256];
		*/
		HMODULE hWS32 = LoadLibraryA("ws2_32.dll");
		pConnect = (realConnect)GetProcAddress(hWS32, "connect");
		RealSend = (SendPtr)GetProcAddress(hWS32, "send");
		//RealRecv = (RecvPtr)GetProcAddress(hWS32, "recv");
		
		/*
		sprintf(buffer, "Adres send() w aplikacji = 0x%x\n", RealSend);
		WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
		sprintf(buffer, "Adres recv() w aplikacji = 0x%x\n", RealRecv);
		WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
		*/

		RealSend = (SendPtr)Detour((BYTE*)RealSend, (BYTE*)&OurSend, 5);
		//RealRecv = (RecvPtr)Detour((BYTE*)RealRecv,(BYTE*)&OurRecv, 5);
		
		/*
		sprintf(buffer, "Adres send() w bibliotece = 0x%x\n", RealSend);
		WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
		sprintf(buffer, "Adres recv() w bibliotece = 0x%x\n", RealRecv);
		WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
		*/
		ReadFile();
	}
	else if (reason == DLL_PROCESS_DETACH)
	{
		CloseHandle(Console);
		//FreeConsole();
	}

	return TRUE;
}

VOID *Detour(BYTE *source, CONST BYTE *destination, CONST INT length)
{
	DWORD back;
	BYTE *jmp = (BYTE*)malloc(length + 5);

	VirtualProtect(source, length, PAGE_READWRITE, &back);
	memcpy(jmp, source, length);
	jmp += length;

	jmp[0] = 0xE9;
	*(DWORD*)(jmp + 1) = (DWORD)(source + length - jmp) - 5;

	source[0] = 0xE9;
	*(DWORD*)(source + 1) = (DWORD)(destination - source) - 5;

	VirtualProtect(source, length, back, &back);

	return (jmp - length);
}

INT WINAPI OurSend(SOCKET sock, CONST CHAR* buf, INT len, INT flags)
{
	//CHAR buffer[256];

	//sprintf(buffer, "SEND -> %s\n", buf);
	//WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
	// OutputDebugString(buffer);
	
	RealSend(Bot, buf, len, flags);
	return RealSend(sock, buf, len, flags);
}

INT WINAPI OurRecv(SOCKET sock, CHAR* buf, INT len, INT flags)
{
	//CHAR buffer[256];

	//sprintf(buffer, "RECV -> %s\n", buf);
	//WriteConsole(Console, buffer, strlen(buffer), NULL, NULL);
	// OutputDebugString(buffer);
	
	//RealRecv(Bot, buf, len, flags);
	return RealRecv(sock, buf, len, flags);
}
