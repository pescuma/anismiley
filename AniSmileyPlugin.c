#include <windows.h>
#include <newpluginapi.h>
#include <tchar.h>
#include <crtdbg.h>

#define ME_SYSTEM_MODULESLOADED "Miranda/System/ModulesLoaded"
#define MS_SYSTEM_GETVERSIONTEXT "Miranda/System/GetVersionText"

#define  InsertAnimatedSmiley
#include "m_anismiley.h"

// {76F66557-6DD8-4fc1-A872-0B6D93D800D5}
#define MIID_ANISMILEY { 0x76f66557, 0x6dd8, 0x4fc1, { 0xa8, 0x72, 0xb, 0x6d, 0x93, 0xd8, 0x0, 0xd5 } }
// {A3A5BF91-73DD-4853-8F42-3D47C88349B2}
#define MUUID_ANISMILEY { 0xa3a5bf91, 0x73dd, 0x4853, { 0x8f, 0x42, 0x3d, 0x47, 0xc8, 0x83, 0x49, 0xb2 } 

extern BOOL g_bGdiPlusFail;
extern HRESULT InitModule();
extern HRESULT UninitModule();
extern int InsertAniSmileyService(WPARAM wParam, LPARAM lParam);
DWORD  g_mirandaVersion=0;

HHOOK hModulesLoaded=NULL;

PLUGINLINK *pluginLink=NULL;
extern void *pLink;

PLUGININFOEX pluginInfo = {
        sizeof(PLUGININFOEX),
        "AniSmiley Service",
        PLUGIN_MAKE_VERSION(0, 1, 6, 0),
        "AniSmiley is service plugin for Miranda IM to provide ability for plugins to insert animated images into native SRMM dialogs.",
        "Artem Shpynov",
        "ashpynov@gmail.com",
        "© 2007-2008 Artem Shpynov, Miranda Project",
        "http://fyr.saddo.ru", //"addons.miranda-im.org/details.php?action=viewfile&id=3779",
        1,
        0, 
        MUUID_ANISMILEY}
};

__declspec(dllexport) PLUGININFO *MirandaPluginInfo(DWORD mirandaVersion)
{    
    pluginInfo.cbSize=sizeof(PLUGININFO);
    g_mirandaVersion =mirandaVersion;
    return (PLUGININFO*)&pluginInfo;
}

__declspec(dllexport) PLUGININFOEX *MirandaPluginInfoEx(DWORD mirandaVersion)
{    
    g_mirandaVersion=mirandaVersion;
    return &pluginInfo;
}
static const MUUID interfaces[] = {MIID_ANISMILEY, MIID_LAST};

__declspec(dllexport) const MUUID* MirandaPluginInterfaces(void)
{
    return interfaces;
}

BOOL CoreCheck()
{
    char version[1024], exepath[1024];
    GetModuleFileNameA(GetModuleHandle(NULL), exepath, sizeof(exepath));
    CallService(MS_SYSTEM_GETVERSIONTEXT, sizeof(version), (LPARAM)version);
    strlwr(version); strlwr(exepath);
    if (strstr(strrchr(exepath, '\\'), "coffee") ||
        strstr(version, "coffee") ||
        !strncmp(version, "1.", 2) || strstr(version, " 1.") ||
        (g_mirandaVersion >= PLUGIN_MAKE_VERSION(1,0,0,0)))
    {
        MessageBoxA(0,
            "AniSmiley plugin was designed to be used with Miranda IM ONLY.\n"
            "For use with any other application, please contact author: ashpynov@gmail.com.\n",
            "AniSmiley Error",
            MB_ICONSTOP|MB_OK);
        TerminateThread(GetCurrentThread(),0);
        return FALSE;
    }
    return TRUE;
}

int ModulesLoadedHook(WPARAM wParam, LPARAM lParam)
{
    if (!CoreCheck()) return 1;
	//CallService("Update/RegisterFL", (WPARAM)3779, (LPARAM)&pluginInfo);
	return 0;
}

int __declspec(dllexport) Load(PLUGINLINK * link)
{
#ifdef _DEBUG
	_CrtSetDbgFlag( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
#endif
#ifdef _DEBUG
	//_CrtSetBreakAlloc(775);
#endif

    InitModule();
    if (g_bGdiPlusFail) return 1;
    pluginLink = link;
	pLink=(void*)link;

	CreateServiceFunction(MS_INSERTANISMILEY, InsertAniSmileyService);    
	hModulesLoaded=HookEvent(ME_SYSTEM_MODULESLOADED, ModulesLoadedHook);
    return 0;
}

int __declspec(dllexport) Unload(void)
{
	UnhookEvent(hModulesLoaded);
    UninitModule();
    return 0;
}

