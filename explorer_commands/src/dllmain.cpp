#include "dllmain.hpp"
#include <cstdio>
#include <wrl.h>

HMODULE g_hinst;

#ifndef NDEBUG
void CreateConsole()
{
    AllocConsole();

    FILE* fp;
    freopen_s(&fp, "CONOUT$", "w", stdout);
    freopen_s(&fp, "CONOUT$", "w", stderr);
    freopen_s(&fp, "CONIN$", "r", stdin);

    SetConsoleTitleW(L"sublime_explorer_commands.dll");
}
#endif

STDAPI DllCanUnloadNow()
{
    if(Microsoft::WRL::Module<Microsoft::WRL::InProc>::GetModule().Terminate()){
        return S_OK;
    }
    return S_FALSE;
}

STDAPI DllGetActivationFactory(HSTRING activatable_clsid, IActivationFactory **factory)
{
    return Microsoft::WRL::Module<Microsoft::WRL::InProc>::GetModule().GetActivationFactory(activatable_clsid, factory);
}

STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void **v)
{
    return Microsoft::WRL::Module<Microsoft::WRL::InProc>::GetModule().GetClassObject(clsid, iid, v);
}

BOOL WINAPI DllMain(HINSTANCE hinst, DWORD reason, void *reserved)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hinst = hinst;

        #ifndef NDEBUG
        CreateConsole();
        #endif

        DisableThreadLibraryCalls(hinst);
    }
    return true;
}
