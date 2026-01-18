#include "OpenHere.hpp"
#include "dllmain.hpp"

#include <array>
#include <filesystem>
#include <ShlObj.h>
#include <Shlwapi.h>
#include <wil/stl.h>
#include <winrt/base.h>

#ifndef NDEBUG
#include <cwchar>
#define LOG(...) do { std::wprintf(__VA_ARGS__); std::putchar('\n'); std::fflush(stdout); } while(0)
#else
#define LOG(...)
#endif

static winrt::hstring GetSublimeTextPath() {
    LOG(L"GetSublimeTextPath");
    std::array<wchar_t, MAX_PATH> module_path;
    auto res = GetModuleFileNameW(g_hinst, module_path.data(), module_path.size());
    if(res == 0 || res == module_path.size()) {
        LOG(L"GetModuleFileNameW failed");
        winrt::throw_last_error();
    }
    std::filesystem::path sublime_path = std::filesystem::path(module_path.data()).parent_path().parent_path().append("sublime_text.exe");
    return sublime_path.c_str();
}

IFACEMETHODIMP OpenHere::Invoke(IShellItemArray* psiItemArray, IBindCtx*)
{   
    LOG(L"Invoke");
    try {
        auto location = GetSublimeTextPath();
        winrt::com_ptr<IShellItemArray> items;
        items.copy_from(psiItemArray);
        winrt::com_ptr<IShellItem> item = GetLocation(items);
        if (item) {

            wil::unique_cotaskmem_string name;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &name))) {
                wil::unique_process_information process_info;
                STARTUPINFOEXW startup_info{ 0 };
                startup_info.StartupInfo.cb = sizeof(startup_info);
                startup_info.StartupInfo.dwFlags = STARTF_USESHOWWINDOW;
                startup_info.StartupInfo.wShowWindow = SW_SHOWNORMAL;

                std::wstring cmd_line = L"-n \"" + std::wstring(name.get()) + L"\"";

                LOG(L"CreateProcessW");

                if (CreateProcessW(location.data(),
                    cmd_line.data(),
                    nullptr,
                    nullptr,
                    false,
                    EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT,
                    nullptr,
                    nullptr,
                    &startup_info.StartupInfo,
                    &process_info)) {
                    LOG(L"OK");
                    return S_OK;
                }
            }
        }
        return E_FAIL;
    } catch (...) {
        return E_FAIL;
    }
}

IFACEMETHODIMP OpenHere::GetToolTip(IShellItemArray*, LPWSTR* ppszInfoTip)
{
    LOG(L"GetToolTip");
    *ppszInfoTip = nullptr;
    return E_NOTIMPL;
}

IFACEMETHODIMP OpenHere::GetTitle(IShellItemArray*, LPWSTR* ppszName)
{
    LOG(L"GetTitle");
    const std::wstring resource = L"Open in Sublime Text";

    auto res = SHStrDupW(resource.c_str(), ppszName);
    return res;
}

IFACEMETHODIMP OpenHere::GetState(IShellItemArray* psiItemArray, BOOL , EXPCMDSTATE* pCmdState)
{
    LOG(L"GetState");
    winrt::com_ptr<IShellItemArray> items;
    items.copy_from(psiItemArray);
    winrt::com_ptr<IShellItem> item = GetLocation(items);

    *pCmdState = ECS_HIDDEN;
    if (item) {
        SFGAOF attributes;
        const bool is_file_system_folder = item->GetAttributes(SFGAO_FILESYSTEM | SFGAO_FOLDER, &attributes) == S_OK;
        const bool is_compressed = item->GetAttributes(SFGAO_STREAM, &attributes) == S_OK;

        if (is_file_system_folder && !is_compressed) {
            *pCmdState = ECS_ENABLED;
        }
    }
    return S_OK;
}

IFACEMETHODIMP OpenHere::GetIcon(IShellItemArray*, LPWSTR* ppszIcon)
{
    LOG(L"GetIcon");
    try {
        auto location = GetSublimeTextPath() + L",-103";
        auto res = SHStrDupW(location.c_str(), ppszIcon);
        return res;
    } catch (...) {
        *ppszIcon = nullptr;
        return E_FAIL;
    }
}

IFACEMETHODIMP OpenHere::GetFlags(EXPCMDFLAGS* pFlags)
{
    LOG(L"GetFlags");
    *pFlags = ECF_DEFAULT;
    return S_OK;
}

IFACEMETHODIMP OpenHere::GetCanonicalName(GUID* pguidCommandName)
{
    *pguidCommandName = __uuidof(decltype(*this));
    return S_OK;
}

IFACEMETHODIMP OpenHere::EnumSubCommands(IEnumExplorerCommand** ppEnum)
{
    LOG(L"EnumSubCommands");
    *ppEnum = nullptr;
    return E_NOTIMPL;
}

IFACEMETHODIMP OpenHere::SetSite(IUnknown* site) noexcept
{
    LOG(L"SetSite");
    this->site.copy_from(site);
    return S_OK;
}

IFACEMETHODIMP OpenHere::GetSite(REFIID iid, void** site) noexcept
{
    LOG(L"GetSite");
    if (this->site) {
        return this->site->QueryInterface(iid, site);
    }
    return E_FAIL;
}

winrt::com_ptr<IShellItem> OpenHere::GetLocation(winrt::com_ptr<IShellItemArray> item_array) {
    LOG(L"GetLocation");
    winrt::com_ptr<IShellItem> item;

    if (item_array) {
        DWORD count{};
        item_array->GetCount(&count);
        if (count) {
            item_array->GetItemAt(0, item.put());
        }
    }

    if (!item) {
        item = GetLocationFromSite();
    }
    return item;
}

winrt::com_ptr<IShellItem> OpenHere::GetLocationFromSite() {
    LOG(L"GetLocationFromSite");
    winrt::com_ptr<IShellItem> item;
    if (site) {
        winrt::com_ptr<IServiceProvider> service_provider;
        if (site.try_as(service_provider)) {
            winrt::com_ptr<IFolderView> folder_view;
            service_provider->QueryService<IFolderView>(SID_SFolderView, folder_view.put());
            if(folder_view){
                item.try_capture(folder_view, &IFolderView::GetFolder);
            }
        }
    }
    return item;
}
