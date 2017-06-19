#include <array>
#include <cstdio>
#include <codecvt>
#include <vector>

#include <windows.h>
#include <devguid.h>
#include <Netcfgx.h>
#include <wrl/client.h>

using namespace Microsoft::WRL;


#define WIDE2(x) L##x
#define WIDEN(x) WIDE2(x)
#define WFILE WIDEN(__FILE__)

#define TRACE_HR(hr) \
  do { \
    if (FAILED(hr)) \
      printf("%s:%d: Failed: 0x%08X (%ls)\n", __FUNCTION__, __LINE__, hr, TranslateNetCfgHResult(hr)); \
  } while (false)

#define RETURN_IF_FAILED(expr) \
  do { \
    HRESULT hr_ = (expr); \
    if (FAILED(hr_)) { \
        TRACE_HR(hr_); \
        return (hr_); \
    } \
  } while (false)

#define RETURN_IF_FAILED_(ret, expr) \
  do { \
    HRESULT hr_ = (expr); \
    if (FAILED(hr_)) { \
        TRACE_HR(hr_); \
        return (ret); \
    } \
  } while (false)

#define CONTINUE_IF_FAILED(expr) \
  do { \
    HRESULT hr_ = (expr); \
    if (FAILED(hr_)) { \
        TRACE_HR(hr_); \
        continue; \
    } \
  } while (false)

namespace
{

struct ComInit
{
    ComInit() { CoInitialize(0); }
    ~ComInit() { CoUninitialize(); }
};


struct GuidAsString
{
    GuidAsString(GUID const& guid)
    {
        StringFromGUID2(guid, buffer.data(), (int)buffer.size());
    }

    wchar_t const* c_str() const { return buffer.data(); }

private:
    std::array<wchar_t, 38 + 1> buffer;
};


struct ReceiveCoTaskMemString
{
    ReceiveCoTaskMemString(std::wstring& str) : str(str) {}

    ~ReceiveCoTaskMemString()
    {
        if (ptr) {
            str.assign(ptr);
            CoTaskMemFree(ptr);
        }
    }

    operator LPWSTR*() { return &ptr; }

private:
    LPWSTR ptr = nullptr;
    std::wstring& str;
};


wchar_t const* TranslateNetCfgHResult(HRESULT hr)
{
#define ENTRY(value) \
    case value: return L ## #value

    switch (hr) {
    ENTRY(NETCFG_E_ALREADY_INITIALIZED);
    ENTRY(NETCFG_E_NOT_INITIALIZED);
    ENTRY(NETCFG_E_IN_USE);
    ENTRY(NETCFG_E_NO_WRITE_LOCK);
    ENTRY(NETCFG_E_NEED_REBOOT);
    ENTRY(NETCFG_E_ACTIVE_RAS_CONNECTIONS);
    ENTRY(NETCFG_E_ADAPTER_NOT_FOUND);
    ENTRY(NETCFG_E_COMPONENT_REMOVED_PENDING_REBOOT);
    ENTRY(NETCFG_E_MAX_FILTER_LIMIT);
    ENTRY(NETCFG_E_VMSWITCH_ACTIVE_OVER_ADAPTER);
    ENTRY(NETCFG_E_DUPLICATE_INSTANCEID);
    ENTRY(NETCFG_S_REBOOT);
    ENTRY(NETCFG_S_DISABLE_QUERY);
    ENTRY(NETCFG_S_STILL_REFERENCED);
    ENTRY(NETCFG_S_CAUSED_SETUP_CHANGE);
    ENTRY(NETCFG_S_COMMIT_NOW);
    default: return L"<unknown error>";
    }

#undef ENTRY
}

wchar_t const* TranslateNetCfgClass(GUID const& guid)
{
    if (guid == GUID_DEVCLASS_NET) return L"GUID_DEVCLASS_NET";
    if (guid == GUID_DEVCLASS_NETTRANS) return L"GUID_DEVCLASS_NETTRANS";
    if (guid == GUID_DEVCLASS_NETCLIENT) return L"GUID_DEVCLASS_NETCLIENT";
    if (guid == GUID_DEVCLASS_NETSERVICE) return L"GUID_DEVCLASS_NETSERVICE";
    return L"<unknown class>";
}


struct NetCfgInit
{
    NetCfgInit(INetCfg& netcfg) : netcfg(netcfg) {}

    HRESULT Initialize()
    {
        HRESULT hr = netcfg.Initialize(nullptr);
        initialized = SUCCEEDED(hr);
        return hr;
    }

    ~NetCfgInit()
    {
        if (initialized)
            netcfg.Uninitialize();
    }

private:
    INetCfg& netcfg;
    bool initialized = false;
};


struct NetCfgWriteLock
{
    NetCfgWriteLock(ComPtr<INetCfg>& netcfg, HRESULT& hr, wchar_t** lockingClient = nullptr)
    {
        hr = netcfg.As(&lock);
        if (!FAILED(hr))
            return;

        hr = lock->AcquireWriteLock(5000, L"PruneNetCfg", lockingClient);
        if (FAILED(hr)) {
            lock.Reset();
            return;
        }

        hr = S_OK;
    }

    ~NetCfgWriteLock()
    {
        if (lock)
            (void)lock->ReleaseWriteLock();
    }

    explicit operator bool() const { return lock != nullptr; }

private:
    ComPtr<INetCfgLock> lock;
};


HRESULT DumpComponent(INetCfgComponent& ncc)
{
    std::wstring id;
    std::wstring displayName;
    std::wstring bindName;
    ULONG status;
    GUID classGuid;
    GUID instanceGuid;
    RETURN_IF_FAILED(ncc.GetId(ReceiveCoTaskMemString(id)));
    RETURN_IF_FAILED(ncc.GetDisplayName(ReceiveCoTaskMemString(displayName)));
    RETURN_IF_FAILED(ncc.GetBindName(ReceiveCoTaskMemString(bindName)));
    RETURN_IF_FAILED(ncc.GetDeviceStatus(&status));
    RETURN_IF_FAILED(ncc.GetClassGuid(&classGuid));
    RETURN_IF_FAILED(ncc.GetInstanceGuid(&instanceGuid));

    printf("- Id:       %ls\n", id.c_str());
    printf("- Name:     %ls\n", displayName.c_str());
    printf("- BindName: %ls\n", bindName.c_str());
    printf("- Status:   %d\n", status);
    printf("- Class:    %ls (%ls)\n", GuidAsString(classGuid).c_str(), TranslateNetCfgClass(classGuid));
    printf("- Instance: %ls\n", GuidAsString(instanceGuid).c_str());
    printf("\n");
    return S_OK;
}


HRESULT EnumerateComponents(std::vector<std::wstring>* components, wchar_t const* filter = nullptr)
{
    HRESULT hr;

    ComPtr<INetCfg> nc;
    RETURN_IF_FAILED(CoCreateInstance(CLSID_CNetCfg, nullptr, CLSCTX_SERVER, IID_INetCfg, &nc));

    NetCfgInit init(*nc.Get());
    RETURN_IF_FAILED(init.Initialize());

    ComPtr<IEnumNetCfgComponent> enumNcc;
    RETURN_IF_FAILED(nc->EnumComponents(&GUID_DEVCLASS_NET, &enumNcc));
    RETURN_IF_FAILED(enumNcc->Reset());

    ComPtr<INetCfgComponent> ncc;
    while ((hr = enumNcc->Next(1, &ncc, nullptr)) == S_OK) {
        std::wstring id;
        std::wstring displayName;
        CONTINUE_IF_FAILED(ncc->GetId(ReceiveCoTaskMemString(id)));
        CONTINUE_IF_FAILED(ncc->GetDisplayName(ReceiveCoTaskMemString(displayName)));

        if (filter && displayName.find(filter) == std::wstring::npos)
            continue;

        DumpComponent(*ncc.Get());
        if (components) {
            while (true) {
                printf("Remove? [Y]es, [N]o: ");
                int c = fgetc(stdin);
                if (c == 'Y' || c == 'y') {
                    components->push_back(std::move(id));
                    while (c != EOF && c != '\n') c = fgetc(stdin);
                    break;
                } else if (c == 'N' || c == 'n') {
                    while (c != EOF && c != '\n') c = fgetc(stdin);
                    break;
                }
            }

            printf("\n");
        }
    }

    RETURN_IF_FAILED(nc->Uninitialize());
    return S_OK;
}


HRESULT EnumerateComponents2(wchar_t const* refComponent,
                             std::vector<std::wstring>* components,
                             wchar_t const* filter = nullptr)
{
    HRESULT hr;

    ComPtr<INetCfg> nc;
    RETURN_IF_FAILED(CoCreateInstance(CLSID_CNetCfg, nullptr, CLSCTX_SERVER, IID_INetCfg, &nc));

    NetCfgInit init(*nc.Get());
    RETURN_IF_FAILED(init.Initialize());

    ComPtr<INetCfgComponent> refNcc;
    RETURN_IF_FAILED(nc->FindComponent(refComponent, &refNcc));

    ComPtr<INetCfgComponentBindings> bindings;
    CONTINUE_IF_FAILED(refNcc->QueryInterface(IID_INetCfgComponentBindings, &bindings));
    
    ComPtr<IEnumNetCfgBindingPath> enumBp;
    CONTINUE_IF_FAILED(bindings->EnumBindingPaths(EBP_BELOW, &enumBp));
    CONTINUE_IF_FAILED(enumBp->Reset());
    
    ComPtr<INetCfgBindingPath> bp;
    while ((hr = enumBp->Next(1, &bp, nullptr)) == S_OK) {
        if (!bp->IsEnabled() == S_OK)
            continue;
    
        ComPtr<IEnumNetCfgBindingInterface> enumBi;
        CONTINUE_IF_FAILED(bp->EnumBindingInterfaces(&enumBi));
        CONTINUE_IF_FAILED(enumBi->Reset());
    
        ComPtr<INetCfgBindingInterface> bi;
        while ((hr = enumBi->Next(1, &bi, nullptr)) == S_OK) {
            ComPtr<INetCfgComponent> mpNcc;
            CONTINUE_IF_FAILED(bi->GetLowerComponent(&mpNcc));
    
            DumpComponent(*mpNcc.Get());
    
            std::wstring id;
            std::wstring displayName;
            CONTINUE_IF_FAILED(mpNcc->GetId(ReceiveCoTaskMemString(id)));
            CONTINUE_IF_FAILED(mpNcc->GetDisplayName(ReceiveCoTaskMemString(displayName)));
    
            if (filter && displayName.find(filter) == std::wstring::npos)
                continue;
    
            if (components) {
                while (true) {
                    printf("Remove? [Y]es, [N]o: ");
                    int c = fgetc(stdin);
                    if (c == 'Y' || c == 'y') {
                        components->push_back(std::move(id));
                        while (c != EOF && c != '\n') c = fgetc(stdin);
                        break;
                    } else if (c == 'N' || c == 'n') {
                        while (c != EOF && c != '\n') c = fgetc(stdin);
                        break;
                    }
                }
    
                printf("\n");
            }
        }
    }

    RETURN_IF_FAILED(nc->Uninitialize());
    return S_OK;
}

HRESULT DeinstallComponents(std::vector<std::wstring> const& componentIds)
{
    ComPtr<INetCfg> nc;
    RETURN_IF_FAILED_(1, CoCreateInstance(CLSID_CNetCfg, nullptr, CLSCTX_SERVER, IID_INetCfg, &nc));

    NetCfgInit init(*nc.Get());
    RETURN_IF_FAILED(init.Initialize());

    HRESULT hr;
    std::wstring lockingClient;

    NetCfgWriteLock lock(nc, hr, ReceiveCoTaskMemString(lockingClient));

    if (FAILED(hr)) {
        if (hr == NETCFG_E_NO_WRITE_LOCK)
            printf("NetCfg is already write-locked by %ls.\n", lockingClient.c_str());
        else
            printf("Failed to acquire NetCfg write lock (hr=0x%08X).\n", hr);
        return E_FAIL;
    }

    for (auto const& componentId : componentIds) {
        printf("Removing %ls\n", componentId.c_str());

        ComPtr<INetCfgComponent> ncc;
        CONTINUE_IF_FAILED(nc->FindComponent(componentId.c_str(), &ncc));

        GUID classGuid;
        CONTINUE_IF_FAILED(ncc->GetClassGuid(&classGuid));

        ComPtr<INetCfgClass> ncClass;
        ComPtr<INetCfgClassSetup> ncClassSetup;
        CONTINUE_IF_FAILED(nc->QueryNetCfgClass(&classGuid, IID_PPV_ARGS(&ncClass)));
        CONTINUE_IF_FAILED(ncClass.As(&ncClassSetup));
        CONTINUE_IF_FAILED(ncClassSetup->DeInstall(ncc.Get(), nullptr, nullptr));
    }

    RETURN_IF_FAILED_(1, nc->Apply());

    return S_OK;
}

} // namespace

int main(int argc, char** argv)
{
    std::wstring_convert<std::codecvt<wchar_t, char, std::mbstate_t>> conv;

    bool doDelete = argc >= 2 && strcmp(argv[1], "-d") == 0;
    std::wstring filter;
    if (doDelete && argc == 3)
        filter = conv.from_bytes(argv[2]);
    else if (argc == 2)
        filter = conv.from_bytes(argv[1]);

    ComInit comInit;

    std::vector<std::wstring> selectedComponents;
    RETURN_IF_FAILED_(1, EnumerateComponents(doDelete ? &selectedComponents : nullptr,
                                             filter.length() > 0 ? filter.c_str() : nullptr));

    if (doDelete) {
        HRESULT hr = DeinstallComponents(selectedComponents);
        if (FAILED(hr)) {
            printf("Failed to deinstall all components: hr=0x%08X\n", hr);
            return 1;
        }
    }

    return 0;
}
