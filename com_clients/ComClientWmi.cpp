#define _WIN32_DCOM
#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>

using namespace std;

#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char** argv)
{
    HRESULT hr;

    // Initialize COM.
    hr = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        cout << "Failed to initialize COM library. "
            << "Error code = 0x"
            << hex << hr << endl;
        return 1;              // Program has failed.
    }

    // Initialize 
    hr = CoInitializeSecurity(
        NULL,
        -1,      // COM negotiates service                  
        NULL,    // Authentication services
        NULL,    // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
        RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
        NULL,             // Authentication info 
        EOAC_NONE,        // Additional capabilities
        NULL              // Reserved
    );


    if (FAILED(hr))
    {
        cout << "Failed to initialize security. "
            << "Error code = 0x"
            << hex << hr << endl;
        CoUninitialize();
        return 1;          // Program has failed.
    }

    // Obtain the initial locator to Windows Management
    // on a particular host computer.
    IWbemLocator* pLoc = 0;

    hr = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hr))
    {
        cout << "Failed to create IWbemLocator object. "
            << "Error code = 0x"
            << hex << hr << endl;
        CoUninitialize();
        return 1;       // Program has failed.
    }

    IWbemServices* pSvc = 0;

    // Connect to the root\cimv2 namespace with the
    // current user and obtain pointer pSvc
    // to make IWbemServices calls.

    hr = pLoc->ConnectServer(

        _bstr_t(L"ROOT\\CIMV2"), // WMI namespace
        NULL,                    // User name
        NULL,                    // User password
        0,                       // Locale
        NULL,                    // Security flags                 
        0,                       // Authority       
        0,                       // Context object
        &pSvc                    // IWbemServices proxy
    );

    if (FAILED(hr))
    {
        cout << "Could not connect. Error code = 0x"
            << hex << hr << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Set the IWbemServices proxy so that impersonation
    // of the user (client) occurs.
    hr = CoSetProxyBlanket(

        pSvc,                         // the proxy to set
        RPC_C_AUTHN_WINNT,            // authentication service
        RPC_C_AUTHZ_NONE,             // authorization service
        NULL,                         // Server principal name
        RPC_C_AUTHN_LEVEL_CALL,       // authentication level
        RPC_C_IMP_LEVEL_IMPERSONATE,  // impersonation level
        NULL,                         // client identity 
        EOAC_NONE                     // proxy capabilities     
    );

    if (FAILED(hr))
    {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hr << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }


    // Use the IWbemServices pointer to make requests of WMI. 
    // Make requests here:

    // For example, query for all the running processes
    IEnumWbemClassObject* pEnumerator = NULL;
    hr = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_Process"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hr))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hr << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;

        while (pEnumerator)
        {
            hr = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            wcout << "Process Name : " << vtProp.bstrVal << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            pclsObj = NULL;
        }

    }

    // Cleanup
    // ========

    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();

    CoUninitialize();

    return 0;  
}
