#include <windows.h>
#include <stdio.h>
#include <comdef.h>
#include <atlbase.h>

#include "..\..\include\AxNetworkConstants.h"
#include "..\..\include\AxNetwork.h"
#include "..\..\include\AxNetwork_i.c"

LPTSTR ReadInput(LPCTSTR lpszTitle, BOOL bAllowEmpty = FALSE);
LPTSTR GetErrorDescription(LONG lLastError, IIcmp* pIcmp);


int main()
{
    IIcmp* pIcmp = NULL;
    HRESULT        hr;
    LONG        lLastError = -1L;
    LONG        lLastDuration = -1L, lLastTtl = -1L;
    LPTSTR        lptszHost = NULL;
    int          i;

    // Initialize COM  
    CoInitialize(NULL);

    // Create objects
    hr = CoCreateInstance(CLSID_Icmp, NULL, CLSCTX_INPROC_SERVER, IID_IIcmp, (void**)&pIcmp);
    if (!SUCCEEDED(hr))
    {
        _tprintf(_T("Unable to create Icmp object.\n"));
        goto _EndMain;
    }

    lptszHost = ReadInput(_T("Enter host: "));
    pIcmp->put_Ttl(255);
    for (i = 0; i < 4; i++)
    {
        pIcmp->Ping(_bstr_t(lptszHost), 5000);  // 5000: timeout in msecs
        pIcmp->get_LastError(&lLastError);
        if (lLastError != 0)
        {
            _tprintf(_T("Ping, result: %ld (%s)\n"), lLastError, GetErrorDescription(lLastError, pIcmp));
        }
        else
        {
            pIcmp->get_LastDuration(&lLastDuration);
            pIcmp->get_LastTtl(&lLastTtl);
            _tprintf(_T("Reply from %s; time=%ld TTL=%ld\n"), lptszHost, lLastDuration, lLastTtl);
        }
        Sleep(1000);
    }

    _tprintf(_T("\n"));

_EndMain:

    if (pIcmp != NULL)
        pIcmp->Release();

    CoUninitialize();

    _tprintf(_T("Ready.\n"));

    return 0;
}



///////////////////////////////////////////////////////////////////////////////////////////

LPTSTR GetErrorDescription(LONG lLastError, IIcmp* pIcmp)
{
    static TCHAR  tszErrorDescription[1024 + 1] = { _T('\0') };
    BSTR      bstrErrDescr = NULL;

    pIcmp->GetErrorDescription(lLastError, &bstrErrDescr);
    if (bstrErrDescr != NULL)
    {
        _stprintf_s(tszErrorDescription, 1024, _T("%ls"), bstrErrDescr);
        SysFreeString(bstrErrDescr);

    }
    return tszErrorDescription;
}



///////////////////////////////////////////////////////////////////////////////////////////

LPTSTR ReadInput(LPCTSTR lptszTitle, BOOL bAllowEmpty)
{
    static TCHAR    tszInput[255 + 1] = { _T('\0') };

    _tprintf(_T("%s:\n"), lptszTitle);
    do
    {
        _tprintf(_T("   > "));
        fflush(stdin);
        fflush(stdout);
        _fgetts(tszInput, 255, stdin);
        if (tszInput[0] != _T('\0') && tszInput[_tcsclen(tszInput) - 1] == _T('\n'))
            tszInput[_tcsclen(tszInput) - 1] = _T('\0');
    } while (_tcsclen(tszInput) == 0 && !bAllowEmpty);

    return tszInput;
}