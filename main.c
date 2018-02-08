#include <stdio.h>
#include <stdlib.h>
#include <windows.h>

#import "msxml4.dll"
using namespace MSXML2;

void dump_com_error(_com_error &e)
{
        printf("Error\n");
        printf("\a\tCode = %08lx\n", e.Error());
        printf("\a\tCode meaning = %s", e.ErrorMessage());
        _bstr_t bstrSource(e.Source());
        _bstr_t bstrDescription(e.Description());
        printf("\a\tSource = %s\n", (LPCSTR) bstrSource);
        printf("\a\tDescription = %s\n", (LPCSTR) bstrDescription);
}


int main(int argc, char* argv[])
{
        CoInitialize(NULL);
        
        
        try{
                IXMLDOMDocumentPtr pXMLDoc;
                HRESULT hr = pXMLDoc.CreateInstance(__uuidof(DOMDocument));
                
                pXMLDoc->async = false; // default - true,
                
                
                hr = pXMLDoc->load("stock.xml");
                
                if(hr!=VARIANT_TRUE)
                {
                        IXMLDOMParseErrorPtr  pError;
                        
                        pError = pXMLDoc->parseError;
                        _bstr_t parseError =_bstr_t("At line ")+ _bstr_t(pError->Getline()) +
_bstr_t("\n")+ _bstr_t(pError->Getreason());
                        MessageBox(NULL,parseError, "Parse Error",MB_OK);
                        return 0;
                }
                
                CComPtr<IStream> pStream;
                hr = CreateStreamOnHGlobal(NULL, true, &pStream);
                hr = pXMLDoc->save(pStream.p);
                
                LARGE_INTEGER pos;
                pos.QuadPart = 0;
                
                //the key is to reset the seek pointer
                pStream->Seek((LARGE_INTEGER)pos, STREAM_SEEK_SET, NULL);

                IXMLDOMDocumentPtr pXMLDocNew;
                hr = pXMLDocNew.CreateInstance(__uuidof(DOMDocument));
                pXMLDocNew->async = false;
                hr = pXMLDocNew->load(pStream.p);
                if(hr!=VARIANT_TRUE)
                {
                        IXMLDOMParseErrorPtr  pError;
                        
                        pError = pXMLDocNew->parseError;
                        _bstr_t parseError =_bstr_t("At line ")+ _bstr_t(pError->Getline()) +
_bstr_t("\n")+ _bstr_t(pError->Getreason());
                        MessageBox(NULL,parseError, "Parse Error",MB_OK);
                        return 0;
                }

                MessageBox(NULL,(LPTSTR)pXMLDocNew->xml, "XML content",MB_OK);          
                
        }
        catch(_com_error &e)
        {
                dump_com_error(e);
        }
      CoUninitialize();

        return 0;
}


