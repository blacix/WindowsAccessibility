// InspectDemo.cpp : Defines the entry point for the console application.

#include "stdafx.h"
#include <windows.h>
#include <stdio.h>
#include <UIAutomation.h>
#include <string>
#include <Oleacc.h>
#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <thread>
#include <sstream> 
#include <fcntl.h>
#include <io.h>
#include <psapi.h> // For access to GetModuleFileNameEx
#include <algorithm>
#include <atlbase.h>
#include <atlsafe.h>

class FocusChangedEventHandler :
	public IUIAutomationFocusChangedEventHandler
{
private:
	LONG _refCount; // for IUKNown
	IUIAutomation* pAutomation = NULL;

	std::wstring currentTitle;
	std::wstring currentUrl;
	std::vector<std::wstring> outputCache;

public:

	FocusChangedEventHandler(IUIAutomation* pui) :
		pAutomation(pui),
		_refCount(1)
	{
	}


	//IUnknown methods.
	ULONG STDMETHODCALLTYPE AddRef()
	{
		ULONG ret = InterlockedIncrement(&_refCount);
		return ret;
	}


	ULONG STDMETHODCALLTYPE Release()
	{
		ULONG ret = InterlockedDecrement(&_refCount);
		if (ret == 0)
		{
			delete this;
			return 0;
		}
		return ret;
	}


	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
	{
		if (riid == __uuidof(IUnknown))
			*ppInterface = static_cast<IUIAutomationFocusChangedEventHandler*>(this);
		else if (riid == __uuidof(IUIAutomationFocusChangedEventHandler))
			*ppInterface = static_cast<IUIAutomationFocusChangedEventHandler*>(this);
		else
		{
			*ppInterface = NULL;
			return E_NOINTERFACE;
		}
		this->AddRef();
		return S_OK;
	}


	// IUIAutomationFocusChangedEventHandler
	//called on a different thread
	HRESULT STDMETHODCALLTYPE HandleFocusChangedEvent(IUIAutomationElement * pSender)
	{
		HRESULT hr;

		IUIAutomationCondition* condition = createConditions();

		// process the element that sent focus changed
		outputCache.push_back(L"[source element]");
		std::wstring strElement = processElement(pSender, true);
		outputCache.push_back(strElement);

		//// ez kb vicc:
		//// msdn-en ott van, hogy a TreeScope_Ancestors not supported. akkor a faszé van az apiban. ^^
		//// nem is mûködik
		//IUIAutomationElementArray* ancestorsArray;
		//pSender->FindAll(TreeScope_Ancestors, condition, &ancestorsArray);
		//processArray(ancestorsArray);
		////ezért saját implementáció: FocusChangedEventHandler::findAllAncestors

		// collect and process ancestor elements in the tree
		outputCache.push_back(L"[ancestors]");
		std::vector<IUIAutomationElement*> ancestors = findAllAncestors(pSender);
		processArray(ancestors);

		// collect and process descendants elements in the tree matching condition
		outputCache.push_back(L"[descendants]");
		IUIAutomationElementArray* descendantsArray;
		pSender->FindAll(TreeScope_Descendants, condition, &descendantsArray);
		processArray(descendantsArray);
		

		// print result
		// flush output cache
		printOutput();

		// release what you want
		release({ descendantsArray, /*ancestorsArray, */ condition,/* titleCondition, docCondition, , tmpCondition*/});

		std::wcout << "--------------------------- event processed --------------------------" << std::endl;		
		return S_OK;
	}


	// own methods

	IUIAutomationCondition * createConditions()
	{
		// the old way...
		// create conditions
		// Levente: check controltypes
		// TODO: move to a method
		// TODO: user pAutomation->CreateAndConditionFromArray()
		//IUIAutomationCondition *docCondition;		
		//VARIANT docVarProp;
		//docVarProp.vt = VT_I4;
		//docVarProp.lVal = UIA_DocumentControlTypeId;
		//hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, docVarProp, &docCondition);

		//IUIAutomationCondition *titleCondition;
		//VARIANT titleVarProp;
		//titleVarProp.vt = VT_I4;
		//titleVarProp.lVal = UIA_TitleBarControlTypeId;
		//hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, titleVarProp, &titleCondition);

		//IUIAutomationCondition *winCondition;
		//VARIANT winVarProp;
		//winVarProp.vt = VT_I4;
		//winVarProp.lVal = UIA_WindowControlTypeId;
		//hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, winVarProp, &winCondition);

		//IUIAutomationCondition* tmpCondition = NULL;
		//pAutomation->CreateOrCondition(docCondition, titleCondition, &tmpCondition);
		//IUIAutomationCondition* condition = NULL;
		//pAutomation->CreateOrCondition(tmpCondition, winCondition, &condition);

		HRESULT hr;
		IUIAutomationCondition* condition = NULL;
		IUIAutomationCondition *docCondition;
		VARIANT docVarProp;
		docVarProp.vt = VT_I4;
		docVarProp.lVal = UIA_DocumentControlTypeId;
		hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, docVarProp, &docCondition);

		IUIAutomationCondition *titleCondition;
		VARIANT titleVarProp;
		titleVarProp.vt = VT_I4;
		titleVarProp.lVal = UIA_TitleBarControlTypeId;
		hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, titleVarProp, &titleCondition);

		IUIAutomationCondition *winCondition;
		VARIANT winVarProp;
		winVarProp.vt = VT_I4;
		winVarProp.lVal = UIA_WindowControlTypeId;
		hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, winVarProp, &winCondition);

		SAFEARRAY* conditions = SafeArrayCreateVector( VT_UNKNOWN, 0, 3 );

		SafeArrayPutElement(conditions, 0, docCondition);
		SafeArrayPutElement(conditions, 0, titleCondition);
		SafeArrayPutElement(conditions, 0, winCondition);

		pAutomation->CreateOrConditionFromArray(conditions, &condition);
		SafeArrayDestroy(conditions);
		release({ docCondition , titleCondition , winCondition });
		return condition;
	}

	void processArray(IUIAutomationElementArray* array)
	{
		if (array == NULL)
			return;

		IUIAutomationElement* e;
		int l = 0;
		array->get_Length(&l);
		for (int i = 0; i < l; ++i)
		{
			array->GetElement(i, &e);
			outputCache.push_back(
				processElement(e)
			);
		}
	}


	// TODO use IUIAutomationElementArray?
	void processArray(std::vector<IUIAutomationElement*> elements)
	{
		for (auto it = elements.begin(); it != elements.end(); it++)
		{
			outputCache.push_back(
				processElement(*it)
			);
		}
	};


	// creates json string of an event for taskfit
	static std::wstring processElement(IUIAutomationElement* element, bool showOffscreen = false)
	{
		std::wstringstream stream;
		if (element == NULL)
			return stream.str();

		HRESULT hr;
		BSTR bstr;
		VARIANT variantValue;
		std::wstring name;
		std::wstring value;
		
		BOOL offScreen = TRUE;
		
		hr = element->get_CurrentIsOffscreen(&offScreen);
		if (FAILED(hr))
			return stream.str();

		// Levente: check this block (offscreen, lentgths)
		if (!offScreen || showOffscreen)
		{
			// get name
			element->get_CurrentName(&bstr);
			if (bstr != NULL)
			{
				std::wstring receivedName(bstr);
				if (receivedName.length() < 300)
					name.append(receivedName);
					//name.append(std::to_wstring(receivedName.length()));
			}

			//get value
			hr = element->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &variantValue);
			if (SUCCEEDED(hr))
			{
				std::wstring receivedValue(variantValue.bstrVal);
				if(receivedValue.length () < 300)
					value.append(receivedValue);
			}

			// get pid
			int pid = 0;
			std::wstring path;
			hr = element->get_CurrentProcessId(&pid);			
			if (SUCCEEDED(hr))
				path = getProcessPath(pid);

			stream << L"name: " << name << std::endl;
			stream << L"value: " << value << std::endl;
			stream << L"pid: " << pid << std::endl;
			stream << L"path : " << getProcessPath(pid) << std::endl;
		}
		
		return stream.str();
	}


	// TODO try to use IUIAutomationElementArray
	// TODO conditions(s) as argument (CONTROLTYPEID, eg UIA_DocumentControlTypeId)
	std::vector<IUIAutomationElement*> findAllAncestors(IUIAutomationElement* pSource)
	{
		std::vector<IUIAutomationElement*> elements;
		if (pSource == NULL)
			return elements;

		IUIAutomationElement* pDesktop = NULL;
		HRESULT hr = pAutomation->GetRootElement(&pDesktop);
		if (FAILED(hr))
			return elements;

		// Levent: nem kellene figyelni azt is,mikor kap a desktop fókuszt?
		// just to make sure, check if this is the desktop
		BOOL same;
		pAutomation->CompareElements(pSource, pDesktop, &same);
		if (same)
		{
			pDesktop->Release();
			return elements;
		}

		IUIAutomationElement* pParent = NULL;
		IUIAutomationElement* pCurrentNode = pSource;

		// Create the treewalker.
		IUIAutomationTreeWalker* pWalker = NULL;
		pAutomation->get_ControlViewWalker(&pWalker);
		if (pWalker == NULL)
			return elements;


		// walk up the tree
		do
		{
			// TODO make these arguments
			// Levente: check controltypes
			CONTROLTYPEID currentControlType;
			pCurrentNode->get_CurrentControlType(&currentControlType);
			if (currentControlType == UIA_DocumentControlTypeId ||
				currentControlType == UIA_TitleBarControlTypeId ||
				currentControlType == UIA_WindowControlTypeId)
			{
				elements.push_back(pCurrentNode);
			}

			// move up
			hr = pWalker->GetParentElement(pCurrentNode, &pParent);
			if (FAILED(hr) || pParent == NULL)
			{
				break;
			}
			pCurrentNode = pParent;
			// check if root is reached
			BOOL same;
			pAutomation->CompareElements(pCurrentNode, pDesktop, &same);
		} while (!same);
		release({ pDesktop , pWalker });
		return elements;
	}


	// prints and flushes output cache
	void printOutput()
	{
		std::vector<std::wstring> output(outputCache);
		for (auto it = output.begin(); it != output.end(); ++it)
		{
			std::wcout << (*it) << std::endl;
		}
		outputCache.clear();
	}

	// releases every instance of IUnknown from argument (e.g.: IUIAutomationElement, etc)
	static void release(std::vector<IUnknown*> releaseMe) 
	{
		for (auto it = releaseMe.begin(); it != releaseMe.end(); it++)
		{
			if (*it != NULL)
				(*it)->Release();
		}
	}


	static std::wstring getProcessPath(int pid)
	{
		HANDLE processHandle = NULL;
		TCHAR filename[MAX_PATH];

		processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, pid);
		if (processHandle != NULL) {
			if (GetModuleFileNameEx(processHandle, NULL, filename, MAX_PATH) == 0) {
				//tcerr << "Failed to get module filename." << endl;
			}
			else {
				//tcout << "Module filename is: " << filename << endl;
			}
			CloseHandle(processHandle);
		}
		else {
			//tcerr << "Failed to open process." << endl;
		}

		return std::wstring(&filename[0]);
	}
};




int main(int argc, char* argv[])
{
	// std::cout/wprintf won't fail on certain unicode charactrs
	_setmode(_fileno(stdout), _O_U16TEXT);

	HRESULT hr;
	int ret = 0;

	FocusChangedEventHandler* pEventHandler = NULL;

	CoInitializeEx(NULL, COINIT_MULTITHREADED);
	IUIAutomation* pAutomation = NULL;
	hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&pAutomation);
	if (FAILED(hr) || pAutomation == NULL)
	{
		ret = 1;
		goto cleanup;
	}

	pEventHandler = new FocusChangedEventHandler(pAutomation);
	if (pEventHandler == NULL)
	{
		ret = 1;
		goto cleanup;
	}

	wprintf(L"Adding Event Handlers. \nit takes a while till events start comming\n\n");
	hr = pAutomation->AddFocusChangedEventHandler(NULL, (IUIAutomationFocusChangedEventHandler*)pEventHandler);
	if (FAILED(hr))
	{
		ret = 1;
		goto cleanup;
	}

	wprintf(L"Press any key to quit\n");
	getchar();
	wprintf(L"-Removing Event Handlers.\n");
	hr = pAutomation->RemoveFocusChangedEventHandler((IUIAutomationFocusChangedEventHandler*)pEventHandler);
	if (FAILED(hr))
	{
		ret = 1;
		goto cleanup;
	}

	// Release resources and terminate.
cleanup:
	if (pEventHandler != NULL)
		pEventHandler->Release();

	if (pAutomation != NULL)
		pAutomation->Release();

	CoUninitialize();
	return ret;
}
