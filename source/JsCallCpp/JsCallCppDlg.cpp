
// JsCallCppDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "JsCallCpp.h"
#include "JsCallCppDlg.h"
#include "afxdialogex.h"
#include <MsHTML.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CJsCallCppDlg 对话框



CJsCallCppDlg::CJsCallCppDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CJsCallCppDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CJsCallCppDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EXPLORER1, m_webbrowser);
}

BEGIN_MESSAGE_MAP(CJsCallCppDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CJsCallCppDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CJsCallCppDlg 消息处理程序

BOOL CJsCallCppDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	//加载资源文件中的HTML,IDR_HTML1就是HTML文件在资源文件中的ID
	wchar_t self_path[MAX_PATH] = { 0 };
	GetModuleFileName(NULL, self_path, MAX_PATH);
	CString res_url;
	res_url.Format(L"res://%s/%d", self_path, IDR_HTML1);
	m_webbrowser.Navigate(res_url, NULL, NULL, NULL, NULL);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CJsCallCppDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CJsCallCppDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//我自己给我的两个函数拟定的数字ID，这个ID可以取0-16384之间的任意数
enum
{
	FUNCTION_ShowMessageBox = 1,
	FUNCTION_GetProcessID = 2,
};

//不用实现，直接返回E_NOTIMPL
HRESULT STDMETHODCALLTYPE CJsCallCppDlg::GetTypeInfoCount(UINT *pctinfo)
{
	return E_NOTIMPL;
}

//不用实现，直接返回E_NOTIMPL
HRESULT STDMETHODCALLTYPE CJsCallCppDlg::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

//JavaScript调用这个对象的方法时，会把方法名，放到rgszNames中，我们需要给这个方法名拟定一个唯一的数字ID，用rgDispId传回给它
//同理JavaScript存取这个对象的属性时，会把属性名放到rgszNames中，我们需要给这个属性名拟定一个唯一的数字ID，用rgDispId传回给它
//紧接着JavaScript会调用Invoke，并把这个ID作为参数传递进来
HRESULT STDMETHODCALLTYPE CJsCallCppDlg::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	//rgszNames是个字符串数组，cNames指明这个数组中有几个字符串，如果不是1个字符串，忽略它
	if (cNames != 1)
		return E_NOTIMPL;
	//如果字符串是ShowMessageBox，说明JavaScript在调用我这个对象的ShowMessageBox方法，我就把我拟定的ID通过rgDispId告诉它
	if (wcscmp(rgszNames[0], L"ShowMessageBox") == 0)
	{
		*rgDispId = FUNCTION_ShowMessageBox;
		return S_OK;
	}
	//同理，如果字符串是GetProcessID，说明JavaScript在调用我这个对象的GetProcessID方法
	else if (wcscmp(rgszNames[0], L"GetProcessID") == 0)
	{
		*rgDispId = FUNCTION_GetProcessID;
		return S_OK;
	}
	else
		return E_NOTIMPL;
}

//JavaScript通过GetIDsOfNames拿到我的对象的方法的ID后，会调用Invoke，dispIdMember就是刚才我告诉它的我自己拟定的ID
//wFlags指明JavaScript对我的对象干了什么事情！
//如果是DISPATCH_METHOD，说明JavaScript在调用这个对象的方法，比如cpp_object.ShowMessageBox();
//如果是DISPATCH_PROPERTYGET，说明JavaScript在获取这个对象的属性，比如var n = cpp_object.num;
//如果是DISPATCH_PROPERTYPUT，说明JavaScript在修改这个对象的属性，比如cpp_object.num = 10;
//如果是DISPATCH_PROPERTYPUTREF，说明JavaScript在通过引用修改这个对象，具体我也不懂
//示例代码并没有涉及到wFlags和对象属性的使用，需要的请自行研究，用法是一样的
//pDispParams就是JavaScript调用我的对象的方法是传递进来的参数，里面有一个数组保存着所有参数
//pDispParams->cArgs就是数组中有多少个参数
//pDispParams->rgvarg就是保存着参数的数组，请使用[]下标来访问，每个参数都是VARIANT类型，可以保存各种类型的值
//具体是什么类型用VARIANT::vt来判断，不多解释了，VARIANT这东西大家都懂
//pVarResult就是我们给JavaScript的返回值
//其它不用管
HRESULT STDMETHODCALLTYPE CJsCallCppDlg::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	//通过ID我就知道JavaScript想调用哪个方法
	if (dispIdMember == FUNCTION_ShowMessageBox)
	{
		//检查是否只有一个参数
		if (pDispParams->cArgs != 1)
			return E_NOTIMPL;
		//检查这个参数是否是字符串类型
		if (pDispParams->rgvarg[0].vt != VT_BSTR)
			return E_NOTIMPL;
		//放心调用
		ShowMessageBox(pDispParams->rgvarg[0].bstrVal);
		return S_OK;
	}
	else if (dispIdMember == FUNCTION_GetProcessID)
	{
		DWORD id = GetProcessID();
		*pVarResult = CComVariant(id);
		return S_OK;
	}
	else
		return E_NOTIMPL;

}


//JavaScript拿到我们传递给它的指针后，由于它不清楚我们的对象是什么东西，会调用QueryInterface来询问我们“你是什么鬼东西？”
//它会通过riid来问我们是什么东西，只有它问到我们是不是IID_IDispatch或我们是不是IID_IUnknown时，我们才能肯定的回答它S_OK
//因为我们的对象继承于IDispatch，而IDispatch又继承于IUnknown，我们只实现了这两个接口，所以只能这样来回答它的询问
HRESULT STDMETHODCALLTYPE CJsCallCppDlg::QueryInterface(REFIID riid, void **ppvObject)
{
	if (riid == IID_IDispatch || riid == IID_IUnknown)
	{
		//对的，我是一个IDispatch，把我自己(this)交给你
		*ppvObject = static_cast<IDispatch*>(this);
		return S_OK;
	}
	else
		return E_NOINTERFACE;
}

//我们知道COM对象使用引用计数来管理对象生命周期，我们的CJsCallCppDlg对象的生命周期就是整个程序的声明周期
//我的这个对象不需要你JavaScript来管，我自己会管，所以我不用实现AddRef()和Release()，这里乱写一些。
//你要return 1;return 2;return 3;return 4;return 5;都可以
ULONG STDMETHODCALLTYPE CJsCallCppDlg::AddRef()
{
	return 1;
}

//同上，不多说了
//题外话：当然如果你要new出一个c++对象来并扔给JavaScript来管，你就需要实现AddRef()和Release()，在引用计数归零时delete this;
ULONG STDMETHODCALLTYPE CJsCallCppDlg::Release()
{
	return 1;
}

//调用JavaScript的SaveCppObject函数，把我自己(this)交给它，SaveCppObject会把我这个对象保存到全局变量var cpp_object;中
//以后JavaScript就可以通过cpp_object来调用我这个C++对象的方法了
void CJsCallCppDlg::OnBnClickedOk()
{
	CComQIPtr<IHTMLDocument2> document = m_webbrowser.get_Document();
	CComDispatchDriver script;
	document->get_Script(&script);
	CComVariant var(static_cast<IDispatch*>(this));
	script.Invoke1(L"SaveCppObject", &var);
}

DWORD CJsCallCppDlg::GetProcessID()
{
	return GetCurrentProcessId();
}

void CJsCallCppDlg::ShowMessageBox(const wchar_t *msg)
{
	MessageBox(msg, L"这是来自javascript的消息");
}

