
// QUERY_TESTDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "QUERY_TEST.h"
#include "QUERY_TESTDlg.h"
#include "afxdialogex.h"
#include <string>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CQUERY_TESTDlg 对话框



CQUERY_TESTDlg::CQUERY_TESTDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_QUERY_TEST_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CQUERY_TESTDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CQUERY_TESTDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_Connect, &CQUERY_TESTDlg::OnBnClickedConnect)
	ON_BN_CLICKED(IDC_Search, &CQUERY_TESTDlg::OnBnClickedSearch)
	ON_BN_CLICKED(IDC_Insert, &CQUERY_TESTDlg::OnBnClickedInsert)
	ON_BN_CLICKED(IDC_Delete, &CQUERY_TESTDlg::OnBnClickedDelete)
	ON_BN_CLICKED(IDC_Update, &CQUERY_TESTDlg::OnBnClickedUpdate)
	ON_EN_CHANGE(IDC_EDIT3, &CQUERY_TESTDlg::OnEnChangeEdit3)
END_MESSAGE_MAP()


// CQUERY_TESTDlg 消息处理程序

BOOL CQUERY_TESTDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CQUERY_TESTDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CQUERY_TESTDlg::OnPaint()
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
HCURSOR CQUERY_TESTDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}






void CQUERY_TESTDlg::OnBnClickedConnect()
{
	HRESULT hr;
	try
	{
		hr = m_pConnection.CreateInstance("ADODB.Connection");///创建Connection对象
		if (SUCCEEDED(hr))
		{
			hr = m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= E:\\test1.mdb", "", "", adModeUnknown);
			MessageBox(_T("连接成功"));
			///连接数据库
		}
	}
	catch (_com_error e)///捕捉异常
	{
		CString errormessage;
		errormessage.Format(_T("连接数据库失败!\r\n错误信:%s",(e.ErrorMessage())));
		AfxMessageBox(errormessage);///显示错误信息
		return;
	}

}


void CQUERY_TESTDlg::OnBnClickedSearch()
{
		CString CFeature;
		try
		{
			CString strSQL;
			CString name;
			GetDlgItem(IDC_EDIT1)->GetWindowText(name);
			strSQL.Format(_T("SELECT * FROM stuinfo where STUNUMBER =%s"), name); // 有条件查找

			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);


			vFeature = m_pRecordset->GetCollect("GRADE");
			CFeature = vFeature;
			GetDlgItem(IDC_EDIT2)->SetWindowText(CFeature);;

			vFeature = m_pRecordset->GetCollect("GRADE2");
			CFeature = vFeature;
			GetDlgItem(IDC_EDIT3)->SetWindowText(CFeature);;
			

		}
		catch (_com_error e)
		{
			CString message;
			message.Format(_T("读取数据库失败!\n 错误信息为:%s"), e.ErrorMessage());
			AfxMessageBox(message);///显示错误信息
		}// TODO: 在此添加控件通知处理程序代码选择读取前1个 测试用所以读的少点

		

}
	


void CQUERY_TESTDlg::OnBnClickedInsert()
{
			try
			{
				CString strSQL;
				CString name;
				CString GRADE1;
				CString GRADE2;
				GetDlgItem(IDC_EDIT1)->GetWindowText(name);
				GetDlgItem(IDC_EDIT2)->GetWindowText(GRADE1);
				GetDlgItem(IDC_EDIT3)->GetWindowText(GRADE2);

				strSQL.Format(_T("insert into stuinfo(STUNUMBER,GRADE,GRADE2) values(%s,%s,%s)"), name, GRADE1,GRADE2);
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);
				MessageBox(_T("插入成功"));

			}
			catch (_com_error e)
			{
				AfxMessageBox(e.Description());
				AfxMessageBox(e.ErrorMessage());
				return;
			}

}


void CQUERY_TESTDlg::OnBnClickedDelete()
{
	try
	{
		CString strSQL;
		CString name;
		GetDlgItem(IDC_EDIT1)->GetWindowText(name);
		strSQL.Format(_T("Delete  FROM stuinfo where STUNUMBER =%s"), name); // 有条件查找


		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);
		MessageBox(_T("删除成功"));
	}
	catch (_com_error e)
	{
		CString message;
		message.Format(_T("读取数据库失败!\n 错误信息为:%s"), e.Description());
		AfxMessageBox(message);
	}
}


void CQUERY_TESTDlg::OnBnClickedUpdate()
{
	try
	{
		CString strSQL;
		CString name, GRADE1,GRADE2;
		GetDlgItem(IDC_EDIT1)->GetWindowText(name);
		GetDlgItem(IDC_EDIT2)->GetWindowText(GRADE1);
		GetDlgItem(IDC_EDIT3)->GetWindowText(GRADE2);
		
		strSQL.Format(_T("Update stuinfo Set GRADE=%s , GRADE2=%s where STUNUMBER =%s"), GRADE1, GRADE2,name); // 有条件查找


		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);
		MessageBox(_T("修改成功"));
	}
	catch (_com_error e)
	{
		CString message;
		message.Format(_T("读取数据库失败!\n 错误信息为:%s"), e.Description());
		AfxMessageBox(message);///显示错误信息
	}// TODO: 

}


void CQUERY_TESTDlg::OnEnChangeEdit3()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}
