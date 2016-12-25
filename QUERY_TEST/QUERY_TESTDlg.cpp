
// QUERY_TESTDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "QUERY_TEST.h"
#include "QUERY_TESTDlg.h"
#include "afxdialogex.h"
#include <string>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CQUERY_TESTDlg �Ի���



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


// CQUERY_TESTDlg ��Ϣ�������

BOOL CQUERY_TESTDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CQUERY_TESTDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CQUERY_TESTDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}






void CQUERY_TESTDlg::OnBnClickedConnect()
{
	HRESULT hr;
	try
	{
		hr = m_pConnection.CreateInstance("ADODB.Connection");///����Connection����
		if (SUCCEEDED(hr))
		{
			hr = m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= E:\\test1.mdb", "", "", adModeUnknown);
			MessageBox(_T("���ӳɹ�"));
			///�������ݿ�
		}
	}
	catch (_com_error e)///��׽�쳣
	{
		CString errormessage;
		errormessage.Format(_T("�������ݿ�ʧ��!\r\n������:%s",(e.ErrorMessage())));
		AfxMessageBox(errormessage);///��ʾ������Ϣ
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
			strSQL.Format(_T("SELECT * FROM stuinfo where STUNUMBER =%s"), name); // ����������

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
			message.Format(_T("��ȡ���ݿ�ʧ��!\n ������ϢΪ:%s"), e.ErrorMessage());
			AfxMessageBox(message);///��ʾ������Ϣ
		}// TODO: �ڴ���ӿؼ�֪ͨ����������ѡ���ȡǰ1�� ���������Զ����ٵ�

		

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
				MessageBox(_T("����ɹ�"));

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
		strSQL.Format(_T("Delete  FROM stuinfo where STUNUMBER =%s"), name); // ����������


		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);
		MessageBox(_T("ɾ���ɹ�"));
	}
	catch (_com_error e)
	{
		CString message;
		message.Format(_T("��ȡ���ݿ�ʧ��!\n ������ϢΪ:%s"), e.Description());
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
		
		strSQL.Format(_T("Update stuinfo Set GRADE=%s , GRADE2=%s where STUNUMBER =%s"), GRADE1, GRADE2,name); // ����������


		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdUnknown);
		MessageBox(_T("�޸ĳɹ�"));
	}
	catch (_com_error e)
	{
		CString message;
		message.Format(_T("��ȡ���ݿ�ʧ��!\n ������ϢΪ:%s"), e.Description());
		AfxMessageBox(message);///��ʾ������Ϣ
	}// TODO: 

}


void CQUERY_TESTDlg::OnEnChangeEdit3()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}
