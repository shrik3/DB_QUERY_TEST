
// QUERY_TESTDlg.h : 头文件
//

#pragma once


// CQUERY_TESTDlg 对话框
class CQUERY_TESTDlg : public CDialogEx
{
// 构造
public:
	CQUERY_TESTDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_QUERY_TEST_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedAbort2();
	_ConnectionPtr m_pConnection;
	_RecordsetPtr  m_pRecordset;//数据集连接
	_variant_t vUsername, vFeature;	//存放取出来的数据的变量


	afx_msg void OnBnClickedConnect();
	afx_msg void OnBnClickedSearch();
	afx_msg void OnBnClickedInsert();
	afx_msg void OnBnClickedDelete();
	afx_msg void OnBnClickedUpdate();
	afx_msg void OnEnChangeEdit3();
};
