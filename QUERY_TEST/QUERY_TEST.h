
// QUERY_TEST.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CQUERY_TESTApp: 
// �йش����ʵ�֣������ QUERY_TEST.cpp
//

class CQUERY_TESTApp : public CWinApp
{
public:
	CQUERY_TESTApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CQUERY_TESTApp theApp;