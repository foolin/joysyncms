<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'Option Explicit		'ǿ������
On Error Resume Next		'�ݴ���
Dim CODEPAGE: CODEPAGE = "936"		'ҳ�����65001|936
Dim CHARSET: CHARSET = "GB2312"		'��������utf-8|gb2312
'=========================================================
' File Name��	config.asp
' Purpose��		ϵͳ�����ļ�
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Created on: 	2010-7-17 16:04:23
' Copyright (c) 2010 ׿�����磨Foolin��All Rights Reserved
'=========================================================

Dim DBPATH		'Access���ݿ�·��
	DBPATH = "database/Fl17#Ek_6C7A80B0FF.mdb"

Dim SITENAME		'��վ����
	SITENAME = "׿������"

Dim HTTPURL		'��վ��ַǰ׺
	HTTPURL = "http://127.0.0.1"

Dim INSTALLDIR		'��վ��װĿ¼����Ŀ¼��Ϊ��/
	INSTALLDIR = "/"

Dim KEYWORDS		'��վ�ؼ���
	KEYWORDS = "׿�����磬׿�����ݹ���ϵͳ,JSCMS��׿������,www.joysyn.com���������£�ling.liufu.org"

Dim DESCRIPTION		'��վ����
	DESCRIPTION = "׿�����ݹ���ϵͳ��JSCMS����һ��С��վ�����ݹ���ϵͳ���ں����¡�ͼƬ�����ԵȻ������ܣ����ҷ���ʹ�ã�"

Dim TEMPLATEDIR		'��վģ��·�������磺default��ʾtemplate/default/
	TEMPLATEDIR = "default"

Dim ISHIDETEMPPATH		'�Ƿ�����ģ��·�����������Ӱ�������ٶ�
	ISHIDETEMPPATH = 0

Dim ISOPENGBOOK		'�Ƿ񿪷����ԣ�Ĭ�Ͽ���
	ISOPENGBOOK = 1

Dim ISAUDITGBOOK		'�Ƿ���Ҫ������ԣ���-1����-0
	ISAUDITGBOOK = 0

Dim GBOOKTIME		'�����������ʱ��������λ�룬Ĭ��60��
	GBOOKTIME = 60

Dim ISCACHE		'�Ƿ񻺴棬�����ǣ����������������
	ISCACHE = 1

Dim CACHEFLAG		'�����־����������Ӣ����ĸ
	CACHEFLAG = "EekkuCms_"

Dim CACHETIME		'����ʱ�䣬Ĭ����60��
	CACHETIME = 60

Dim ISWEBLOG		'�Ƿ��¼��̨���������¼
	ISWEBLOG = 1

Dim LIMITIP		'����IP������|���зָ�
	LIMITIP = ""

Dim DIRTYWORDS		'�໰���ˣ�����|���зָ�
	DIRTYWORDS = "fuck|sex"

%>

