<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'Option Explicit		'ǿ������
On Error Resume Next		'�ݴ�����
Dim CODEPAGE: CODEPAGE = "936"		'ҳ�����65001|936
Dim CHARSET: CHARSET = "GB2312"		'��������utf-8|gb2312
'=========================================================
' File Name��	config.asp
' Purpose��		ϵͳ�����ļ�
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Created on: 	2009-9-9 10:27:17
' Update on: 	2009-10-21 23:24:44
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================

Dim DBPATH		'Access���ݿ�·��
	DBPATH = "database/Fl28#Ek_7348D432AF.mdb"

Dim SITENAME		'��վ����
	SITENAME = "E��Ƽ���"

Dim HTTPURL		'��վ��ַǰ׺
	HTTPURL = "http://localhost"

Dim INSTALLDIR		'��վ��װĿ¼����Ŀ¼��Ϊ��/
	INSTALLDIR = "/eekku"

Dim KEYWORDS		'��վ�ؼ���
	KEYWORDS = "E������E��Cms��E�Ṥ����,www.eekku.com���������£�ling.liufu.org___"

Dim DESCRIPTION		'��վ����
	DESCRIPTION = "E��Cms(EekkuCMS)��һ��С��վ�����ݹ���ϵͳ���ں����¡�ͼƬ�����ԵȻ������ܣ����ҷ���ʹ��!__"

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
	ISCACHE = 0

Dim CACHEFLAG		'�����־����������Ӣ����ĸ
	CACHEFLAG = "Eekku_"

Dim CACHETIME		'����ʱ�䣬Ĭ����60��
	CACHETIME = 60

Dim ISWEBLOG		'�Ƿ��¼��̨����������¼
	ISWEBLOG = 1

Dim LIMITIP		'����IP������|���зָ�
	LIMITIP = "127.0.0.2|127.0.0.3"

Dim DIRTYWORDS		'�໰����,����|���зָ�
	DIRTYWORDS = "fuck|sex"

%>
