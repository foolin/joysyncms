<%
'=========================================================
' Class Name��	ClassCreateHtml
' Purpose��		�����ļ���
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2010-7-17 22:15:28
' Modify log:	
' Updated on: 	
' Copyright (c) 2010 ׿�����磨JoySyn.com��All Rights Reserved
'=========================================================
Class ClassCreateHtml

	Dim mTpl	'ģ����ʵ��
	Dim mHtmlDir
	Dim mHtmlArtDir
	Dim mHtmlPicDir
	Dim mHtmlDiyDir
	
	Private mMessage	'�����Ϣ
	Public Property Get Message
		Message = mMessage
	End Property
	
	Private mErrMessage	'������Ϣ
	Public Property Get ErrMessage
		ErrMessage = mErrMessage
	End Property

	'��ʼ������
	Private Sub Class_Initialize()
		Call ChkLogin()		'����¼
		Set mTpl = New ClassTemplate	'����ģ��
		mHtmlDir = "/Html/"	'HtmlĿ¼
		mHtmlArtDir = mHtmlDir & "Article/"
		mHtmlPicDir = mHtmlDir & "Picture/"
		mHtmlDiyDir = mHtmlDir & "DiyPage/"
		If ExistFolder(mHtmlDir) = False Then
			Call CreateFolder(mHtmlDir)	'����Html�ļ���
		End If
		mMessage = ""
	End Sub
	
	'��������
	Private Sub Class_Terminate()
		Set mTpl = Nothing	'�ͷ�ģ��
	End Sub


	'������ҳ
	Public Function CreateIndex()
		Dim blnFlag: blnFlag = False

		'����ȫ�� ���� {sys: var}
		page = 1
		id = colId
		Title = "�����б�": If id > 0 Then Title = GetNameOfColumn(id, "ARTICLE") & " - �����б�"
		SitePath = ColPath(id, 0)
		
		'��ʼ����
		Dim tplIndexPath: tplIndexPath = mHtmlDir & "index.html"
		Call mTpl.LoadTpl("index.html")		'����ģ��
		Call mTpl.Compile_Index()		'���б�ǩ����
		Call CreateFile(mTpl.Content, tplIndexPath)	'������ҳ�ļ�
		mMessage = mMessage & "������ҳindex.html�ɹ���<br />"
		blnFlag = True
		CreateIndex = blnFlag
	End Function
	
	'===================���������б� ================
	
	'���������б�&����
	Public Function CreateAllArticles()
		Dim blnFlag: blnFlag = False
		
		Dim strMessage
		mMessage = mMessage & "��ʼ����ȫ�����¼��б�ҳ...<br />"
		'����Html/Article�ļ���
		Call CreateFolder(mHtmlArtDir)
		'����������ĿĿ¼
		Call CreateAllArtFolder()
		'�������������б�
		CreateAllArtListHtml()
		'������������
		CreateAllArtHtml()
		strMessage = mMessage & "����ȫ�����¼��б�ҳ�ɹ���<br />"
		mMessage = strMessage
		
		blnFlag = True
		CreateAllArticles = blnFlag
	End Function
	
	'���´��������б�&����
	Public Function ReCreateAllArticles()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "���´���ȫ������ҳ...<br />"
		'���ArticleĿ¼�������ļ�
		Call ClearAllArticles()
		'��������ȫ��
		CreateAllArticles()
		mMessage = mMessage & "���´���ȫ������ҳ��ϣ�<br />"
		
		blnFlag = True
		ReCreateAllArticles = blnFlag
	End Function
	
	
	'���������б�
	Public Function CreateAllArtListHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "��ʼ����ȫ�������б�ҳ...<br />"
		'����ȫ�������б�
		Call CreateArtListHtmlByID(0)
		'ѭ��������Ŀ����
		Dim Rs
		Set Rs = DB( "SELECT ID FROM ArtColumn", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Call CreateArtListHtmlByID(Rs("ID"))
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ȫ�������б�ҳ�ɹ���<br />"
		
		blnFlag = True
		CreateAllArtListHtml = blnFlag
	End Function
	
	'�������������б�ҳ
	Public Function CreateArtListHtmlByID(colId)
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "���������б�ҳ...<br />"
		
		'�����б���ر���
		Dim pageCount: pageCount = 1
		Dim listDir: listDir = mHtmlArtDir
		
		'����ȫ�� ���� {sys: var}
		page = 1
		id = colId
		Title = "�����б�": If id > 0 Then Title = GetNameOfColumn(id, "ARTICLE") & " - �����б�"
		SitePath = ColPath(id, 0)	
		
		
		'ѭ����������
		If CInt(id) > 0 Then
			Dim Rs
			Set Rs = DB( "SELECT CreateHtmlPath FROM ArtColumn WHERE ID=" & id, 1)
			If Not Rs.Eof Then
				listDir = Rs("CreateHtmlPath")
			Else
				mErrMessage = mErrMessage & "��������ĿIDΪ[" & id & "]�ļ�¼��"
				CreateArtListHtmlByID = False
			End If
		End If
		Do While True
			mTpl.Page = page						'���õ�ǰҳ
			Call mTpl.LoadTpl("artlist.html")		'����ģ��
			Call mTpl.Compile_List(id)				'���б�ǩ����
			pageCount = mTpl.PageCount
			Dim listPath: listPath = ""
			'����List
			If id > 0 Then
				If Page>1 Then
					listPath = listDir & "list_" & id & "_" & page & ".html"	'ĳ��Ŀ�����б�list_��ĿID_ҳ��.html
				Else
					listPath = listDir & "list_" & id & ".html"	'ĳ��Ŀ�����б�list_��ĿID.html ����ǵ�һҳ
				End If
			Else
				If Page>1 Then
					listPath = listDir & "list_" & page & ".html"	'ȫ�������б�list_ҳ��.html
				Else
					listPath = listDir & "list" & ".html"	'ȫ�������б�list.html ����ǵ�һҳ
				End If
			End If
			Call CreateFile(mTpl.Content, listPath)
			mMessage = mMessage & "�ɹ�����" & listPath & "!<br />"
			page = page + 1	'ҳ��+1
			If Page>pageCount Then Exit Do '���������ѭ��
		Loop
		mMessage = mMessage & "���������б���ϣ�<br />"
		
		blnFlag = True
		CreateArtListHtmlByID = blnFlag
	End Function
	
	
	'������������ҳ
	Public Function CreateAllArtHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "����ȫ������...<br />"
			
		'ѭ����������
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "����": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)
			
				'����Ŀ¼
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Call mTpl.LoadTpl("article.html")		'����ģ��
				Call mTpl.Compile_Field(id, False)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strArtPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����������ϣ�<br />"
		
		blnFlag = True
		CreateAllArtHtml = blnFlag
	End Function
	
	'����ĳ��Ŀ�µ���������
	Public Function CreateArtHtmlByColID(colId)
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "��������...<br />"
		Dim colIds: colIds = GetColIds(colId, "article")	'��ȡ��������ĿID������ID
		'ѭ����������
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.ColID IN (" & colIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "����": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)	
			
				'����Ŀ¼
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Call mTpl.LoadTpl("article.html")		'����ģ��
				Call mTpl.Compile_Field(id, False)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strArtPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����������ϣ�<br />"

		blnFlag = True
		CreateArtHtmlByColID = blnFlag
	End Function
	
	'����ָ��ID������
	Public Function CreateArtHtmlByID(artIds)
		Dim blnFlag: blnFlag = False

		'��������ҳ
		mMessage = mMessage & "��������...<br />"
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.ID IN(" & artIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "����": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)	
			
				'����Ŀ¼
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Call mTpl.LoadTpl("article.html")		'����ģ��
				Call mTpl.Compile_Field(id, False)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strArtPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����������ϣ�<br />"

		blnFlag = True
		CreateArtHtmlByID = blnFlag
	End Function
	
	
	'����������¼��ļ���
	Public Function ClearAllArticles()
		Dim blnFlag: blnFlag = False
		
		'������ڸ�Ŀ¼�������
		If ExistFolder(mHtmlArtDir) = True Then
			DeleteFolder(mHtmlArtDir)
		End If
		mMessage = mMessage & "�Ѿ�ɾ��" & mHtmlArtDir & "Ŀ¼����Ŀ¼�µ������ļ�...<br />"

		blnFlag = True
		ClearAllArticles = blnFlag
	End Function
	
	
	
	'��������Ŀ¼����һ����Ŀ����
	Public Function CreateAllArtFolder()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "��ʼ����ȫ������Ŀ¼...<br />"
		Dim Rs
		Set Rs = DB( "SELECT * FROM ArtColumn WHERE ParentID = 0 ORDER BY Sort DESC,ID", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Dim colDirPath
				'Echo("<option value=""" & Rs("ID") & """>" & Rs("Name") & "</option>" & Chr(10) & Chr(9) & Chr(9))
				If Rs("DirName")<>"" Then
					colDirPath = mHtmlArtDir & Rs("DirName") & "/"
				Else
					colDirPath = mHtmlArtDir & "Col_" & Rs("ID") & "/"
				End If
				'����Ŀ¼·�������ݿ�
				Call DB("Update ArtColumn set CreateHtmlPath='" & colDirPath & "' where ID=" & Rs("ID"), 0)
				'����Ŀ¼
				Call CreateFolder(colDirPath)
				mMessage = mMessage & "�ɹ�����Ŀ¼" & colDirPath & "<br />"
				Call P_CreateSubArtFolder(Rs("ID"), colDirPath) 'ѭ���Ӽ�����
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ȫ������Ŀ¼��ϣ�<br />"
		
		blnFlag = True
		CreateAllArtFolder = blnFlag
	End Function
	
	'����Ŀ����
	Private Function P_CreateSubArtFolder(FID,StrDis)
		Dim Rs1
		Set Rs1 = DB("SELECT * FROM ArtColumn WHERE ParentID = " & FID & " ORDER BY Sort DESC,ID", 1)
		If Not Rs1.Eof Then
			Do While Not Rs1.Eof
				'Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
				If Rs1("DirName")<>"" Then
					StrDis = StrDis & Rs1("DirName") & "/"
				Else
					StrDis = StrDis & "Col_" & Rs1("ID") & "/"
				End If
				
				Call DB("Update ArtColumn set CreateHtmlPath='" & StrDis & "' where ID=" & Rs1("ID"), 0)
				Call CreateFolder(StrDis)
				mMessage = mMessage & "�ɹ�����Ŀ¼" & StrDis & "<br />"
				Call P_CreateSubArtFolder(Trim(Rs1("ID")), Strdis) '�ݹ��Ӽ�����
			Rs1.Movenext:Loop
			If Rs1.Eof Then
				Rs1.Close: Set Rs1 = Nothing
				Exit Function
			End If
		End If
		Rs1.Close: Set Rs1 = Nothing
	End Function
	
	
	'===================����ͼƬ�б� ================
	
	'����ͼƬ�б�&ͼƬ
	Public Function CreateAllPictures()
		Dim blnFlag: blnFlag = False
		
		Dim strMessage
		strMessage = strMessage & "��ʼ����ȫ��ͼƬҳ..."
		'����Html�ļ���
		Call CreateFolder(mHtmlPicDir)	
		'ѭ������Ŀ¼
		Call CreateAllPicFolder()	
		'����ͼƬ�б�
		CreateAllPicListHtml()	
		'����ͼƬ
		CreateAllPicHtml()
		strMessage = mMessage & "����ȫ��ͼƬҳ�ɹ���<br />"

		mMessage = strMessage

		blnFlag = True
		CreateAllPictures = blnFlag
	End Function
	
	'���´��������б�&����
	Public Function ReCreateAllPictures()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "���´���ȫ��ͼƬҳ...<br />"
		'���ArticleĿ¼�������ļ�
		Call ClearAllPictures()
		'��������ȫ��
		CreateAllPictures()
		mMessage = mMessage & "���´���ȫ������ҳ��ϣ�<br />"
		
		blnFlag = True
		ReCreateAllPictures = blnFlag
	End Function
	
	'����ͼƬ�б�
	Public Function CreateAllPicListHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "��ʼ����ȫ��ͼƬ�б�ҳ...<br />"
		'����Ĭ������ͼƬ�б�
		Call CreatePicListHtmlByID(0)
		'ѭ��������ĿͼƬ
		Dim Rs
		Set Rs = DB( "SELECT ID FROM PicColumn", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Call CreatePicListHtmlByID(Rs("ID"))
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ȫ��ͼƬ�б�ҳ��ϣ�<br />"
		
		blnFlag = True
		CreateAllPicListHtml = blnFlag
	End Function
	
	'��������ͼƬ�б�ҳ
	Public Function CreatePicListHtmlByID(colId)
		Dim blnFlag: blnFlag = False
		'�����б���ر���
		Dim pageCount: pageCount = 1
		Dim listDir: listDir = mHtmlPicDir
		
		'����ȫ�� ���� {sys: var}
		page = 1
		id = colId
		Title = "ͼƬ�б�": If id > 0 Then Title = GetNameOfColumn(id, "PICTURE") & " - ͼƬ�б�"
		SitePath = ColPath(id, 0)	
		
		
		'ѭ������ͼƬ
		If CInt(id) > 0 Then
			Dim Rs
			Set Rs = DB( "SELECT CreateHtmlPath FROM PicColumn WHERE ID=" & id, 1)
			If Not Rs.Eof Then
				listDir = Rs("CreateHtmlPath")
			Else
				mErrMessage = mErrMessage & "��������ĿIDΪ[" & id & "]�ļ�¼��<br />"
				CreatePicListHtmlByID = False
			End If
		End If
		Do While True
			mTpl.Page = page						'���õ�ǰҳ
			Call mTpl.LoadTpl("piclist.html")		'����ģ��
			Call mTpl.Compile_List(id)				'���б�ǩ����
			pageCount = mTpl.PageCount
			Dim listPath: listPath = ""
			'����List
			If id > 0 Then
				If Page>1 Then
					listPath = listDir & "list_" & id & "_" & page & ".html"	'ĳ��ĿͼƬ�б�list_��ĿID_ҳ��.html
				Else
					listPath = listDir & "list_" & id & ".html"	'ĳ��ĿͼƬ�б�list_��ĿID.html ����ǵ�һҳ
				End If
			Else
				If Page>1 Then
					listPath = listDir & "list_" & page & ".html"	'ȫ��ͼƬ�б�list_ҳ��.html
				Else
					listPath = listDir & "list" & ".html"	'ȫ��ͼƬ�б�list.html ����ǵ�һҳ
				End If
			End If
			Call CreateFile(mTpl.Content, listPath)
			mMessage = mMessage & "�ɹ�����" & listPath & "!<br />"
			page = page + 1	'ҳ��+1
			If Page>pageCount Then Exit Do '���������ѭ��
		Loop
		
		blnFlag = True
		CreatePicListHtmlByID = blnFlag
	End Function
	
	
	'��������ͼƬҳ
	Public Function CreateAllPicHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "����ͼƬ...<br />"
			
		'ѭ������ͼƬ
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "ͼƬ": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)
				

			
				'ͼƬĿ¼
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'����ģ��
				Call mTpl.Compile_Field(id, True)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strPicPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strPicPath & "<br />"

			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ͼƬ��ϣ�<br />"

		blnFlag = True
		CreateAllPicHtml = blnFlag
	End Function
	
	'����ĳ��Ŀ�µ�����ͼƬ
	Public Function CreatePicHtmlByColID(colId)
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "����ͼƬ...<br />"
		Dim colIds: colIds = GetColIds(colId, "picture")	'��ȡ��������ĿID������ID
		'ѭ������ͼƬ
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.ColID IN (" & colIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "ͼƬ": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)	
			
				'ͼƬĿ¼
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'����ģ��
				Call mTpl.Compile_Field(id, True)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strPicPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strPicPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ͼƬ��ϣ�<br />"

		blnFlag = True
		CreatePicHtmlByColID = blnFlag
	End Function
	
	'����ָ��ID��ͼƬ
	Public Function CreatePicHtmlByID(picIds)
		Dim blnFlag: blnFlag = False

		'����ͼƬ
		mMessage = mMessage & "����ͼƬ...<br />"
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.ID IN(" & picIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				id = Rs("ID")
				Title = "ͼƬ": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)	
			
				'ͼƬĿ¼
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'��������
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'����ģ��
				Call mTpl.Compile_Field(id, True)	'���б�ǩ����
				'�����ļ�
				Call CreateFile(mTpl.Content, strPicPath)			'�������
				mMessage = mMessage & "�ɹ�����" & strPicPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ͼƬ��ϣ�<br />"

		blnFlag = True
		CreatePicHtmlByID = blnFlag
	End Function


	
	'����������¼��ļ���
	Public Function ClearAllPictures()
		Dim blnFlag: blnFlag = False
		
		'������ڸ�Ŀ¼�������
		If ExistFolder(mHtmlPicDir) = True Then
			DeleteFolder(mHtmlPicDir)
		End If
		mMessage = mMessage & "ɾ��" & mHtmlPicDir & "Ŀ¼�������ļ��ɹ���<br />"

		blnFlag = True
		ClearAllPictures = blnFlag
	End Function
	
	'����ͼƬĿ¼����һ����Ŀ
	Public Function CreateAllPicFolder()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "��ʼ����ȫ��ͼƬĿ¼...<br />"
		Dim Rs
		Set Rs = DB( "SELECT * FROM PicColumn WHERE ParentID = 0 ORDER BY Sort DESC,ID", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Dim colDirPath
				If Rs("DirName")<>"" Then
					colDirPath = mHtmlPicDir & Rs("DirName") & "/"
				Else
					colDirPath = mHtmlPicDir & "Col_" & Rs("ID") & "/"
				End If
				'����Ŀ¼·�������ݿ�
				Call DB("Update PicColumn set CreateHtmlPath='" & colDirPath & "' where ID=" & Rs("ID"), 0)
				'����Ŀ¼
				Call CreateFolder(colDirPath)
				mMessage = mMessage & "����Ŀ¼" & colDirPath & "<br />"
				'ѭ���Ӽ�����
				Call P_CreateSubPicFolder(Rs("ID"), colDirPath) 
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "��ʼ����ȫ��ͼƬĿ¼�ɹ�!<br />"
		
		blnFlag = True
		CreateAllPicFolder = blnFlag
	End Function
	
	'����Ŀ
	Private Function P_CreateSubPicFolder(FID,StrDis)
		Dim Rs1
		Set Rs1 = DB("SELECT * FROM PicColumn WHERE ParentID = " & FID & " ORDER BY Sort DESC,ID", 1)
		If Not Rs1.Eof Then
			Do While Not Rs1.Eof
				'Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
				If Rs1("DirName")<>"" Then
					StrDis = StrDis & Rs1("DirName") & "/"
				Else
					StrDis = StrDis & "Col_" & Rs1("ID") & "/"
				End If
				'����Ŀ¼�����ݿ�
				Call DB("Update PicColumn set CreateHtmlPath='" & StrDis & "' where ID=" & Rs1("ID"), 0)
				'����Ŀ¼
				Call CreateFolder(StrDis)
				mMessage = mMessage & "����Ŀ¼" & StrDis & "<br />"
				Call P_CreateSubPicFolder(Trim(Rs1("ID")), Strdis) '�ݹ��Ӽ�����
			Rs1.Movenext:Loop
			If Rs1.Eof Then
				Rs1.Close: Set Rs1 = Nothing
				Exit Function
			End If
		End If
		Rs1.Close: Set Rs1 = Nothing
	End Function
	
	
	'===================�����Զ���ҳ�� ================
	
	'�����Զ���ҳ��
	Public Function CreateAllDiyHtml()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "��ʼ����ȫ��Diyҳ��...<br />"
		'����DIYҳ��
		Dim Rs,strSql
		strSql = "SELECT * FROM DiyPage WHERE State = 1"
		Set Rs = DB(strSql, 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'����ȫ�� ���� {sys: var}
				page = 1
				'id = Rs("ID")
				Title = "DIYҳ��": Title = GetTitleOfDiypage(Rs("ID"))
				SitePath = DiyPagePath(Rs("ID"))
				
				Dim strDiyPath
				If Rs("PageName")<>"" Then
					strDiyPath = mHtmlDiyDir & Rs("PageName")
				Else
					strDiyPath = mHtmlDiyDir & "html_" & Rs("ID") & ".html"
				End If
				mTpl.Page = page						'���õ�ǰҳ
				Call mTpl.Compile_DiyPage(Rs("ID"))		'���б�ǩ����
				Call CreateFile(mTpl.Content, strDiyPath)
				mMessage = mMessage & "�ɹ�����DIYҳ��" & strDiyPath & "��<br />"
				
			Rs.MoveNext
			If Rs.Eof Then Exit Do '���������ѭ��
			Loop
		Else
			mErrMessage = mErrMessage & "����ʧ�ܣ�������DIYҳ��[" & param & "]�ļ�¼��<br />"
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "����ȫ��Diyҳ��!<br />"
		
		blnFlag = True
		CreateAllDiyHtml = blnFlag
	End Function
	
	Public Function CreateDiyHtmlByParam(param)
		Dim blnFlag: blnFlag = False
		If Len(param)=0 Then mErrMassage = "����CreateDiyHtmlByID(param)������param����Ϊ�գ�": Exit Function
		
		'����ȫ�� ���� {sys: var}
		page = 1
		'id = Rs("ID")
		Title = "DIYҳ��": Title = GetTitleOfDiypage(param)
		SitePath = DiyPagePath(param)
		
		
		'����DIYҳ��
		Dim Rs,strSql
		If IsNumeric(param) Then
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND ID =  " & param
		Else
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND PageName = '" & param &"'"
		End If
		Set Rs = DB(strSql, 1)
		If Not Rs.Eof Then
			Dim strDiyPath
			If Rs("PageName")<>"" Then
				strDiyPath = mHtmlDiyDir & Rs("PageName")
			Else
				strDiyPath = mHtmlDiyDir & "html_" & Rs("ID") & ".html"
			End If
			mTpl.Page = page						'���õ�ǰҳ
			Call mTpl.Compile_DiyPage(param)			'���б�ǩ����
			Call CreateFile(mTpl.Content, strDiyPath)
			mMessage = mMessage & "�ɹ�����DIYҳ��" & strDiyPath & "��<br />"
		Else
			mErrMessage = mErrMessage & "����ʧ�ܣ�������DIYҳ��[" & param & "]�ļ�¼��<br />"
		End If
		Rs.Close: Set Rs = Nothing
		
		blnFlag = True
		CreateDiyHtmlByParam = blnFlag
	End Function
	
	'����������¼��ļ���
	Public Function ClearAllDiyPages()
		Dim blnFlag: blnFlag = False
		
		'������ڸ�Ŀ¼�������
		If ExistFolder(mHtmlDiyDir) = True Then
			Call DeleteFolder(mHtmlDiyDir)
		End If
		mMessage = mMessage & "ɾ��" & mHtmlPicDir & "Ŀ¼�������ļ��ɹ���<br />"

		blnFlag = True
		ClearAllDiyPages = blnFlag
	End Function
	
	'����������¼��ļ���
	Public Function ReCreateAllDiyPages()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "��ʼ��������ȫ��DIYҳ��...<br />"
		Call ClearAllDiyPages()
		Call CreateAllDiyHtml()
		mMessage = mMessage & "��������ȫ��DIYҳ�����!<br />"

		blnFlag = True
		ReCreateAllDiyPages = blnFlag
	End Function
	
End Class
%>