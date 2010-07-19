<%
'=========================================================
' Class Name：	ClassCreateHtml
' Purpose：		创建文件类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2010-7-17 22:15:28
' Modify log:	
' Updated on: 	
' Copyright (c) 2010 卓新网络（JoySyn.com）All Rights Reserved
'=========================================================
Class ClassCreateHtml

	Dim mTpl	'模板类实例
	Dim mHtmlDir
	Dim mHtmlArtDir
	Dim mHtmlPicDir
	Dim mHtmlDiyDir
	
	Private mMessage	'输出信息
	Public Property Get Message
		Message = mMessage
	End Property
	
	Private mErrMessage	'错误信息
	Public Property Get ErrMessage
		ErrMessage = mErrMessage
	End Property

	'初始化函数
	Private Sub Class_Initialize()
		Call ChkLogin()		'检查登录
		Set mTpl = New ClassTemplate	'创建模板
		mHtmlDir = "/Html/"	'Html目录
		mHtmlArtDir = mHtmlDir & "Article/"
		mHtmlPicDir = mHtmlDir & "Picture/"
		mHtmlDiyDir = mHtmlDir & "DiyPage/"
		If ExistFolder(mHtmlDir) = False Then
			Call CreateFolder(mHtmlDir)	'创建Html文件夹
		End If
		mMessage = ""
	End Sub
	
	'结束函数
	Private Sub Class_Terminate()
		Set mTpl = Nothing	'释放模板
	End Sub


	'创建首页
	Public Function CreateIndex()
		Dim blnFlag: blnFlag = False

		'定义全局 变量 {sys: var}
		page = 1
		id = colId
		Title = "文章列表": If id > 0 Then Title = GetNameOfColumn(id, "ARTICLE") & " - 文章列表"
		SitePath = ColPath(id, 0)
		
		'开始创建
		Dim tplIndexPath: tplIndexPath = mHtmlDir & "index.html"
		Call mTpl.LoadTpl("index.html")		'载入模板
		Call mTpl.Compile_Index()		'运行标签分析
		Call CreateFile(mTpl.Content, tplIndexPath)	'创建首页文件
		mMessage = mMessage & "创建首页index.html成功！<br />"
		blnFlag = True
		CreateIndex = blnFlag
	End Function
	
	'===================创建文章列表 ================
	
	'创建文章列表&文章
	Public Function CreateAllArticles()
		Dim blnFlag: blnFlag = False
		
		Dim strMessage
		mMessage = mMessage & "开始创建全部文章及列表页...<br />"
		'创建Html/Article文件夹
		Call CreateFolder(mHtmlArtDir)
		'创建所有栏目目录
		Call CreateAllArtFolder()
		'创建所有文章列表
		CreateAllArtListHtml()
		'创建所有文章
		CreateAllArtHtml()
		strMessage = mMessage & "创建全部文章及列表页成功！<br />"
		mMessage = strMessage
		
		blnFlag = True
		CreateAllArticles = blnFlag
	End Function
	
	'重新创建文章列表&文章
	Public Function ReCreateAllArticles()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "重新创建全部文章页...<br />"
		'清空Article目录及其下文件
		Call ClearAllArticles()
		'创建所有全部
		CreateAllArticles()
		mMessage = mMessage & "重新创建全部文章页完毕！<br />"
		
		blnFlag = True
		ReCreateAllArticles = blnFlag
	End Function
	
	
	'创建文章列表
	Public Function CreateAllArtListHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "开始创建全部文章列表页...<br />"
		'创建全部文章列表
		Call CreateArtListHtmlByID(0)
		'循环创建栏目文章
		Dim Rs
		Set Rs = DB( "SELECT ID FROM ArtColumn", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Call CreateArtListHtmlByID(Rs("ID"))
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建全部文章列表页成功！<br />"
		
		blnFlag = True
		CreateAllArtListHtml = blnFlag
	End Function
	
	'创建所有文章列表页
	Public Function CreateArtListHtmlByID(colId)
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "创建文章列表页...<br />"
		
		'定义列表相关变量
		Dim pageCount: pageCount = 1
		Dim listDir: listDir = mHtmlArtDir
		
		'定义全局 变量 {sys: var}
		page = 1
		id = colId
		Title = "文章列表": If id > 0 Then Title = GetNameOfColumn(id, "ARTICLE") & " - 文章列表"
		SitePath = ColPath(id, 0)	
		
		
		'循环创建文章
		If CInt(id) > 0 Then
			Dim Rs
			Set Rs = DB( "SELECT CreateHtmlPath FROM ArtColumn WHERE ID=" & id, 1)
			If Not Rs.Eof Then
				listDir = Rs("CreateHtmlPath")
			Else
				mErrMessage = mErrMessage & "不存在栏目ID为[" & id & "]的记录！"
				CreateArtListHtmlByID = False
			End If
		End If
		Do While True
			mTpl.Page = page						'设置当前页
			Call mTpl.LoadTpl("artlist.html")		'载入模板
			Call mTpl.Compile_List(id)				'运行标签分析
			pageCount = mTpl.PageCount
			Dim listPath: listPath = ""
			'创建List
			If id > 0 Then
				If Page>1 Then
					listPath = listDir & "list_" & id & "_" & page & ".html"	'某栏目文章列表，list_栏目ID_页码.html
				Else
					listPath = listDir & "list_" & id & ".html"	'某栏目文章列表，list_栏目ID.html 如果是第一页
				End If
			Else
				If Page>1 Then
					listPath = listDir & "list_" & page & ".html"	'全部文章列表，list_页码.html
				Else
					listPath = listDir & "list" & ".html"	'全部文章列表，list.html 如果是第一页
				End If
			End If
			Call CreateFile(mTpl.Content, listPath)
			mMessage = mMessage & "成功创建" & listPath & "!<br />"
			page = page + 1	'页数+1
			If Page>pageCount Then Exit Do '防上造成死循环
		Loop
		mMessage = mMessage & "创建文章列表完毕！<br />"
		
		blnFlag = True
		CreateArtListHtmlByID = blnFlag
	End Function
	
	
	'创建所有文章页
	Public Function CreateAllArtHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "创建全部文章...<br />"
			
		'循环创建文章
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "文章": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)
			
				'文章目录
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Call mTpl.LoadTpl("article.html")		'载入模板
				Call mTpl.Compile_Field(id, False)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strArtPath)			'输出内容
				mMessage = mMessage & "成功创建" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建文章完毕！<br />"
		
		blnFlag = True
		CreateAllArtHtml = blnFlag
	End Function
	
	'创建某栏目下的所有文章
	Public Function CreateArtHtmlByColID(colId)
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "创建文章...<br />"
		Dim colIds: colIds = GetColIds(colId, "article")	'获取所有子栏目ID及本身ID
		'循环创建文章
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.ColID IN (" & colIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "文章": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)	
			
				'文章目录
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Call mTpl.LoadTpl("article.html")		'载入模板
				Call mTpl.Compile_Field(id, False)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strArtPath)			'输出内容
				mMessage = mMessage & "成功创建" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建文章完毕！<br />"

		blnFlag = True
		CreateArtHtmlByColID = blnFlag
	End Function
	
	'创建指定ID的文章
	Public Function CreateArtHtmlByID(artIds)
		Dim blnFlag: blnFlag = False

		'创建文章页
		mMessage = mMessage & "创建文章...<br />"
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Article A Left Join ArtColumn B ON A.ColID=B.ID WHERE A.ID IN(" & artIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "文章": If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
				SitePath = ArtPath(id)	
			
				'文章目录
				Dim strArtPath
				If Rs("CreateHtmlPath")<>"" Then
					strArtPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strArtPath = mHtmlArtDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Call mTpl.LoadTpl("article.html")		'载入模板
				Call mTpl.Compile_Field(id, False)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strArtPath)			'输出内容
				mMessage = mMessage & "成功创建" & strArtPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建文章完毕！<br />"

		blnFlag = True
		CreateArtHtmlByID = blnFlag
	End Function
	
	
	'清空所有文章及文件夹
	Public Function ClearAllArticles()
		Dim blnFlag: blnFlag = False
		
		'如果存在该目录，则清空
		If ExistFolder(mHtmlArtDir) = True Then
			DeleteFolder(mHtmlArtDir)
		End If
		mMessage = mMessage & "已经删除" & mHtmlArtDir & "目录及该目录下的所有文件...<br />"

		blnFlag = True
		ClearAllArticles = blnFlag
	End Function
	
	
	
	'创建文章目录，第一级栏目分类
	Public Function CreateAllArtFolder()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "开始创建全部文章目录...<br />"
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
				'更新目录路径到数据库
				Call DB("Update ArtColumn set CreateHtmlPath='" & colDirPath & "' where ID=" & Rs("ID"), 0)
				'创建目录
				Call CreateFolder(colDirPath)
				mMessage = mMessage & "成功创建目录" & colDirPath & "<br />"
				Call P_CreateSubArtFolder(Rs("ID"), colDirPath) '循环子级分类
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建全部文章目录完毕！<br />"
		
		blnFlag = True
		CreateAllArtFolder = blnFlag
	End Function
	
	'子栏目分类
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
				mMessage = mMessage & "成功创建目录" & StrDis & "<br />"
				Call P_CreateSubArtFolder(Trim(Rs1("ID")), Strdis) '递归子级分类
			Rs1.Movenext:Loop
			If Rs1.Eof Then
				Rs1.Close: Set Rs1 = Nothing
				Exit Function
			End If
		End If
		Rs1.Close: Set Rs1 = Nothing
	End Function
	
	
	'===================创建图片列表 ================
	
	'创建图片列表&图片
	Public Function CreateAllPictures()
		Dim blnFlag: blnFlag = False
		
		Dim strMessage
		strMessage = strMessage & "开始创建全部图片页..."
		'创建Html文件夹
		Call CreateFolder(mHtmlPicDir)	
		'循环创建目录
		Call CreateAllPicFolder()	
		'创建图片列表
		CreateAllPicListHtml()	
		'创建图片
		CreateAllPicHtml()
		strMessage = mMessage & "创建全部图片页成功！<br />"

		mMessage = strMessage

		blnFlag = True
		CreateAllPictures = blnFlag
	End Function
	
	'重新创建文章列表&文章
	Public Function ReCreateAllPictures()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "重新创建全部图片页...<br />"
		'清空Article目录及其下文件
		Call ClearAllPictures()
		'创建所有全部
		CreateAllPictures()
		mMessage = mMessage & "重新创建全部文章页完毕！<br />"
		
		blnFlag = True
		ReCreateAllPictures = blnFlag
	End Function
	
	'创建图片列表
	Public Function CreateAllPicListHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "开始创建全部图片列表页...<br />"
		'创建默认所有图片列表
		Call CreatePicListHtmlByID(0)
		'循环创建栏目图片
		Dim Rs
		Set Rs = DB( "SELECT ID FROM PicColumn", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Call CreatePicListHtmlByID(Rs("ID"))
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建全部图片列表页完毕！<br />"
		
		blnFlag = True
		CreateAllPicListHtml = blnFlag
	End Function
	
	'创建所有图片列表页
	Public Function CreatePicListHtmlByID(colId)
		Dim blnFlag: blnFlag = False
		'定义列表相关变量
		Dim pageCount: pageCount = 1
		Dim listDir: listDir = mHtmlPicDir
		
		'定义全局 变量 {sys: var}
		page = 1
		id = colId
		Title = "图片列表": If id > 0 Then Title = GetNameOfColumn(id, "PICTURE") & " - 图片列表"
		SitePath = ColPath(id, 0)	
		
		
		'循环创建图片
		If CInt(id) > 0 Then
			Dim Rs
			Set Rs = DB( "SELECT CreateHtmlPath FROM PicColumn WHERE ID=" & id, 1)
			If Not Rs.Eof Then
				listDir = Rs("CreateHtmlPath")
			Else
				mErrMessage = mErrMessage & "不存在栏目ID为[" & id & "]的记录！<br />"
				CreatePicListHtmlByID = False
			End If
		End If
		Do While True
			mTpl.Page = page						'设置当前页
			Call mTpl.LoadTpl("piclist.html")		'载入模板
			Call mTpl.Compile_List(id)				'运行标签分析
			pageCount = mTpl.PageCount
			Dim listPath: listPath = ""
			'创建List
			If id > 0 Then
				If Page>1 Then
					listPath = listDir & "list_" & id & "_" & page & ".html"	'某栏目图片列表，list_栏目ID_页码.html
				Else
					listPath = listDir & "list_" & id & ".html"	'某栏目图片列表，list_栏目ID.html 如果是第一页
				End If
			Else
				If Page>1 Then
					listPath = listDir & "list_" & page & ".html"	'全部图片列表，list_页码.html
				Else
					listPath = listDir & "list" & ".html"	'全部图片列表，list.html 如果是第一页
				End If
			End If
			Call CreateFile(mTpl.Content, listPath)
			mMessage = mMessage & "成功创建" & listPath & "!<br />"
			page = page + 1	'页数+1
			If Page>pageCount Then Exit Do '防上造成死循环
		Loop
		
		blnFlag = True
		CreatePicListHtmlByID = blnFlag
	End Function
	
	
	'创建所有图片页
	Public Function CreateAllPicHtml()
		Dim blnFlag: blnFlag = False
		mMessage = mMessage & "创建图片...<br />"
			
		'循环创建图片
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "图片": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)
				

			
				'图片目录
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'载入模板
				Call mTpl.Compile_Field(id, True)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strPicPath)			'输出内容
				mMessage = mMessage & "成功创建" & strPicPath & "<br />"

			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建图片完毕！<br />"

		blnFlag = True
		CreateAllPicHtml = blnFlag
	End Function
	
	'创建某栏目下的所有图片
	Public Function CreatePicHtmlByColID(colId)
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "创建图片...<br />"
		Dim colIds: colIds = GetColIds(colId, "picture")	'获取所有子栏目ID及本身ID
		'循环创建图片
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.ColID IN (" & colIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "图片": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)	
			
				'图片目录
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'载入模板
				Call mTpl.Compile_Field(id, True)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strPicPath)			'输出内容
				mMessage = mMessage & "成功创建" & strPicPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建图片完毕！<br />"

		blnFlag = True
		CreatePicHtmlByColID = blnFlag
	End Function
	
	'创建指定ID的图片
	Public Function CreatePicHtmlByID(picIds)
		Dim blnFlag: blnFlag = False

		'创建图片
		mMessage = mMessage & "创建图片...<br />"
		Dim Rs
		Set Rs = DB( "SELECT A.*,B.CreateHtmlPath FROM Picture A Left Join PicColumn B ON A.ColID=B.ID WHERE A.ID IN(" & picIds & ") AND A.State=1", 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				id = Rs("ID")
				Title = "图片": If id > 0 Then Title = GetTitleOfArtOrPic(id, "PICTURE")
				SitePath = PicPath(id)	
			
				'图片目录
				Dim strPicPath
				If Rs("CreateHtmlPath")<>"" Then
					strPicPath = Rs("CreateHtmlPath") & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				Else
					strPicPath = mHtmlPicDir & "html_" & Rs("ColID") & "_" & Rs("ID") & ".html"
				End If
				'分析内容
				Dim id: id = Rs("ID")
				Call mTpl.LoadTpl("picture.html")		'载入模板
				Call mTpl.Compile_Field(id, True)	'运行标签分析
				'创建文件
				Call CreateFile(mTpl.Content, strPicPath)			'输出内容
				mMessage = mMessage & "成功创建" & strPicPath & "<br />"
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建图片完毕！<br />"

		blnFlag = True
		CreatePicHtmlByID = blnFlag
	End Function


	
	'清空所有文章及文件夹
	Public Function ClearAllPictures()
		Dim blnFlag: blnFlag = False
		
		'如果存在该目录，则清空
		If ExistFolder(mHtmlPicDir) = True Then
			DeleteFolder(mHtmlPicDir)
		End If
		mMessage = mMessage & "删除" & mHtmlPicDir & "目录及其下文件成功！<br />"

		blnFlag = True
		ClearAllPictures = blnFlag
	End Function
	
	'创建图片目录，第一级栏目
	Public Function CreateAllPicFolder()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "开始创建全部图片目录...<br />"
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
				'更新目录路径到数据库
				Call DB("Update PicColumn set CreateHtmlPath='" & colDirPath & "' where ID=" & Rs("ID"), 0)
				'创建目录
				Call CreateFolder(colDirPath)
				mMessage = mMessage & "创建目录" & colDirPath & "<br />"
				'循环子级分类
				Call P_CreateSubPicFolder(Rs("ID"), colDirPath) 
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "开始创建全部图片目录成功!<br />"
		
		blnFlag = True
		CreateAllPicFolder = blnFlag
	End Function
	
	'子栏目
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
				'更新目录到数据库
				Call DB("Update PicColumn set CreateHtmlPath='" & StrDis & "' where ID=" & Rs1("ID"), 0)
				'创建目录
				Call CreateFolder(StrDis)
				mMessage = mMessage & "创建目录" & StrDis & "<br />"
				Call P_CreateSubPicFolder(Trim(Rs1("ID")), Strdis) '递归子级分类
			Rs1.Movenext:Loop
			If Rs1.Eof Then
				Rs1.Close: Set Rs1 = Nothing
				Exit Function
			End If
		End If
		Rs1.Close: Set Rs1 = Nothing
	End Function
	
	
	'===================创建自定义页面 ================
	
	'创建自定义页面
	Public Function CreateAllDiyHtml()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "开始创建全部Diy页面...<br />"
		'创建DIY页面
		Dim Rs,strSql
		strSql = "SELECT * FROM DiyPage WHERE State = 1"
		Set Rs = DB(strSql, 1)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
			
				'定义全局 变量 {sys: var}
				page = 1
				'id = Rs("ID")
				Title = "DIY页面": Title = GetTitleOfDiypage(Rs("ID"))
				SitePath = DiyPagePath(Rs("ID"))
				
				Dim strDiyPath
				If Rs("PageName")<>"" Then
					strDiyPath = mHtmlDiyDir & Rs("PageName")
				Else
					strDiyPath = mHtmlDiyDir & "html_" & Rs("ID") & ".html"
				End If
				mTpl.Page = page						'设置当前页
				Call mTpl.Compile_DiyPage(Rs("ID"))		'运行标签分析
				Call CreateFile(mTpl.Content, strDiyPath)
				mMessage = mMessage & "成功创建DIY页面" & strDiyPath & "！<br />"
				
			Rs.MoveNext
			If Rs.Eof Then Exit Do '防上造成死循环
			Loop
		Else
			mErrMessage = mErrMessage & "创建失败！不存在DIY页面[" & param & "]的记录！<br />"
		End If
		Rs.Close: Set Rs = Nothing
		mMessage = mMessage & "创建全部Diy页面!<br />"
		
		blnFlag = True
		CreateAllDiyHtml = blnFlag
	End Function
	
	Public Function CreateDiyHtmlByParam(param)
		Dim blnFlag: blnFlag = False
		If Len(param)=0 Then mErrMassage = "调用CreateDiyHtmlByID(param)函数的param不能为空！": Exit Function
		
		'定义全局 变量 {sys: var}
		page = 1
		'id = Rs("ID")
		Title = "DIY页面": Title = GetTitleOfDiypage(param)
		SitePath = DiyPagePath(param)
		
		
		'创建DIY页面
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
			mTpl.Page = page						'设置当前页
			Call mTpl.Compile_DiyPage(param)			'运行标签分析
			Call CreateFile(mTpl.Content, strDiyPath)
			mMessage = mMessage & "成功创建DIY页面" & strDiyPath & "！<br />"
		Else
			mErrMessage = mErrMessage & "创建失败！不存在DIY页面[" & param & "]的记录！<br />"
		End If
		Rs.Close: Set Rs = Nothing
		
		blnFlag = True
		CreateDiyHtmlByParam = blnFlag
	End Function
	
	'清空所有文章及文件夹
	Public Function ClearAllDiyPages()
		Dim blnFlag: blnFlag = False
		
		'如果存在该目录，则清空
		If ExistFolder(mHtmlDiyDir) = True Then
			Call DeleteFolder(mHtmlDiyDir)
		End If
		mMessage = mMessage & "删除" & mHtmlPicDir & "目录及其下文件成功！<br />"

		blnFlag = True
		ClearAllDiyPages = blnFlag
	End Function
	
	'清空所有文章及文件夹
	Public Function ReCreateAllDiyPages()
		Dim blnFlag: blnFlag = False
		
		mMessage = mMessage & "开始重新生成全部DIY页面...<br />"
		Call ClearAllDiyPages()
		Call CreateAllDiyHtml()
		mMessage = mMessage & "重新生成全部DIY页面完毕!<br />"

		blnFlag = True
		ReCreateAllDiyPages = blnFlag
	End Function
	
End Class
%>