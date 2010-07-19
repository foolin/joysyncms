<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/class_createhtml.asp"-->
<!--#include file="../inc/class_template.asp"-->
<!--#include file="../inc/func_sitepath.asp"-->
<%
'全局变量
Dim page: page = CPage(Req("page"))	'当前页数
Dim id: id = Req("id")
'当前页标题
Dim Title: Title = ""
If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
Dim SitePath: SitePath = ""

Function CreateIndex()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("正在创建Index.html...<br />")
		If objCreate.CreateIndex()=True Then	'创建首页
			Response.Write("创建Index.html成功！<br />")
		End If
	Set objCreate = Nothing
End Function


Function CreateAllArticles()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("正在重新创建所有文章页...<br />")
		If objCreate.CreateAllArticles()=True Then	'生成全部文章
			Response.Write("创建所有文章页成功！<br />")
		End If
	Set objCreate = Nothing
End Function

Function CreateAllPictures()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("正在重新创建所有图片页...<br />")
		If objCreate.CreateAllArticles()=True Then	'生成全部文章
			Response.Write("创建所有图片页成功！<br />")
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtl()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtml()=True Then	'生成全部文章
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtlByColID(colID)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHmtlByColID(ColID)=True Then	'生成全部文章
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtlByID(id)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtmlByID(id)=True Then	'生成全部文章
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtListHmtlByID(id)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtmlByID(id)=True Then	'生成全部文章
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function



	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.ReCreateAllPictures()=True Then	'生成全部文章
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing

'Call CreateArtHmtlByID(508)
'Call CreateArtListHmtlByID(6)
'Call CreateIndex()

'Dim i
'For i=2000 To 2500
'	Call DB("INSERT INTO Article(ColID,Title,Content,State) VALUES("& ((i Mod 5) + 1) &",'测试卓新内容管理系统（JSCMS）"& i &"','测试卓新内容管理系统（JSCMS）,这里是内容，JSCMS内容管理系统简便但不简单！<br>Foolin作品<br>敬请期待！<br>开发商：卓新网络（Joysyn.com）<br>官网：www.joysyn.com',1)", 0)
'
'Next
'	Response.Write("完成！")
%>
