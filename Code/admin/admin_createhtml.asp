<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/class_createhtml.asp"-->
<!--#include file="../inc/class_template.asp"-->
<!--#include file="../inc/func_sitepath.asp"-->
<%
'ȫ�ֱ���
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim id: id = Req("id")
'��ǰҳ����
Dim Title: Title = ""
If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
Dim SitePath: SitePath = ""

Function CreateIndex()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("���ڴ���Index.html...<br />")
		If objCreate.CreateIndex()=True Then	'������ҳ
			Response.Write("����Index.html�ɹ���<br />")
		End If
	Set objCreate = Nothing
End Function


Function CreateAllArticles()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("�������´�����������ҳ...<br />")
		If objCreate.CreateAllArticles()=True Then	'����ȫ������
			Response.Write("������������ҳ�ɹ���<br />")
		End If
	Set objCreate = Nothing
End Function

Function CreateAllPictures()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		Response.Write("�������´�������ͼƬҳ...<br />")
		If objCreate.CreateAllArticles()=True Then	'����ȫ������
			Response.Write("��������ͼƬҳ�ɹ���<br />")
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtl()
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtml()=True Then	'����ȫ������
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtlByColID(colID)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHmtlByColID(ColID)=True Then	'����ȫ������
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtHmtlByID(id)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtmlByID(id)=True Then	'����ȫ������
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function

Function CreateArtListHmtlByID(id)
	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.CreateArtHtmlByID(id)=True Then	'����ȫ������
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing
End Function



	Dim objCreate
	Set objCreate = New ClassCreateHtml
		If objCreate.ReCreateAllPictures()=True Then	'����ȫ������
			Response.Write(objCreate.Message)
		End If
	Set objCreate = Nothing

'Call CreateArtHmtlByID(508)
'Call CreateArtListHmtlByID(6)
'Call CreateIndex()

'Dim i
'For i=2000 To 2500
'	Call DB("INSERT INTO Article(ColID,Title,Content,State) VALUES("& ((i Mod 5) + 1) &",'����׿�����ݹ���ϵͳ��JSCMS��"& i &"','����׿�����ݹ���ϵͳ��JSCMS��,���������ݣ�JSCMS���ݹ���ϵͳ��㵫���򵥣�<br>Foolin��Ʒ<br>�����ڴ���<br>�����̣�׿�����磨Joysyn.com��<br>������www.joysyn.com',1)", 0)
'
'Next
'	Response.Write("��ɣ�")
%>
