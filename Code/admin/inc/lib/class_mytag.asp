<%
'=========================================================
' Class Name：	ClassMyTag
' Purpose：		自定义标签
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-16 16:21:23
' Modify log:	
' Updated on: 	
' Copyright (c) 2010 卓新网络（JoySyn.com）All Rights Reserved
'=========================================================
Class ClassMyTag

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vName
	Private vInfo
	Private vCode
	'm前缀：Menber，类成员
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'Name
	Public Property Let Name(ByVal pName): vName = pName: End Property
	Public Property Get Name: Name = vName: End Property
	'Info
	Public Property Let Info(ByVal pInfo): vInfo = pInfo: End Property
	Public Property Get Info: Info = vInfo: End Property
	'Code
	Public Property Let Code(ByVal pCode): vCode = pCode: End Property
	Public Property Get Code: Code = vCode: End Property
	'LastError
	Public Property Let LastError(ByVal pLastError): mLastError = pLastError: End Property
	Public Property Get LastError: LastError = mLastError: End Property
	
	Private Sub Class_Initialize()
		Call ChkLogin()		'检查登录
		Call Initialize()	'初始化
	End Sub

	Private Sub Class_Terminate()
		Call Initialize()
	End Sub

	Public Function Initialize()
		vID = 0
		vName = ""
		vInfo = ""
		vCode = ""
		mLastError = ""
	End Function
	
	'--------------------------------------------------------------
	' Function name：	SetValue
	' Description: 		从表单获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function SetValue()
		vName = Request.Form("fName")
		vInfo = Request.Form("fInfo")
		vCode = Request.Form("fCode")
		If Len(vName) < 1 Or Len(vName) > 50 Then mLastError = "标题的长度请控制在 1 至 50 位" : SetValue = False : Exit Function
		If Len(vCode) = 0 Then mLastError = "标签代码不能为空" : SetValue = False : Exit Function
		If Len(vInfo) = 0 Then vInfo = ""
		SetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	LetValue
	' Description: 		从数据库获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function LetValue()
		Dim Rs
		Set Rs = DB("Select * From [MyTags] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
		vName = Rs("Name")
		vInfo = Rs("Info")
		vCode = Rs("Code")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Create()
	' Description: 		创建记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:40:46
	'--------------------------------------------------------------
	Public Function Create()
		If ExistTag(vName) = True Then mLastError = "标签已经[" & vName & "]已存在!": Create = False: Exit Function
		Dim Rs
		Set Rs = DB("Select * From [MyTags]",3)
		Rs.AddNew
		Rs("Name") = vName
		Rs("Info") = vInfo
		Rs("Code") = vCode
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Modify()
	' Description: 		修改记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [MyTags] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Modify = False : Exit Function
		If Rs("Name") <> vName Then
			If ExistTag(vName) = True Then mLastError = "标签已经[" & vName & "]已存在!":  Modify = False : Exit Function
		End If
		Rs("Name") = vName
		Rs("Info") = vInfo
		Rs("Code") = vCode
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Modify = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Delete()
	' Description: 		删除记录
	' Params: 			none
	' Return:			True
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Delete()
		DB "Delete From [MyTags] Where [ID] In (" & vID &")" ,0
		Delete = True
	End Function
	
	'判断标签是否存在
	Public Function ExistTag(Byval tagName)
		Dim Rs, Flag
		Set Rs = DB("Select [ID] From [MyTags] Where [Name]='" & tagName & "'",1)
		If Not Rs.Eof Then
			Flag = True
		Else
			Flag = False
		End If
		Rs.Close : Set Rs = Nothing
		ExistTag = Flag
	End Function

End Class
%>