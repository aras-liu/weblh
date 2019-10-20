<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Ajax.asp
'文件用途：信息切换拉取页面
'版权所有：方卡在线
'==========================================

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call Menu_Module() '读取顶部菜单
	Case 2
		Call GetKeyword() '获取关键字
	Case 3
		Call GetDescription() '获取描述
End Select

'==============================
'函 数 名：ShowModuleSelect
'作    用：输出ModuleSelect列表
'参    数：要输出的菜单MenuIds
'==============================
Public Function Menu_Module()
	Id=Clng(Request.QueryString("Id"))
	If Request.QueryString("Temp")<>1 Then
		Response.Write("0|||||一级模块,,,,,")
		Response.Write("{$ModuleId$}|||||当前模块,,,,,")
	End If
	Call ShowModuleSelectM(Id,0,"")
End Function
Public Function ShowModuleSelectM(MenuIds,LevelId,TitleBack)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,1
	While Not Rs2.Eof
		Response.Write(Rs2("Fk_Module_Id")&"|||||"&TitleBack&Rs2("Fk_Module_Name")&"["&FKFun.CheckModule(Rs2("Fk_Module_Type"))&"],,,,,")
		If LevelId=0 Then
			TitleBacks="   ├"
		Else
			TitleBacks="   "&TitleBack
		End If
		Call ShowModuleSelectM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks)
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：GetKeyword
'作    用：获取关键字
'参    数：
'==============================
Public Function GetKeyword()
	Dim KeyWord,C1,C2,Content
	Id=Clng(Request.QueryString("Id"))
	Select Case Id
		Case 1
			Content=Request.Form("Fk_Article_Content")
		Case 2
			Content=Request.Form("Fk_Product_Content")
		Case 3
			Content=Request.Form("Fk_Down_Content")
		Case Else
			Response.End()
	End Select
	KeyWord=Trim(FKFun.UnEscape(FKFso.FsoFileRead("KeyWord.dat")))
	If KeyWord<>"" And Content<>"" Then
		TempArr=Split(KeyWord," ")
		KeyWord=""
		C1=Len(Content)
		For Each Temp In TempArr
			If Temp<>"" Then
				C2=(C1-Len(Replace(Content,Temp,"")))/C1
				If C2>=0.02 And C2<=0.08 Then
					If Len(KeyWord)<90 Then
						If KeyWord="" Then
							KeyWord=Temp
						Else
							KeyWord=KeyWord&","&Temp
						End If
					End If
				End If
			End If
		Next
		Response.Write(KeyWord)
	End If
End Function

'==============================
'函 数 名：GetDescription
'作    用：获取描述
'参    数：
'==============================
Public Function GetDescription()
	Dim Content
	Id=Clng(Request.QueryString("Id"))
	Select Case Id
		Case 1
			Content=Request.Form("Fk_Article_Content")
		Case 2
			Content=Request.Form("Fk_Product_Content")
		Case 3
			Content=Request.Form("Fk_Down_Content")
		Case Else
			Response.End()
	End Select
	Response.Write(Left(Replace(FKFun.RemoveHTML(Content)," ",""),100))
End Function
%>