<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：GBookDo.asp
'文件用途：留言提交
'版权所有：
'==========================================

'定义页面变量
Dim Fk_GBook_Module,Fk_Module_GModel,Fk_Module_GUrl,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MinStr,Fk_GModel_MaxStr,Fk_GBook_Content,Fk_GModel_Content,Fk_GModel_GoUrl,Fk_GModel_Contents
Dim TempArr3

'获取功能选项参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GBookAddDo() '添加留言
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==============================
'函 数 名：GBookAddDo
'作    用：添加留言
'参    数：
'==============================
Sub GBookAddDo()
	Dim Str1,Str2
	Fk_GBook_Module=Trim(Request.Form("Fk_GBook_Module"))
	Call FKFun.AlertNum(Fk_GBook_Module,"留言板系统参数错误，可能是模板配置问题！")
	Sqlstr="Select Fk_Module_GModel,Fk_Module_GUrl From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Type=4 And Fk_Module_Id=" & Fk_GBook_Module
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_GModel=Rs("Fk_Module_GModel")
		Fk_Module_GUrl=Rs("Fk_Module_GUrl")
	Else
		Rs.Close
		Call FKFun.AlertInfo("模块不存在！",SiteDir)
	End If
	Rs.Close
	Sqlstr="Select Fk_GModel_Content,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MinStr,Fk_GModel_MaxStr,Fk_GModel_GoUrl From [Fk_GModel] Where Fk_GModel_Id=" & Fk_Module_GModel
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		TempArr=Split(Rs("Fk_GModel_Content"),"|-_-|Fangka|-_-|")
		Fk_GModel_GoUrl=Rs("Fk_GModel_GoUrl")
		Fk_GModel_Succeed=Rs("Fk_GModel_Succeed")
		Fk_GModel_Repeat=Rs("Fk_GModel_Repeat")
		Fk_GModel_NoTrash=Rs("Fk_GModel_NoTrash")
		Fk_GModel_TrashPint=Rs("Fk_GModel_TrashPint")
		Fk_GModel_MinStr=Rs("Fk_GModel_MinStr")
		Fk_GModel_MaxStr=Rs("Fk_GModel_MaxStr")
		If Fk_GModel_GoUrl="" Or IsNull(Fk_GModel_GoUrl) Then
			Fk_GModel_GoUrl=SiteDir
		End If
		If Fk_Module_GUrl<>"" Then
			Fk_GModel_GoUrl=Fk_Module_GUrl
		End If
	Else
		Rs.Close
		Call FKFun.AlertInfo("留言模型不存在！",SiteDir)
	End If
	Rs.Close
	For i=0 To UBound(TempArr)
		If Instr(TempArr(i),"|-^-|Fangka|-^-|") Then
			TempArr3=Split(TempArr(i),"|-^-|Fangka|-^-|")
			Temp=FKFun.HTMLEncode(Trim(Request.Form(TempArr3(3))))
			Str1=Replace(Replace(Fk_GModel_MinStr,"{-留言条目-}",TempArr3(0)),"{-留言长度-}",TempArr3(1))
			Str2=Replace(Replace(Fk_GModel_MaxStr,"{-留言条目-}",TempArr3(0)),"{-留言长度-}",TempArr3(2))
			Call FKFun.AlertString(Temp,Clng(TempArr3(1)),Clng(TempArr3(2)),0,Str1,Str2)
			If Fk_GBook_Content="" Then
				Fk_GBook_Content=TempArr3(0)&"|-^-|Fangka|-^-|"&Temp&"|-^-|Fangka|-^-|"&TempArr3(3)
			Else
				Fk_GBook_Content=Fk_GBook_Content&"|-_-|Fangka|-_-|"&TempArr3(0)&"|-^-|Fangka|-^-|"&Temp&"|-^-|Fangka|-^-|"&TempArr3(3)
			End If
			Fk_GModel_Contents=Fk_GModel_Contents&TempArr3(0)&"："&Temp&"<br />"
		End If
	Next
	If Fk_GModel_NoTrash>0 Then
		Call FKFun.NoTrash(Fk_GBook_Content,Fk_GModel_NoTrash,Fk_GModel_TrashPint)
	End If
	Sqlstr="Select Fk_GBook_Id,Fk_GBook_Content,Fk_GBook_Module,Fk_GBook_Ip From [Fk_GBook] Where Fk_GBook_Module="&Fk_GBook_Module&" And Fk_GBook_Content='"&Fk_GBook_Content&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_GBook_Content")=Fk_GBook_Content
		Rs("Fk_GBook_Module")=Fk_GBook_Module
		Rs("Fk_GBook_Ip")=Request.ServerVariables("REMOTE_ADDR")
		Rs.Update()
		Application.UnLock()
		Rs.Close
		If Fk_Site_Mail=1 Then
			Temp=FKFun.Jmail(Mail_Address,"“"&Fk_Site_Name&"”新信息提醒","“"&Fk_Site_Name&"”有新信息，请注意及时处理！<br />"&Fk_GModel_Contents,"GB2312","text/html") 
		End If
		Call FKFun.AlertInfo(Fk_GModel_Succeed,Fk_GModel_GoUrl)
	Else
		Rs.Close
		Call FKFun.AlertInfo(Fk_GModel_Repeat,"1")
	End If
End Sub
%>
<!--#Include File="Code.asp"-->
