<!--#Include File="Inc/Config.asp"-->
<!--#Include File="Inc/PageCode.asp"-->
<!--#Include File="Class/Cls_Template.asp"-->
<!--#Include File="Class/Cls_PageCode.asp"--><%
'==========================================
'文 件 名：Include.asp
'文件用途：前台总控文件
'版权所有：
'==========================================
If Fk_Site_Open=0 Then
	Call FKFun.ShowErr(Fk_Site_CloseStr,0)
End If
%>
