<!--#Include File="../Inc/Config.asp"-->
<!--#Include File="../Class/Cls_Admin.asp"-->
<!--#Include File="../Class/Cls_Html.asp"-->
<!--#Include File="../Class/Cls_PageCode.asp"-->
<!--#Include File="../Class/Cls_Template.asp"--><%
'==========================================
'文 件 名：Admin/Include.asp
'文件用途：管理员总控
'版权所有：方卡在线
'==========================================
'类赋值，判断管理权限
Set FKAdmin=New Cls_Admin
Set FKHtml=New Cls_Html
Set FKTemplate=New Cls_Template
Login=FKAdmin.AdminCheck(1,"","")
%>