<!--#Include File="../../Include.asp"--><%
'==========================================
'文 件 名：Plugin/Flash/Index.asp
'文件用途：Flash轮换
'版权所有：
'插件版本：V1.0.0
'==========================================

Dim Height,Width,Menu,Module
Dim Pic,Text,Url,ArticleUrl,ProductUrl,DownUrl
Dim TempTitle,TempMUrl,TempPic
Set FKTemplate=New Cls_Template
Dim TempDir
If Fk_Site_Dir<>"" Then
	TempDir=Fk_Site_Dir
Else
	TempDir=SiteDir
End If

'获取参数
Types=Clng(Request.QueryString("Type"))
Menu=Clng(Request.QueryString("Menu"))
Module=Clng(Request.QueryString("Module"))
Height=Request.QueryString("Height")
Width=Request.QueryString("Width")
If Height="" Then
	Height=100
Else
	Height=Clng(Height)
End If
If Width="" Then
	Width=100
Else
	Width=Clng(Width)
End If

Select Case Types
	Case 1
		Call Flash_1()
End Select

Sub Flash_1()
	Sqlstr="Select Fk_Module_Type From [Fk_Module] Where Fk_Module_Id="&Module&" And Fk_Module_Show=1"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Temp=Rs("Fk_Module_Type")
		If Instr(",1,2,7,",","&Temp&",")=0 Then
			Rs.Close
			Response.End()
		End If
	Else
		Rs.Close
		Response.End()
	End If
	Rs.Close
	Select Case Temp
		Case 1
			Sqlstr="Select Top 5 Fk_Article_Id,Fk_Article_Title,Fk_Article_FileName,Fk_Article_Pic,Fk_Module_Id,Fk_Module_Type,Fk_Module_MUrl From [Fk_ArticleList] Where (Fk_Article_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Article_Show=1 And Fk_Article_Pic<>'' Order By Fk_Article_Id Desc"
		Case 2
			Sqlstr="Select Top 5 Fk_Product_Id,Fk_Product_Title,Fk_Product_FileName,Fk_Product_Pic,Fk_Module_Id,Fk_Module_Type,Fk_Module_MUrl From [Fk_ProductList] Where (Fk_Product_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Product_Show=1 And Fk_Product_Pic<>'' Order By Fk_Product_Id Desc"
		Case 7
			Sqlstr="Select Top 5 Fk_Down_Id,Fk_Down_Title,Fk_Down_FileName,Fk_Down_Pic,Fk_Module_Id,Fk_Module_Type,Fk_Module_MUrl From [Fk_DownList] Where (Fk_Down_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Down_Show=1 And Fk_Down_Pic<>'' Order By Fk_Down_Id Desc"
	End Select
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		Select Case Temp
			Case 1
				TempTitle=Rs("Fk_Article_Title")
				If Rs("Fk_Article_FileName")<>"" Then
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Article_FileName")&FKTemplate.GetHtmlSuffix()
				Else
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Article_Id")&FKTemplate.GetHtmlSuffix()
				End If
				TempPic=Rs("Fk_Article_Pic")
			Case 2
				TempTitle=Rs("Fk_Product_Title")
				If Rs("Fk_Product_FileName")<>"" Then
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Product_FileName")&FKTemplate.GetHtmlSuffix()
				Else
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Product_Id")&FKTemplate.GetHtmlSuffix()
				End If
				TempPic=Rs("Fk_Product_Pic")
			Case 7
				TempTitle=Rs("Fk_Down_Title")
				If Rs("Fk_Down_FileName")<>"" Then
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Down_FileName")&FKTemplate.GetHtmlSuffix()
				Else
					TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&Rs("Fk_Down_Id")&FKTemplate.GetHtmlSuffix()
				End If
				TempPic=Rs("Fk_Down_Pic")
		End Select
		If Pic="" Then
			Pic=TempPic
		Else
			Pic=Pic&"|"&TempPic
		End If
		If Text="" Then
			Text=TempTitle
		Else
			Text=Text&"|"&TempTitle
		End If
		If Url="" Then
			Url=TempMUrl
		Else
			Url=Url&"|"&TempMUrl
		End If
		Rs.MoveNext
	Wend
	Rs.Close
%>
var focus_width=<%=Width%>     //场景宽
var focus_height=<%=Height%>　　//场景高
var text_height=0　　　//文字说明字高，为0时不显示文本
var swf_height = focus_height+text_height
var pics='<%=Pic%>'
var links='#'
var texts='<%=Text%>'
document.write('<object ID="focus_flash" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="<%=TempDir%>Plugin/Flash/1/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
document.write('<embed ID="focus_flash" src="<%=TempDir%>Plugin/Flash/1/focus.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" class=tablebody1 quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer"/>');  document.write('</object>');
<%
End Sub
%>
<!--#Include File="../../Code.asp"-->
