<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'==========================================
'文 件 名：Inc/Config.asp
'文件用途：系统配置
'版权所有：
'==========================================
Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
Response.Expires=-999
Session.Timeout=999
Dim StartTime,EndTime
StartTime=Timer()
%>
<!--#Include File="System.asp"-->
<!--#Include File="../Class/Cls_DB.asp"-->
<!--#Include File="../Class/Cls_Fso.asp"-->
<!--#Include File="../Class/Cls_Fun.asp"-->
<%

'定义页面常量
Dim Id,i,j,Types
Dim Temp,TempArr,Login,UrlIndexIn
Dim Fso,F,objAdoStream
Dim Conn,Rs,Sqlstr
Dim TemplateTempArr,TemplateTemp,Templates,NoDirStr
Dim SearchStr,SearchType,SearchTemplate,SearchField,SearchFieldList
Dim PageNow,PageCounts,PageSizes,PageAll,PageFirst,PageNext,PagePrev,PageLast,TempPageSize
Dim Fk_Site_Name,Fk_Site_Url,Fk_Site_Keyword,Fk_Site_Description,Fk_Site_Open,Fk_Site_CloseStr,Fk_Site_Template,Fk_Site_Html,Fk_Site_HtmlType,Fk_Site_HtmlSuffix,Fk_Site_PageSize,Fk_Site_ToPinyin,Fk_Site_DelWord,Fk_Site_SkinTest,Fk_Site_Dir,Fk_Site_Index,Fk_Site_FetionNum,Fk_Site_FetionPass,Fk_Site_LinkExt,Fk_Site_PicExt,Fk_Site_FlashExt,Fk_Site_MediaExt,Fk_Site_Jpeg,Fk_Site_Edit,Fk_Site_Mail,Fk_Site_MailStr,Fk_Site_Field,Fk_Site_Sign,Fk_Site_PageSign,Fk_Site_SysHidden
Dim Jpeg_Pic,Jpeg_Pic_w,Jpeg_Pic_h,Jpeg_EditPic,Jpeg_EditPic_w,Jpeg_EditPic_h,Jpeg_Water,Jpeg_WaterText,Jpeg_WaterFontSize,Jpeg_WaterFont,Jpeg_WaterFontColor,Jpeg_WaterFontWeight,Jpeg_WaterPic,Jpeg_WaterPicTransparence,Jpeg_WaterPicBgColor,Jpeg_WaterPosition,Jpeg_Water_x,Jpeg_Water_y
Dim EditorClass,FileDir
Dim MenuTepmlate,IsIndex,t_MenuIds
Dim Mail_Address,Mail_Name,Mail_Pass,Mail_Smtp
Dim FKDB,FKFun,FKFso,FKHtml,FKTemplate,FKJpeg,FKAdmin,FKPageCode
Dim FKModuleId,FKModuleName
Dim SiteDir,SiteDBDir,SiteData
Dim MyModuleId,MyId,MyType
Dim City,CityName
Dim Product_All_Url
Dim News_All_Url
Dim Article_All_Url
%>
<!--#Include File="Conn.asp"-->
<!--#Include File="Md5.asp"-->
<!--#Include File="Model.asp"-->
<%

'置默认值
Product_All_Url="?Product/"
News_All_Url="?News/"
Article_All_Url="?Article38/"
City=""
CityName=""
Set FKDB=New Cls_DB
Set FKFun=New Cls_Fun
Set FKFso=New Cls_Fso
Call FKDB.DB_Open()
NoDirStr="Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','gbookshow','page','subject','job','subject','top','bottom','downlist','down','search') And Not (Fk_Template_Name Like '%%/index' Or Fk_Template_Name Like '%%/info' Or Fk_Template_Name Like '%%/articlelist' Or Fk_Template_Name Like '%%/article' Or Fk_Template_Name Like '%%/productlist' Or Fk_Template_Name Like '%%/product' Or Fk_Template_Name Like '%%/gbook' Or Fk_Template_Name Like '%%/gbookshow' Or Fk_Template_Name Like '%%/page' Or Fk_Template_Name Like '%%/job' Or Fk_Template_Name Like '%%/top' Or Fk_Template_Name Like '%%/bottom' Or Fk_Template_Name Like '%%/downlist' Or Fk_Template_Name Like '%%/down')"
Sqlstr="Select * From [Fk_Site]"
Rs.Open Sqlstr,Conn,1,1
If Not Rs.Eof Then
	Fk_Site_Name=Rs("Fk_Site_Name")
	Fk_Site_Url=Rs("Fk_Site_Url")
	Fk_Site_Keyword=Rs("Fk_Site_Keyword")
	Fk_Site_Description=Rs("Fk_Site_Description")
	Fk_Site_Open=Rs("Fk_Site_Open")
	Fk_Site_CloseStr=Rs("Fk_Site_CloseStr")
	dim MoblieUrl,reExp,MbStr
Set reExp = New RegExp
MbStr="Android|iPhone|UC|Windows Phone|webOS|BlackBerry|iPod"
reExp.pattern=".*("&MbStr&").*"
reExp.IgnoreCase = True
reExp.Global = True
If reExp.test(Request.ServerVariables("HTTP_USER_AGENT")) Then
	Fk_Site_Template="mobi"
Else
	Fk_Site_Template=Rs("Fk_Site_Template")
End If
	Fk_Site_Html=Rs("Fk_Site_Html")
	Fk_Site_HtmlType=Rs("Fk_Site_HtmlType")
	Fk_Site_HtmlSuffix=Rs("Fk_Site_HtmlSuffix")
	Fk_Site_PageSize=Rs("Fk_Site_PageSize")
	Fk_Site_ToPinyin=Rs("Fk_Site_ToPinyin")
	Fk_Site_DelWord=Rs("Fk_Site_DelWord")
	Fk_Site_SkinTest=Rs("Fk_Site_SkinTest")
	Fk_Site_Dir=Rs("Fk_Site_Dir")
	Fk_Site_Index=Rs("Fk_Site_Index")
	Fk_Site_LinkExt=Rs("Fk_Site_LinkExt")
	Fk_Site_PicExt=Rs("Fk_Site_PicExt")
	Fk_Site_FlashExt=Rs("Fk_Site_FlashExt")
	Fk_Site_MediaExt=Rs("Fk_Site_MediaExt")
	Fk_Site_Jpeg=Rs("Fk_Site_Jpeg")
	Fk_Site_Edit=Rs("Fk_Site_Edit")
	Fk_Site_Mail=Rs("Fk_Site_Mail")
	Fk_Site_MailStr=Rs("Fk_Site_MailStr")
	Fk_Site_Sign=Rs("Fk_Site_Sign")
	Fk_Site_PageSign=Rs("Fk_Site_PageSign")
	Fk_Site_SysHidden=Rs("Fk_Site_SysHidden")
	If IsNull(Rs("Fk_Site_Field")) Or Rs("Fk_Site_Field")="" Then
		Fk_Site_Field=Split("-_-|-Fangka_Field-|1")
	Else
		Fk_Site_Field=Split(Rs("Fk_Site_Field"),"[-Fangka_Field-]")
	End If
	If Fk_Site_MailStr<>"" Then
		If UBound(Split(Fk_Site_MailStr,"||"))=3 Then
			TempArr=Split(Fk_Site_MailStr,"||")
			Mail_Address=TempArr(0)
			Mail_Name=TempArr(1)
			Mail_Pass=TempArr(2)
			Mail_Smtp=TempArr(3)
		End If
	End If
	Select Case Fk_Site_Edit
		Case 0
			EditorClass="Editer"
		Case 1
			EditorClass="KinEditer"
		Case 2
			EditorClass="UEditor"
	End Select
	If Fk_Site_Index=0 Then
		UrlIndexIn=""
	Else
		UrlIndexIn="Index.asp"
	End If
	FileDir=SiteDir
	If Fk_Site_Dir<>"" Then
		SiteDir=Fk_Site_Dir
	End If
Else
	Call FKFun.ShowErr("站点配置读取失败！",0)
End If
Rs.Close
%>
