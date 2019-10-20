<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Index.asp
'文件用途：首页
'版权所有：
'==========================================
'定义变量
Dim PageCode,PageCodes,PageUrl,ModuleUrl
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode
Dim ct
If Fk_Site_Html=2 Then
	Response.Redirect(SiteDir&"Index.html")
ElseIf Fk_Site_Html=0 Then

	
	PageUrl=FKFun.HTMLEncode(Request.QueryString())
	
	'response.write("<hr/>")
	'City="henan"
	'response.write("PageUrl="&PageUrl&"<br/>")
	If Instr(PageUrl,"/")>0 Then
		ct=Split(PageUrl,"/")(0)
		'ct="hengshui"
		'response.write("ct="&ct&"<br/>")
		CityName = FKFun.GetCityName(ct)
		'response.write("CityName="&CityName&"<br/>")
		If CityName<>"" Then
			City=ct
			Fk_Site_Url= Fk_Site_Url & "?" &City & "/"
			Product_All_Url="?"&City&"/Product/"
			News_All_Url="?"&City&"/News/"
			Article_All_Url="?"&City&"/Article38/"
			'response.write("Fk_Site_Url="&Fk_Site_Url&"<br/>")
			PageUrl=Replace(PageUrl,ct&"/","")
		Else
			Fk_Site_Url ="www.hblhjgj.com"
			Product_All_Url="?Product/"
			News_All_Url="?News/"
			Article_All_Url="?Article38/"
			City=""
			CityName=""
		End If
	End If
	'response.write("<hr/>")
	
	If Instr(PageUrl,"&")>0 Then
		PageUrl=Split(PageUrl,"&")(0)
	End If
	PageUrl=Replace(PageUrl,"%2F","/")
	PageUrl=Replace(PageUrl,"%2f","/")
	If Fk_Site_Sign<>"" Then
		If Instr(PageUrl,Fk_Site_PageSign) Then
			PageUrl=PageUrl&FKTemplate.GetHtmlSuffix()
		End If
		PageUrl=Replace(PageUrl,Fk_Site_Sign,"/")
		PageUrl=Replace(PageUrl,Fk_Site_PageSign,"/Index_")
	End If
	PageNow=1
	If PageUrl="" Then
		MyModuleId=0
		MyId=0
		MyType=0
	Else
		If PageUrl<>"" And Instr(PageUrl,FKTemplate.GetHtmlSuffix())=0 And Right(PageUrl,1)<>"/" Then
			PageUrl=PageUrl&"/"
		End If
		If Instr(PageUrl,"/")>0 And Instr(PageUrl,FKTemplate.GetHtmlSuffix())>0 And Instr(PageUrl,"/Index")=0 Then
			TempArr=Split(PageUrl,"/")
			ModuleUrl=Replace(PageUrl,TempArr(UBound(TempArr)),"")
			MyId=Split(TempArr(UBound(TempArr)),".")(0)
			Sqlstr="Select Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_MUrl='"&ModuleUrl&"'"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				MyType=Rs("Fk_Module_Type")
				MyModuleId=Rs("Fk_Module_Id")
			Else
				Call FKFun.ShowErr("内容页面没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
			End If
			Rs.Close
			If Not IsNumeric(MyId) Then
				If MyType=1 Then
					Sqlstr="Select Fk_Article_Id From [Fk_Article] Where Fk_Article_Show=1 And Fk_Article_Module="&MyModuleId&" And Fk_Article_FileName='"&MyId&"'"
					Rs.Open Sqlstr,Conn,1,1
					If Not Rs.Eof Then
						MyId=Rs("Fk_Article_Id")
					Else
						Call FKFun.ShowErr("文章没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
					End If
					Rs.Close
				ElseIf MyType=2 Then
					Sqlstr="Select Fk_Product_Id From [Fk_Product] Where Fk_Product_Show=1 And Fk_Product_Module="&MyModuleId&" And Fk_Product_FileName='"&MyId&"'"
					Rs.Open Sqlstr,Conn,1,1
					If Not Rs.Eof Then
						MyId=Rs("Fk_Product_Id")
					Else
						Call FKFun.ShowErr("产品没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
					End If
					Rs.Close
				ElseIf MyType=7 Then
					Sqlstr="Select Fk_Down_Id From [Fk_Down] Where Fk_Down_Show=1 And Fk_Down_Module="&MyModuleId&" And Fk_Down_FileName='"&MyId&"'"
					Rs.Open Sqlstr,Conn,1,1
					If Not Rs.Eof Then
						MyId=Rs("Fk_Down_Id")
					Else
						Call FKFun.ShowErr("下载没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
					End If
					Rs.Close
				End If
			End If
		Else
			MyId=0
			If Instr(LCase(PageUrl),"/")>0 Then
				If Instr(LCase(PageUrl),"/index")>0 Then
					If Instr(LCase(PageUrl),"/index_")>0 Then
						TempArr=Split(PageUrl,"/")
						PageNow=Clng(Split(Split(LCase(PageUrl),"index_")(1),".")(0))
						ModuleUrl=Replace(PageUrl,TempArr(UBound(TempArr)),"")
					Else
						TempArr=Split(PageUrl,"/")
						ModuleUrl=Replace(PageUrl,TempArr(UBound(TempArr)),"")
					End If
				Else
					TempArr=Split(PageUrl,"/")
					ModuleUrl=Replace(PageUrl,TempArr(UBound(TempArr)),"")
				End If
			ElseIf Instr(LCase(PageUrl),"__")>0 Then
				PageNow=Clng(Split(Split(PageUrl,"__")(1),".")(0))
				ModuleUrl=Split(PageUrl,"__")(0)&FKTemplate.GetHtmlSuffix()
			Else
				ModuleUrl=PageUrl
			End If
			Sqlstr="Select Fk_Module_Id,Fk_Module_Type From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_MUrl='"&ModuleUrl&"'"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				MyType=Rs("Fk_Module_Type")
				MyModuleId=Rs("Fk_Module_Id")
			Else
				Call FKFun.ShowErr("内容页面没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
			End If
			Rs.Close
		End If
	End If
ElseIf Fk_Site_Html=1 Then
	MyModuleId=FKFun.GetNumeric("Module",0)
	MyId=FKFun.GetNumeric("Id",0)
	MyType=FKFun.GetNumeric("Type",0)
	PageNow=FKFun.GetNumeric("Page",1)
End If
If MyType=0 And MyModuleId=0 Then
	PageCode=FKPageCode.cIndex()
ElseIf MyModuleId>0 And MyId=0 Then
	PageCode=FKPageCode.cModule(MyModuleId,MyType)
ElseIf MyModuleId>0 And MyId>0 Then
	PageCode=FKPageCode.cPage(MyId,MyModuleId,MyType)
End If
Response.Write(PageCode)
%>
<!--#Include File="Code.asp"-->
