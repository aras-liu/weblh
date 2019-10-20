<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Inc/PageCode.asp"--><%
'==========================================
'文 件 名：Admin/Html.asp
'文件用途：HTML生成拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System3",Request.Cookies("FkAdminLimit1"))

'定义常量
Set FKHtml=New Cls_Html
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode
Dim t_PageSizes,t_i,t_gList

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call HtmlBox() '读取HTML生成器
	Case 2
		Call HtmlCreate() '生成常规HTML
	Case 3
		Call HtmlSelectCreate() '生成选择项HTML
	Case 4
		Call HtmlDayCreate() '一键生成HTML
End Select

'==========================================
'函 数 名：HtmlBox()
'作    用：读取HTML生成器
'参    数：
'==========================================
Sub HtmlBox()
%>
<div id="BoxTop" style="width:700px;">HTML生成器[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
<%
	If Fk_Site_Html=2 Then
%>
<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=1';">生成首页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=2';">生成所有信息页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=3';">生成所有静态页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=4';">生成所有留言页</a></td>
        </tr>
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=5';">生成招聘页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=6';">生成所有专题页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=7';">生成所有文章页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=8';">生成所有产品页</a></td>
    </tr>
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=9';">生成所有下载页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=2&Id=10';">生成所有跳转页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=4&CreatDay=1';">一键今日更新</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Html.asp?Type=4&CreatDay=2';">一键2日内更新</a></td>
    </tr>
    <tr>
        <td height="30" align="right">单独模块生成：</td>
        <td height="30" colspan="3" style="padding:5px;">
        <select name="MenuId" class="Input" id="MenuId" onchange="ChangeSelect('Ajax.asp?Type=1&Temp=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId');">
            <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>
            <select name="ModuleId" class="Input" id="ModuleId">
                <option value="">请先选择菜单</option>
                </select><br />
            <select name="CreatType" class="Input" id="CreatType">
                <option value="0">生成分类页+内容页</option>
                <option value="1">只生成分类页</option>
                <option value="2">只生成内容页</option>
                </select>
            <select name="CreatDay" class="Input" id="CreatDay">
                <option value="0">生成所有</option>
                <option value="1">生成1天内</option>
                <option value="2">生成2天内</option>
                <option value="7">生成7天内</option>
                </select>
            <input type="button" onclick="document.getElementById('Gets').src='Html.asp?Type=3&MenuId='+document.all.MenuId.options[document.all.MenuId.selectedIndex].value+'&ModuleId='+document.all.ModuleId.options[document.all.ModuleId.selectedIndex].value+'&CreatType='+document.all.CreatType.options[document.all.CreatType.selectedIndex].value+'&CreatDay='+document.all.CreatDay.options[document.all.CreatDay.selectedIndex].value;" class="Button" name="button2" id="button2" value="生 成" />
            </td>
        </tr>
    <tr>
        <td height="25" colspan="4" align="center">&nbsp;&nbsp;标签生成结果</td>
    </tr>
    <tr>
        <td height="25" colspan="4" id="Template" style="padding:10px; line-height:22px; font-size:14px;"><iframe src="Html.asp" id="Gets" width="600px" height="200px"></iframe></td>
        </tr>
</table>
<%
	Else
%>
<p style="text-align:center;">系统设置为动态模式或伪静态模式，无需HTML生成！</p>
<%
	End If
%>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==============================
'函 数 名：HtmlCreate()
'作    用：生成常规HTML
'参    数：
'==============================
Sub HtmlCreate()
	Dim cRs
	Dim CreatType,CreatModuleId,ModuleType,ModuleId,ModuleIsIndex,ModuleUrl,ContentId,ContentUrl,ContentFileName,ContentTitle
	Dim CreatNow,t_i
	Set cRs=Server.Createobject("Adodb.RecordSet")
	CreatType=Clng(Request.QueryString("Id"))
	PageNow=FKFun.GetNumeric("Page",1)
	CreatNow=FKFun.GetNumeric("CreatNow",0)
	CreatModuleId=FKFun.GetNumeric("CreatModuleId",0)
	Call HtmlCss()
	Select Case CreatType
		Case 1 '生成首页
			Call FKHtml.CreatIndex()
		Case 2
			ModuleType=3
		Case 3
			ModuleType=0
		Case 4
			ModuleType=4
		Case 5
			ModuleType=6
		Case 6  '生成专题页
			Sqlstr="Select Fk_Subject_Id,Fk_Subject_Name,Fk_Subject_Dir,Fk_Subject_Template From [Fk_Subject] Order By Fk_Subject_Id Asc"
			cRs.Open Sqlstr,Conn,1,1
			While Not cRs.Eof
				Id=cRs("Fk_Subject_Id")
				Call FKHtml.CreatSubject(Id,cRs("Fk_Subject_Template"),cRs("Fk_Subject_Dir"),cRs("Fk_Subject_Name"))
				cRs.MoveNext
			Wend
			cRs.Close
		Case 7
			ModuleType=1
		Case 8
			ModuleType=2
		Case 9
			ModuleType=7
		Case 10
			ModuleType=5
	End Select
	If Instr(",2,3,5,10,",","&CreatType&",")>0 Then
		PageNow=1
		CreatNow=1
	End If
	If Instr(",4,",","&CreatType&",")>0 Then
		CreatNow=1
	End If
	If Instr(",7,8,9,",","&CreatType&",")>0 And CreatNow=0 Then
		Select Case CreatType
			Case 7
				ModuleType=1
				Sqlstr="Select Fk_Article_Id,Fk_Article_Title,Fk_Article_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Show=1 And (Fk_Article_Url='' Or Fk_Article_Url Is Null) Order By Fk_Article_Id Desc"
			Case 8
				ModuleType=2
				Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Show=1 And (Fk_Product_Url='' Or Fk_Product_Url Is Null) Order By Fk_Product_Id Desc"
			Case 9
				ModuleType=7
				Sqlstr="Select Fk_Down_Id,Fk_Down_Title,Fk_Down_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Show=1 And (Fk_Down_Url='' Or Fk_Down_Url Is Null) Order By Fk_Down_Id Desc"
		End Select
		cRs.Open Sqlstr,Conn,1,1
		If Not cRs.Eof Then
			cRs.PageSize=PageSizes
			If PageNow>cRs.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=cRs.PageCount
			cRs.AbsolutePage=PageNow
			PageAll=cRs.RecordCount
			t_i=1
			While (Not cRs.Eof) And t_i<PageSizes+1
				Select Case ModuleType
					Case 1
						ContentId=cRs("Fk_Article_Id")
						ContentFileName=cRs("Fk_Article_FileName")
						ContentTitle=cRs("Fk_Article_Title")
					Case 2
						ContentId=cRs("Fk_Product_Id")
						ContentFileName=cRs("Fk_Product_FileName")
						ContentTitle=cRs("Fk_Product_Title")
					Case 7
						ContentId=cRs("Fk_Down_Id")
						ContentFileName=cRs("Fk_Down_FileName")
						ContentTitle=cRs("Fk_Down_Title")
				End Select
				ContentUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),cRs("Fk_Module_Type"),cRs("Fk_Module_Id"))
				If ContentFileName<>"" Then
					ContentUrl=ContentUrl&ContentFileName&FKTemplate.GetHtmlSuffix()
				Else
					ContentUrl=ContentUrl&ContentId&FKTemplate.GetHtmlSuffix()
				End If
				Call FKHtml.CreatPage(ContentId,cRs("Fk_Module_Id"),cRs("Fk_Module_Type"),ContentUrl,ContentTitle)
				cRs.MoveNext
				t_i=t_i+1
			Wend
		End If
		cRs.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=Html.asp?Type=2&Id=<%=CreatType%>&Page=<%=PageNow+1%>">
<%
			Set cRs=Nothing
			Call FKFun.ShowErr("本页生成完成，正在转向下一页，共"&PageCounts&"页，当前第"&PageNow&"页！",2)
		Else
			Select Case ModuleType
				Case 7
					Response.Write("文章部分生成完毕！<br />")
				Case 8
					Response.Write("产品部分生成完毕！<br />")
				Case 9
					Response.Write("下载部分生成完毕！<br />")
			End Select
			Response.Flush()
			Response.Clear()
			CreatNow=1
			PageNow=1
		End If
	End If
	If Instr(",2,3,5,10,",","&CreatType&",")>0 And CreatNow=1 Then
		Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Type="&ModuleType&" And Fk_Module_Show=1"
		If ModuleType=5 Then
			Sqlstr=Sqlstr&" And Fk_Module_UrlType=1"
		End If
		Sqlstr=Sqlstr&" Order By Fk_Module_Id Desc"
		cRs.Open Sqlstr,Conn,1,1
		While Not cRs.Eof
			ModuleId=cRs("Fk_Module_Id")
			ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
			Call FKHtml.CreatModule(ModuleId,ModuleType,ModuleUrl,cRs("Fk_Module_Name"),"",0,1)
			cRs.MoveNext
		Wend
		cRs.Close
	End If
	If Instr(",4,7,8,9,",","&CreatType&",")>0 And CreatNow=1 Then
		Sqlstr="Select Top 1 Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Type="&ModuleType&" And Fk_Module_Show=1 And Fk_Module_Id>"&CreatModuleId&" Order By Fk_Module_Id Asc"
		cRs.Open Sqlstr,Conn,1,1
		If Not cRs.Eof Then
			ModuleId=cRs("Fk_Module_Id")
			ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
			Call FKHtml.CreatModule(ModuleId,ModuleType,ModuleUrl,cRs("Fk_Module_Name"),"Html.asp?Type=2&Id="&CreatType&"&CreatNow=1&CreatModuleId=",CreatModuleId,1)
		End If
		cRs.Close
	End If
	Response.Write("生成完毕！")
End Sub

'==============================
'函 数 名：HtmlSelectCreate()
'作    用：生成选择项HTML
'参    数：
'==============================
Sub HtmlSelectCreate()
	Dim cRs
	Dim MenuId,ModuleType,ModuleId,ModuleIsIndex,ModuleUrl,ContentId,ContentUrl,ContentFileName,ContentTitle,CreatType,CreatDay
	Dim CreatNow,CreatModuleId,CreatDate,t_i
	Set cRs=Server.Createobject("Adodb.RecordSet")
	MenuId=FKFun.GetNumeric("MenuId",0)
	ModuleId=FKFun.GetNumeric("ModuleId",0)
	CreatType=FKFun.GetNumeric("CreatType",0)
	CreatDay=FKFun.GetNumeric("CreatDay",0)
	PageNow=FKFun.GetNumeric("Page",1)
	CreatNow=FKFun.GetNumeric("CreatNow",0)
	CreatModuleId=FKFun.GetNumeric("CreatModuleId",0)
	Call HtmlCss()
	If CreatDay=1 Then
		CreatDate=Date()
	Else
		CreatDate=DateAdd("d",-(CreatDay-1),Date())
	End If
	If MenuId=0 Then
		Call FKFun.ShowErr("请先选择主菜单！",0)
	End If
	If ModuleId=0 Then
		Call FKFun.ShowErr("请选择需要生成的模块！",0)
	End If
	Sqlstr="Select Fk_Module_Name,Fk_Module_Type,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Id="&ModuleId&" And Fk_Module_Show=1"
	cRs.Open Sqlstr,Conn,1,1
	If Not cRs.Eof Then
		ModuleType=cRs("Fk_Module_Type")
		ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
	Else
		cRs.Close
		Call FKFun.ShowErr("模块不存在！",0)
	End If
	cRs.Close
	If Instr(",1,2,7,",","&ModuleType&",")>0 And CreatType<>1 And CreatNow=0 Then
		Select Case ModuleType
			Case 1
				Sqlstr="Select Fk_Article_Id,Fk_Article_Title,Fk_Article_FileName,Fk_Module_MUrl From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Show=1 And (Fk_Article_Url='' Or Fk_Article_Url Is Null) And Fk_Article_Module="&ModuleId&""
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And Fk_Article_Time>=#"&CreatDate&"#"
				End If
				Sqlstr=Sqlstr&" Order By Fk_Article_Id Desc"
			Case 2
				Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_FileName,Fk_Module_MUrl From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Show=1 And (Fk_Product_Url='' Or Fk_Product_Url Is Null) And Fk_Product_Module="&ModuleId&""
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And Fk_Product_Time>=#"&CreatDate&"#"
				End If
				Sqlstr=Sqlstr&" Order By Fk_Product_Id Desc"
			Case 7
				Sqlstr="Select Fk_Down_Id,Fk_Down_Title,Fk_Down_FileName,Fk_Module_MUrl From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Show=1 And (Fk_Down_Url='' Or Fk_Down_Url Is Null) And Fk_Down_Module="&ModuleId&""
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And Fk_Down_Time>=#"&CreatDate&"#"
				End If
				Sqlstr=Sqlstr&" Order By Fk_Down_Id Desc"
		End Select
		cRs.Open Sqlstr,Conn,1,1
		If Not cRs.Eof Then
			cRs.PageSize=PageSizes
			If PageNow>cRs.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=cRs.PageCount
			cRs.AbsolutePage=PageNow
			PageAll=cRs.RecordCount
			t_i=1
			While (Not cRs.Eof) And t_i<PageSizes+1
				Select Case ModuleType
					Case 1
						ContentId=cRs("Fk_Article_Id")
						ContentFileName=cRs("Fk_Article_FileName")
						ContentTitle=cRs("Fk_Article_Title")
					Case 2
						ContentId=cRs("Fk_Product_Id")
						ContentFileName=cRs("Fk_Product_FileName")
						ContentTitle=cRs("Fk_Product_Title")
					Case 7
						ContentId=cRs("Fk_Down_Id")
						ContentFileName=cRs("Fk_Down_FileName")
						ContentTitle=cRs("Fk_Down_Title")
				End Select
				ContentUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
				If ContentFileName<>"" Then
					ContentUrl=ContentUrl&ContentFileName&FKTemplate.GetHtmlSuffix()
				Else
					ContentUrl=ContentUrl&ContentId&FKTemplate.GetHtmlSuffix()
				End If
				Call FKHtml.CreatPage(ContentId,ModuleId,ModuleType,ContentUrl,ContentTitle)
				cRs.MoveNext
				t_i=t_i+1
			Wend
		End If
		cRs.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=Html.asp?Type=3&MenuId=<%=MenuId%>&ModuleId=<%=ModuleId%>&CreatType=<%=CreatType%>&CreatDay=<%=CreatDay%>&Page=<%=PageNow+1%>">
<%
			Set cRs=Nothing
			Call FKFun.ShowErr("本页生成完成，正在转向下一页，共"&PageCounts&"页，当前第"&PageNow&"页！",2)
		Else
			Select Case ModuleType
				Case 1
					Response.Write("文章部分生成完毕！<br />")
				Case 2
					Response.Write("产品部分生成完毕！<br />")
				Case 7
					Response.Write("下载部分生成完毕！<br />")
			End Select
			Response.Flush()
			Response.Clear()
			CreatNow=1
		End If
	Else
		CreatNow=1
	End If
	If CreatNow=1 And CreatType<>2 Then
		If Instr(",4,1,2,7,",","&ModuleType&",")>0 Then
			Sqlstr="Select Top 1 Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Type="&ModuleType&" And Fk_Module_Show=1 And Fk_Module_Id>"&CreatModuleId&" And Fk_Module_Id="&ModuleId&" Order By Fk_Module_Id Asc"
			cRs.Open Sqlstr,Conn,1,1
			If Not cRs.Eof Then
				ModuleId=cRs("Fk_Module_Id")
				ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
				Call FKHtml.CreatModule(ModuleId,ModuleType,ModuleUrl,cRs("Fk_Module_Name"),"Html.asp?Type=3&MenuId="&MenuId&"&ModuleId="&ModuleId&"&CreatType="&CreatType&"&CreatDay="&CreatDay&"&CreatNow=1&CreatModuleId=",CreatModuleId,1)
			End If
			cRs.Close
		Else
			Sqlstr="Select Fk_Module_Name,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Id="&ModuleId&" And Fk_Module_Show=1"
			If ModuleType=5 Then
				Sqlstr=Sqlstr&" And Fk_Module_UrlType=1"
			End If
			Sqlstr=Sqlstr&" Order By Fk_Module_Id Desc"
			cRs.Open Sqlstr,Conn,1,1
			If Not cRs.Eof Then
				ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),ModuleType,ModuleId)
				Call FKHtml.CreatModule(ModuleId,ModuleType,ModuleUrl,cRs("Fk_Module_Name"),"",0,1)
			End If
			cRs.Close
		End If
	End If
	Response.Write("生成完毕！")
End Sub

'==============================
'函 数 名：HtmlDayCreate()
'作    用：一键生成HTML
'参    数：
'==============================
Sub HtmlDayCreate()
	Dim cRs
	Dim CreatDay,CreatType,CreatDate,ModuleList
	Dim CreatNow,CreatModuleId
	Dim ContentId,ContentFileName,ContentTitle,ContentUrl
	Dim ModuleId,ModuleUrl
	Set cRs=Server.Createobject("Adodb.RecordSet")
	CreatDay=FKFun.GetNumeric("CreatDay",1)
	CreatType=FKFun.GetNumeric("CreatType",1)
	PageNow=FKFun.GetNumeric("Page",1)
	ModuleList=FKFun.HTMLEncode(Request.QueryString("ModuleList"))
	CreatNow=FKFun.GetNumeric("CreatNow",0)
	CreatModuleId=FKFun.GetNumeric("CreatModuleId",0)
	If CreatDay=1 Then
		CreatDate=Date()
	Else
		CreatDate=DateAdd("d",-(CreatDay-1),Date())
	End If
	Call HtmlCss()
	If CreatNow=0 Then
		Select Case CreatType
			Case 1
				Sqlstr="Select Fk_Article_Id,Fk_Article_Title,Fk_Article_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Show=1 And (Fk_Article_Url='' Or Fk_Article_Url Is Null) And Fk_Article_Time>=#"&CreatDate&"# Order By Fk_Article_Id Desc"
			Case 2
				Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Show=1 And (Fk_Product_Url='' Or Fk_Product_Url Is Null) And Fk_Product_Time>=#"&CreatDate&"# Order By Fk_Product_Id Desc"
			Case 7
				Sqlstr="Select Fk_Down_Id,Fk_Down_Title,Fk_Down_FileName,Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl,Fk_Module_Menu From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Show=1 And (Fk_Down_Url='' Or Fk_Down_Url Is Null) And Fk_Down_Time>=#"&CreatDate&"# Order By Fk_Down_Id Desc"
		End Select
		cRs.Open Sqlstr,Conn,1,1
		If Not cRs.Eof Then
			cRs.PageSize=PageSizes
			If PageNow>cRs.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=cRs.PageCount
			cRs.AbsolutePage=PageNow
			PageAll=cRs.RecordCount
			t_i=1
			While (Not cRs.Eof) And t_i<PageSizes+1
				If ModuleList="" Then
					ModuleList=cRs("Fk_Module_Id")
				Else
					If Instr(",,"&ModuleList&",,",","&cRs("Fk_Module_Id")&",")>0 Then
						ModuleList=ModuleList&","&cRs("Fk_Module_Id")
					End If
				End If
				Select Case cRs("Fk_Module_Type")
					Case 1
						ContentId=cRs("Fk_Article_Id")
						ContentFileName=cRs("Fk_Article_FileName")
						ContentTitle=cRs("Fk_Article_Title")
					Case 2
						ContentId=cRs("Fk_Product_Id")
						ContentFileName=cRs("Fk_Product_FileName")
						ContentTitle=cRs("Fk_Product_Title")
					Case 7
						ContentId=cRs("Fk_Down_Id")
						ContentFileName=cRs("Fk_Down_FileName")
						ContentTitle=cRs("Fk_Down_Title")
				End Select
				ContentUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),cRs("Fk_Module_Type"),cRs("Fk_Module_Id"))
				If ContentFileName<>"" Then
					ContentUrl=ContentUrl&ContentFileName&FKTemplate.GetHtmlSuffix()
				Else
					ContentUrl=ContentUrl&ContentId&FKTemplate.GetHtmlSuffix()
				End If
				Call FKHtml.CreatPage(ContentId,cRs("Fk_Module_Id"),cRs("Fk_Module_Type"),ContentUrl,ContentTitle)
				cRs.MoveNext
				t_i=t_i+1
			Wend
		End If
		cRs.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=Html.asp?Type=4&CreatDay=<%=CreatDay%>&CreatType=<%=CreatType%>&ModuleList=<%=Server.URLEncode(ModuleList)%>&Page=<%=PageNow+1%>">
<%
			Response.Write("本页生成完成，正在转向下一页，共"&PageCounts&"页，当前第"&PageNow&"页！")
			Set cRs=Nothing
			Call FKDB.DB_Close()
			Response.End()
		Else
			Select Case CreatType
				Case 1
					Response.Write("文章部分生成完毕！<br />")
				Case 2
					Response.Write("产品部分生成完毕！<br />")
				Case 7
					Response.Write("下载部分生成完毕！<br />")
			End Select
			Response.Flush()
			Response.Clear()
			CreatNow=1
		End If
	End If
	If ModuleList<>"" And CreatNow=1 Then
		Sqlstr="Select Top 1 Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Id In ("&ModuleList&") And Fk_Module_Show=1 And Fk_Module_Id>"&CreatModuleId&" Order By Fk_Module_Id Asc"
		cRs.Open Sqlstr,Conn,1,1
		If Not cRs.Eof Then
			ModuleId=cRs("Fk_Module_Id")
			ModuleUrl=FKTemplate.GetModuleUrl(cRs("Fk_Module_MUrl"),CreatType,ModuleId)
			Call FKHtml.CreatModule(ModuleId,CreatType,ModuleUrl,cRs("Fk_Module_Name"),"Html.asp?Type=4&CreatDay="&CreatDay&"&CreatType="&CreatType&"&ModuleList="&Server.URLEncode(ModuleList)&"&CreatNow=1&CreatModuleId=",CreatModuleId,1)
		End If
		cRs.Close
	End If
	Select Case CreatType
		Case 1
%>
<meta http-equiv="refresh" content="1;URL=Html.asp?Type=4&CreatDay=<%=CreatDay%>&CreatType=2">
<%
			Call FKFun.ShowErr("正在生成，请稍候！",2)
		Case 2
%>
<meta http-equiv="refresh" content="1;URL=Html.asp?Type=4&CreatDay=<%=CreatDay%>&CreatType=7">
<%
			Call FKFun.ShowErr("正在生成，请稍候！",2)
		Case 7
			Call FKHtml.CreatIndex()
	End Select
	Response.Write("生成完毕！")
End Sub

Sub HtmlCss()
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
End Sub
%>
<!--#Include File="../Code.asp"-->