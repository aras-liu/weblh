<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Job.asp
'文件用途：招聘管理拉取页面
'版权所有：方卡在线
'==========================================

'定义页面变量
Dim Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date,Fk_Job_Module,Fk_Job_Menu,Fk_Job_Field
Dim Fk_Module_Id,Fk_Module_Name,Fk_Module_Menu

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call JobList() '招聘列表
	Case 2
		Call JobAddForm() '添加招聘表单
	Case 3
		Call JobAddDo() '执行添加招聘
	Case 4
		Call JobEditForm() '修改招聘表单
	Case 5
		Call JobEditDo() '执行修改招聘
	Case 6
		Call JobDelDo() '执行删除招聘
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：JobList()
'作    用：招聘列表
'参    数：
'==========================================
Sub JobList()
	Session("NowPage")=FkFun.GetNowUrl()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select Fk_Module_Name,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Rs.Close
		Call FKFun.ShowErr("模块不存在！",2)
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Job.asp?Type=2&ModuleId=<%=Fk_Module_Id%>');">添加新招聘</a></li>
    </ul>
</div>
<div id="ListTop">
    “<%=Fk_Module_Name%>”管理
    &nbsp;&nbsp;快速通道：<select name="D1" id="D1" onChange="eval(this.options[this.selectedIndex].value);" class="Input">
      <option value="alert('请选择模块');">请选择模块</option>
<%
Call FKAdmin.GetModuleList(0,Fk_Module_Menu,0,Fk_Module_Id,"")
%>
</select>
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">职位</td>
            <td align="center" class="ListTdTop">人数</td>
            <td align="center" class="ListTdTop">工作地点</td>
            <td align="center" class="ListTdTop">有效期</td>
            <td align="center" class="ListTdTop">添加时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Job_Id,Fk_Job_Name,Fk_Job_Count,Fk_Job_Area,Fk_Job_Date,Fk_Job_Time From [Fk_Job] Where Fk_Job_Module="&Fk_Module_Id&" Order By Fk_Job_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_Job_Name")%></td>
            <td align="center"><%=Rs("Fk_Job_Count")%></td>
            <td align="center"><%=Rs("Fk_Job_Area")%></td>
            <td align="center"><%If Rs("Fk_Job_Date")=0 Then%>长期<%Else%><%=Rs("Fk_Job_Date")%><%End If%></td>
            <td align="center"><%=Rs("Fk_Job_Time")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Job.asp?Type=4&Id=<%=Rs("Fk_Job_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Job_Name")%>”？此操作不可逆！','Job.asp?Type=6&Id=<%=Rs("Fk_Job_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="7" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="7">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：JobAddForm()
'作    用：添加招聘表单
'参    数：
'==========================================
Sub JobAddForm()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select Fk_Module_Name,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="JobAdd" name="JobAdd" method="post" action="Job.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加“<%=Fk_Module_Name%>”模块新招聘[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">职位：</td>
	        <td>&nbsp;<input name="Fk_Job_Name" type="text" class="Input" id="Fk_Job_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>职位名称，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">招聘人数：</td>
            <td>&nbsp;<input name="Fk_Job_Count" type="text" class="Input" id="Fk_Job_Count" />&nbsp;&nbsp;<span class="qbox" title="<p>招聘人数，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">工作地点：</td>
            <td>&nbsp;<input name="Fk_Job_Area" type="text" class="Input" id="Fk_Job_Area" />&nbsp;&nbsp;<span class="qbox" title="<p>工作地点，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">有效期：</td>
            <td>&nbsp;<input name="Fk_Job_Date" type="text" class="Input" id="Fk_Job_Date" />&nbsp;天（请输入数字，如果长期有效请填0）&nbsp;&nbsp;<span class="qbox" title="<p>有效期，不能为空，必须输入数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
<%
	Call FKAdmin.ShowField(0,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
	Call FKAdmin.ShowField(2,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
	Call FKAdmin.ShowField(3,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
%>
        <tr>
            <td height="30" align="right">招聘要求<span class="qbox" title="<p>招聘要求，不能为空，请输入1-5000个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td><textarea name="Fk_Job_About" cols="60" rows="15" class="<%=EditorClass%>" id="Fk_Job_About"></textarea></td>
        </tr>
<%
	Call FKAdmin.ShowField(1,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,EditorClass)
%>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('JobAdd','Job.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：JobAddDo
'作    用：执行添加招聘
'参    数：
'==============================
Sub JobAddDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("您无此权限！",2)
	End If
	Fk_Job_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Name")))
	Fk_Job_Count=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Count")))
	Fk_Job_About=Request.Form("Fk_Job_About")
	Fk_Job_Area=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Area")))
	Fk_Job_Date=Trim(Request.Form("Fk_Job_Date"))
	Call FKFun.ShowString(Fk_Job_Name,1,50,0,"请输入招聘职位！","招聘职位不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_Count,1,50,0,"请输入招聘数量！","招聘数量不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_About,1,5000,0,"请输入招聘要求！","招聘要求不能大于5000个字符！")
	Call FKFun.ShowString(Fk_Job_Area,1,50,0,"请输入工作地点！","工作地点不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Job_Date,"请输入有效期，单位为天，必须是数字，长期有效请填0！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Fk_Job_Field=FKAdmin.GetFieldData(0,"(Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')")
	Sqlstr="Select Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Rs.Close
		Call FKFun.ShowErr("模块不存在！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Job_Id,Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date,Fk_Job_Field,Fk_Job_Module,Fk_Job_Menu From [Fk_Job] Where Fk_Job_Name='"&Fk_Job_Name&"' And Fk_Job_About='"&Fk_Job_Area&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Job_Name")=Fk_Job_Name
		Rs("Fk_Job_Count")=Fk_Job_Count
		Rs("Fk_Job_About")=Fk_Job_About
		Rs("Fk_Job_Area")=Fk_Job_Area
		Rs("Fk_Job_Date")=Fk_Job_Date
		Rs("Fk_Job_Field")=Fk_Job_Field
		Rs("Fk_Job_Module")=Fk_Module_Id
		Rs("Fk_Job_Menu")=Fk_Module_Menu
		Rs.Update()
		Application.UnLock()
		Response.Write("新招聘添加成功！")
	Else
		Response.Write("该招聘职位已经被发布，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：JobEditForm()
'作    用：修改招聘表单
'参    数：
'==========================================
Sub JobEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date,Fk_Job_Field,Fk_Job_Module From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Job_Name=FKFun.HTMLDncode(Rs("Fk_Job_Name"))
		Fk_Job_Count=FKFun.HTMLDncode(Rs("Fk_Job_Count"))
		Fk_Job_About=FKFun.HTMLDncode(Rs("Fk_Job_About"))
		Fk_Job_Area=FKFun.HTMLDncode(Rs("Fk_Job_Area"))
		Fk_Job_Date=Rs("Fk_Job_Date")
		Fk_Job_Module=Rs("Fk_Job_Module")
		If IsNull(Rs("Fk_Job_Field")) Or Rs("Fk_Job_Field")="" Then
			Fk_Job_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Job_Field=Split(Rs("Fk_Job_Field"),"[-Fangka_Field-]")
		End If
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到招聘，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Job_Module,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
%>
<form id="JobEdit" name="JobEdit" method="post" action="Job.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改招聘[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">职位：</td>
	        <td>&nbsp;<input name="Fk_Job_Name" value="<%=Fk_Job_Name%>" type="text" class="Input" id="Fk_Job_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>职位名称，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">招聘人数：</td>
            <td>&nbsp;<input name="Fk_Job_Count" value="<%=Fk_Job_Count%>" type="text" class="Input" id="Fk_Job_Count" />&nbsp;&nbsp;<span class="qbox" title="<p>招聘人数，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">工作地点：</td>
            <td>&nbsp;<input name="Fk_Job_Area" value="<%=Fk_Job_Area%>" type="text" class="Input" id="Fk_Job_Area" />&nbsp;&nbsp;<span class="qbox" title="<p>工作地点，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">有效期：</td>
            <td>&nbsp;<input name="Fk_Job_Date" value="<%=Fk_Job_Date%>" type="text" class="Input" id="Fk_Job_Date" />&nbsp;天（请输入数字，如果长期有效请填0）&nbsp;&nbsp;<span class="qbox" title="<p>有效期，不能为空，必须输入数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
<%
	Call FKAdmin.ShowField(0,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Job_Module&",%%')",Fk_Job_Field,"")
	Call FKAdmin.ShowField(2,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Job_Module&",%%')",Fk_Job_Field,"")
	Call FKAdmin.ShowField(3,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Job_Module&",%%')",Fk_Job_Field,"")
%>
        <tr>
            <td height="30" align="right">招聘要求<span class="qbox" title="<p>招聘要求，不能为空，请输入1-5000个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td><textarea name="Fk_Job_About" cols="60" rows="15" class="<%=EditorClass%>" id="Fk_Job_About"><%=Fk_Job_About%></textarea></td>
        </tr>
<%
	Call FKAdmin.ShowField(1,0," And (Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Job_Module&",%%')",Fk_Job_Field,EditorClass)
%>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
		<input type="hidden" name="ModuleId" value="<%=Fk_Job_Module%>" />
        <input type="submit" onclick="Sends('JobEdit','Job.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：JobEditDo
'作    用：执行修改招聘
'参    数：
'==============================
Sub JobEditDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	Fk_Job_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Name")))
	Fk_Job_Count=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Count")))
	Fk_Job_About=Request.Form("Fk_Job_About")
	Fk_Job_Area=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Area")))
	Fk_Job_Date=Trim(Request.Form("Fk_Job_Date"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Job_Name,1,255,0,"请输入招聘职位！","招聘职位不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_Count,1,50,0,"请输入招聘数量！","招聘数量不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_About,1,5000,0,"请输入招聘要求！","招聘要求不能大于5000个字符！")
	Call FKFun.ShowString(Fk_Job_Area,1,50,0,"请输入工作地点！","工作地点不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Job_Date,"请输入有效期，单位为天，必须是数字，长期有效请填0！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Call FKFun.ShowNum(Fk_Module_Id,"系统参数错误，请刷新页面！")
	Fk_Job_Field=FKAdmin.GetFieldData(0,"(Fk_Field_Content Like '%%,Job,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')")
	Sqlstr="Select Fk_Job_Id,Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date,Fk_Job_Field,Fk_Job_Module From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Job_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		Application.Lock()
		Rs("Fk_Job_Name")=Fk_Job_Name
		Rs("Fk_Job_Count")=Fk_Job_Count
		Rs("Fk_Job_About")=Fk_Job_About
		Rs("Fk_Job_Area")=Fk_Job_Area
		Rs("Fk_Job_Date")=Fk_Job_Date
		Rs("Fk_Job_Field")=Fk_Job_Field
		Rs.Update()
		Application.UnLock()
		Response.Write("招聘修改成功！")
	Else
		Response.Write("招聘不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：JobDelDo
'作    用：执行删除招聘
'参    数：
'==============================
Sub JobDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Job_Id,Fk_Job_Module From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Job_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("招聘删除成功！")
	Else
		Response.Write("招聘不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->