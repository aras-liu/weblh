<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Admin.asp
'文件用途：管理员管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"0","")

'定义页面变量
Dim Fk_Admin_LoginName,Fk_Admin_LoginPass1,Fk_Admin_LoginPass2,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call AdminList() '管理员列表
	Case 2
		Call AdminAddForm() '添加管理员表单
	Case 3
		Call AdminAddDo() '执行添加管理员
	Case 4
		Call AdminEditForm() '修改管理员表单
	Case 5
		Call AdminEditDo() '执行修改管理员
	Case 6
		Call AdminDelDo() '执行删除管理员
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：AdminList()
'作    用：管理员列表
'参    数：
'==========================================
Sub AdminList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Admin.asp?Type=2');">添加新管理员</a></li>
    </ul>
</div>
<div id="ListTop">
    管理员管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">序号</td>
            <td align="center" class="ListTdTop">用户名</td>
            <td align="center" class="ListTdTop">姓名</td>
            <td align="center" class="ListTdTop">权限</td>
            <td align="center" class="ListTdTop">状态</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Admin_Id,Fk_Admin_LoginName,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit From [Fk_Admin] Order By Fk_Admin_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim LimitName
		Dim Rs2
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		i=1
		While Not Rs.Eof
			If Rs("Fk_Admin_Limit")>0 Then
				Sqlstr="Select Fk_Limit_Name From [Fk_Limit] Where Fk_Limit_Id=" & Rs("Fk_Admin_Limit")
				Rs2.Open Sqlstr,Conn,1,1
				If Not Rs2.Eof Then
					LimitName=Rs2("Fk_Limit_Name")
				Else
					LimitName="未知权限"
				End If
				Rs2.Close
			Else
				LimitName="超级管理员"
			End If
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_Admin_LoginName")%></td>
            <td align="center"><%=Rs("Fk_Admin_Name")%></td>
            <td align="center"><%=LimitName%></td>
            <td align="center"><%If Rs("Fk_Admin_User")=0 Then%>禁用<%Else%>正常<%End If%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Admin.asp?Type=4&Id=<%=Rs("Fk_Admin_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Admin_LoginName")%>”帐号？此操作不可逆！','Admin.asp?Type=6&Id=<%=Rs("Fk_Admin_Id")%>','MainRight','Admin.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		Set Rs2=Nothing
	Else
%>
        <tr>
            <td height="25" colspan="6" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="6">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：AdminAddForm()
'作    用：添加管理员表单
'参    数：
'==========================================
Sub AdminAddForm()
%>
<form id="AdminAdd" name="AdminAdd" method="post" action="Admin.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:500px;">添加新管理员[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right">登录名：</td>
            <td>&nbsp;<input name="Fk_Admin_LoginName" type="text" class="Input" id="Fk_Admin_LoginName" />&nbsp;&nbsp;<span class="qbox" title="<p>管理用户登录用，建议用英文、数字，登录名不可重复，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">密码：</td>
            <td>&nbsp;<input name="Fk_Admin_LoginPass1" type="password" class="Input" id="Fk_Admin_LoginPass1" />&nbsp;&nbsp;<span class="qbox" title="<p>管理用户密码，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">重复密码：</td>
            <td>&nbsp;<input name="Fk_Admin_LoginPass2" type="password" class="Input" id="Fk_Admin_LoginPass2" />&nbsp;&nbsp;<span class="qbox" title="<p>重新输入用户密码，两次密码需要一致，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">姓名：</td>
            <td>&nbsp;<input name="Fk_Admin_Name" type="text" class="Input" id="Fk_Admin_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>管理员姓名，便于管理，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">权限：</td>
            <td>&nbsp;<select name="Fk_Admin_Limit" class="Input" id="Fk_Admin_Limit">
                    <option value="0">超级管理员</option>
<%
	Sqlstr="Select Fk_Limit_Id,Fk_Limit_Name From [Fk_Limit] Order By Fk_Limit_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Limit_Id")%>"><%=Rs("Fk_Limit_Name")%></option>
<%
		Rs.MoveNext
	Wend
Rs.Close
%>
                    </select>&nbsp;&nbsp;<span class="qbox" title="<p>管理员权限，普通管理员权限在“权限管理”中进行设置（此选项对超级管理员admin无效）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">状态：</td>
            <td>&nbsp;<input name="Fk_Admin_User" class="Input" type="radio" id="Fk_Admin_User" value="1" checked="checked" />正常
                <input type="radio" class="Input" name="Fk_Admin_User" id="Fk_Admin_User" value="0" />禁用&nbsp;&nbsp;<span class="qbox" title="<p>设置管理员是否可用，如设为禁用，则不可登录（此选项对超级管理员admin无效）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:480px;">
        <input type="submit" onclick="Sends('AdminAdd','Admin.asp?Type=3',0,'',0,1,'MainRight','Admin.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：AdminAddDo
'作    用：执行添加管理员
'参    数：
'==============================
Sub AdminAddDo()
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginName")))
	Fk_Admin_LoginPass1=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass1")))
	Fk_Admin_LoginPass2=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass2")))
	Fk_Admin_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_Name")))
	Fk_Admin_User=Trim(Request.Form("Fk_Admin_User"))
	Fk_Admin_Limit=Trim(Request.Form("Fk_Admin_Limit"))
	Call FKFun.ShowString(Fk_Admin_LoginName,1,50,0,"请输入登录名！","登录名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_LoginPass1,1,50,0,"请输入登录密码！","登录密码不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_Name,1,50,0,"请输入姓名！","姓名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Admin_Limit,"请选择权限！")
	If Fk_Admin_LoginPass1<>Fk_Admin_LoginPass2 Then
		Call FKFun.ShowErr("两次密码不一致！",2)
	End If
	Sqlstr="Select Fk_Admin_Id,Fk_Admin_LoginName,Fk_Admin_LoginPass,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit From [Fk_Admin] Where Fk_Admin_LoginName='"&Fk_Admin_LoginName&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Admin_LoginName")=Fk_Admin_LoginName
		Rs("Fk_Admin_LoginPass")=Md5(Md5(Fk_Admin_LoginPass1,32),16)
		Rs("Fk_Admin_Name")=Fk_Admin_Name
		Rs("Fk_Admin_User")=Fk_Admin_User
		Rs("Fk_Admin_Limit")=Fk_Admin_Limit
		Rs.Update()
		Application.UnLock()
		Response.Write("新管理员添加成功！")
	Else
		Response.Write("该登录名已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：AdminEditForm()
'作    用：修改管理员表单
'参    数：
'==========================================
Sub AdminEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Admin_LoginName,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit From [Fk_Admin] Where Fk_Admin_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Admin_LoginName=Rs("Fk_Admin_LoginName")
		Fk_Admin_Name=Rs("Fk_Admin_Name")
		Fk_Admin_User=Rs("Fk_Admin_User")
		Fk_Admin_Limit=Rs("Fk_Admin_Limit")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此管理员，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="AdminEdit" name="AdminEdit" method="post" action="Admin.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:500px;">修改管理员[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="30" align="right">登录名：</td>
        <td>&nbsp;<input name="Fk_Admin_LoginName" disabled="disabled" type="text" class="Input" id="Fk_Admin_LoginName" value="<%=Fk_Admin_LoginName%>" readonly="readonly" />&nbsp;&nbsp;<span class="qbox" title="<p>修改管理员时登录名不可修改。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="30" align="right">密码：</td>
        <td>&nbsp;<input name="Fk_Admin_LoginPass1" type="password" class="Input" id="Fk_Admin_LoginPass1" />*不改密码请留空&nbsp;&nbsp;<span class="qbox" title="<p>不修改密码的话请留空，如修改密码请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="30" align="right">重复密码：</td>
        <td>&nbsp;<input name="Fk_Admin_LoginPass2" type="password" class="Input" id="Fk_Admin_LoginPass2" />*不改密码请留空&nbsp;&nbsp;<span class="qbox" title="<p>不修改密码的话请留空，如修改密码请重复上面的输入的字符串。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="30" align="right">真实姓名：</td>
        <td>&nbsp;<input name="Fk_Admin_Name" type="text" class="Input" id="Fk_Admin_Name" value="<%=Fk_Admin_Name%>" />&nbsp;&nbsp;<span class="qbox" title="<p>管理员姓名，便于管理，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="30" align="right">权限：</td>
        <td>&nbsp;<select name="Fk_Admin_Limit" class="Input" id="Fk_Admin_Limit">
                <option value="0"<%=FKFun.BeSelect(Fk_Admin_Limit,0)%>>超级管理员</option>
<%
	Sqlstr="Select Fk_Limit_Id,Fk_Limit_Name From [Fk_Limit] Order By Fk_Limit_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Limit_Id")%>"<%=FKFun.BeSelect(Fk_Admin_Limit,Rs("Fk_Limit_Id"))%>><%=Rs("Fk_Limit_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                    </select>&nbsp;&nbsp;<span class="qbox" title="<p>管理员权限，普通管理员权限在“权限管理”中进行设置（此选项对超级管理员admin无效）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">状态：</td>
            <td>&nbsp;<input name="Fk_Admin_User" class="Input" type="radio" id="Fk_Admin_User" value="1"<%=FKFun.BeCheck(Fk_Admin_User,1)%> />正常
            <input type="radio" name="Fk_Admin_User" class="Input" id="Fk_Admin_User" value="0"<%=FKFun.BeCheck(Fk_Admin_User,0)%> />禁用&nbsp;&nbsp;<span class="qbox" title="<p>设置管理员是否可用，如设为禁用，则不可登录（此选项对超级管理员admin无效）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:480px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('AdminEdit','Admin.asp?Type=5',0,'',0,1,'MainRight','Admin.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：AdminEditDo
'作    用：执行修改管理员
'参    数：
'==============================
Sub AdminEditDo()
	Fk_Admin_LoginPass1=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass1")))
	Fk_Admin_LoginPass2=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass2")))
	Fk_Admin_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_Name")))
	Fk_Admin_User=Trim(Request.Form("Fk_Admin_User"))
	Fk_Admin_Limit=Trim(Request.Form("Fk_Admin_Limit"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Admin_LoginPass1,0,50,0,"请输入登录密码！","登录密码不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_Name,1,50,0,"请输入姓名！","姓名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Admin_User,"请选择账户状态！")
	Call FKFun.ShowNum(Fk_Admin_Limit,"请选择权限！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	If Fk_Admin_LoginPass1<>Fk_Admin_LoginPass2 And Fk_Admin_LoginPass1<>"" Then
		Call FKFun.ShowErr("两次密码不一致！",2)
	End If
	If Id=1 Then
		Fk_Admin_User=1
		Fk_Admin_Limit=0
	End If
	Sqlstr="Select Fk_Admin_Id,Fk_Admin_LoginPass,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit From [Fk_Admin] Where Fk_Admin_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		If Fk_Admin_LoginPass1<>"" Then
			Rs("Fk_Admin_LoginPass")=Md5(Md5(Fk_Admin_LoginPass1,32),16)
		End If
		Rs("Fk_Admin_Name")=Fk_Admin_Name
		Rs("Fk_Admin_User")=Fk_Admin_User
		Rs("Fk_Admin_Limit")=Fk_Admin_Limit
		Rs.Update()
		If Request.Cookies("FkAdminId")=Id And Fk_Admin_LoginPass1<>"" Then
			Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass1,32),16)
		End If
		Application.UnLock()
		Response.Write("管理员修改成功！")
	Else
		Response.Write("管理员不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：AdminDelDo
'作    用：执行删除管理员
'参    数：
'==============================
Sub AdminDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	If Id=1 Then
		Call FKFun.ShowErr("原始帐号无法删除！",2)
	End If
	Sqlstr="Select Fk_Admin_Id From [Fk_Admin] Where Fk_Admin_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("管理员删除成功！")
	Else
		Response.Write("管理员不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->