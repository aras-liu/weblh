<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Admin/Login.asp
'文件用途：用户登录拉取页面
'版权所有：方卡在线
'==========================================

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call LoginBox() '读取登录信息
	Case 2
		Call LoginDo() '登录操作
End Select

'==========================================
'函 数 名：LoginBox()
'作    用：读取登录信息
'参    数：
'==========================================
Sub LoginBox()
%>
<form id="AdminLogin" name="AdminLogin" method="post" action="Login.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:300px;">用户登录[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:300px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">用户名：</td>
	        <td>&nbsp;<input type="text" name="AdminName" id="AdminName" class="Input Input150" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">密码：</td>
	        <td>&nbsp;<input type="password" name="AdminPass" id="AdminPass" class="Input Input150" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:280px;">
        <input type="submit" onclick="Sends('AdminLogin','Login.asp?Type=2',1,'Index.asp',0,0,'','');" class="Button" name="Enter" id="Enter" value="登 录" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：LoginDo()
'作    用：登录操作
'参    数：
'==========================================
Sub LoginDo()
	Dim Fk_Admin_LoginName,Fk_Admin_LoginPass
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request.Form("AdminName")))
	Fk_Admin_LoginPass=FKFun.HTMLEncode(Trim(Request.Form("AdminPass")))
	Call FKFun.ShowString(Fk_Admin_LoginName,1,50,0,"请输入登录名！","登录名名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_LoginPass,1,50,0,"请输入登录密码！","登录密码不能大于50个字符！")
	Sqlstr="Select Fk_Admin_Id From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&Fk_Admin_LoginName&"' And Fk_Admin_LoginPass='"&Md5(Md5(Fk_Admin_LoginPass,32),16)&"'"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Response.Cookies("FkAdminName")=Fk_Admin_LoginName
		Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass,32),16)
		Response.Cookies("FkAdminIp")=Request.ServerVariables("REMOTE_ADDR")
		Response.Cookies("FkAdminTime")=Now()
		Response.Cookies("CloseSysHidden")="0"
		If Fk_Site_Dir<>"" Then
			Response.Cookies("FkAdminName").Path="/"
			Response.Cookies("FkAdminPass").Path="/"
			Response.Cookies("FkAdminIp").Path="/"
			Response.Cookies("FkAdminTime").Path="/"
			Response.Cookies("CloseSysHidden").Path="/"
		End If
		Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		Application.Lock()
		Conn.Execute(Sqlstr)
		Application.UnLock()
		Response.Write("用户登录成功！")
	Else
		Response.Write("用户名或密码错误！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->