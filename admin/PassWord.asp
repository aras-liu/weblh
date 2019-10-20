<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/PassWord.asp
'文件用途：修改密码拉取页面
'版权所有：方卡在线
'==========================================

'定义页面变量
Dim Fk_Admin_LoginPass1,Fk_Admin_LoginPass2

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call PassWordForm() '读取修改密码表单
	Case 2
		Call PassWordDo() '修改密码操作
End Select

'==========================================
'函 数 名：PassWordForm()
'作    用：读取修改密码表单
'参    数：
'==========================================
Sub PassWordForm()
%>
<form id="ChangePass" name="ChangePass" method="post" action="PassWord.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:300px;">修改密码[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:300px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">用户名：</td>
	        <td>&nbsp;<input type="text" name="AdminName" id="AdminName" class="Input Input150" value="<%=Request.Cookies("FkAdminName")%>" readonly="readonly" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">密码：</td>
	        <td>&nbsp;<input type="password" name="Fk_Admin_LoginPass1" id="Fk_Admin_LoginPass1" class="Input Input150" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">重复密码：</td>
	        <td>&nbsp;<input type="password" name="Fk_Admin_LoginPass2" id="Fk_Admin_LoginPass2" class="Input Input150" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:280px;">
        <input type="submit" onclick="Sends('ChangePass','PassWord.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：PassWordDo()
'作    用：修改密码操作
'参    数：
'==========================================
Sub PassWordDo()
	Fk_Admin_LoginPass1=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass1")))
	Fk_Admin_LoginPass2=FKFun.HTMLEncode(Trim(Request.Form("Fk_Admin_LoginPass2")))
	Call FKFun.ShowString(Fk_Admin_LoginPass1,1,50,0,"请输入密码！","密码不能大于50个字符！")
	If Fk_Admin_LoginPass1<>Fk_Admin_LoginPass2 Then
		Call FKFun.ShowErr("两次密码不一致！",2)
	End If
	Sqlstr="Select Fk_Admin_Id,Fk_Admin_LoginPass From [Fk_Admin] Where Fk_Admin_Id=" & Request.Cookies("FkAdminId")
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Admin_LoginPass")=Md5(Md5(Fk_Admin_LoginPass1,32),16)
		Rs.Update()
		Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass1,32),16)
		Application.UnLock()
		Response.Write("密码修改成功！")
	Else
		Response.Write("用户不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->