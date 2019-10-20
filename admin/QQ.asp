<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/QQ.asp
'文件用途：客服浮窗拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System10",Request.Cookies("FkAdminLimit2"))

Dim Fk_QQ_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call QQBox() '读取客服浮窗
	Case 2
		Call QQDo() '设置客服浮窗
End Select

'==========================================
'函 数 名：QQBox()
'作    用：读取客服浮窗
'参    数：
'==========================================
Sub QQBox()
	Sqlstr="Select * From [Fk_QQ]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_QQ_Content=FKFun.HTMLDncode(Rs("Fk_QQ_Content"))
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此客服浮窗代码，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="QQSet" name="QQSet" method="post" action="QQ.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">客服浮窗设置[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td width="11%" height="30" align="right" class="MainTableTop">客服代码&nbsp;&nbsp;<span class="qbox" title="<p>客服代码请根据规则编写。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td width="89%"><textarea name="Fk_QQ_Content" style="width:100%;" rows="20" class="TextArea" id="Fk_QQ_Content"><%=Fk_QQ_Content%></textarea></td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:880px;">
        <input type="submit" onclick="Sends('QQSet','QQ.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="设 置" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：QQDo()
'作    用：设置客服浮窗
'参    数：
'==========================================
Sub QQDo()
	Fk_QQ_Content=FKFun.HTMLEncode(Replace(Request.Form("Fk_QQ_Content"),"%","&#37;"))
	Sqlstr="Update [Fk_QQ] Set Fk_QQ_Content='"&Fk_QQ_Content&"'"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("客服浮窗内容修改成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->