<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：KeyWord.asp
'文件用途：关键字拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System5",Request.Cookies("FkAdminLimit1"))

Dim KeyWord

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call KeyWordBox() '读取关键字
	Case 2
		Call KeyWordDo() '设置关键字
End Select

'==========================================
'函 数 名：KeyWordBox()
'作    用：读取关键字
'参    数：
'==========================================
Sub KeyWordBox()
	KeyWord=FKFso.FsoFileRead("KeyWord.dat")
%>
<form id="KeyWordSet" name="KeyWordSet" method="post" action="KeyWord.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">关键字设置[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right" class="MainTableTop">关键字&nbsp;&nbsp;<span class="qbox" title="<p>关键字设置后，可以在文章、产品等内容区间自动提取关键字，多个关键字用空格隔开。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td>&nbsp;<textarea name="KeyWord" cols="50" rows="20" class="TextArea" id="KeyWord"><%=KeyWord%></textarea><br /><span style="color:#F00">（多个关键字用空格隔开）</span></td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:580px;">
        <input type="submit" onclick="$('#KeyWord').text(escape($('#KeyWord').val()));Sends('KeyWordSet','KeyWord.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="设 置" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：KeyWordDo()
'作    用：设置关键字
'参    数：
'==========================================
Sub KeyWordDo()
	KeyWord=Request.Form("KeyWord")
	Call FKFso.CreateFile("KeyWord.dat",KeyWord)
	Response.Write("过滤字符修改成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->