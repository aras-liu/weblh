<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/DelWord.asp
'文件用途：过滤字符拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System4",Request.Cookies("FkAdminLimit1"))

Dim DelWord

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call DelWordBox() '读取过滤字符
	Case 2
		Call DelWordDo() '设置过滤字符
End Select

'==========================================
'函 数 名：DelWordBox()
'作    用：读取过滤字符
'参    数：
'==========================================
Sub DelWordBox()
	DelWord=FKFso.FsoFileRead("DelWord.dat")
%>
<form id="DelWordSet" name="DelWordSet" method="post" action="DelWord.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">过滤字符设置[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right" class="MainTableTop">过滤字符&nbsp;&nbsp;<span class="qbox" title="<p>过滤字符设置用于设置敏感词，设置的敏感词可以对文章、产品等等内容进行过滤，多个过滤字符串用空格隔开。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td>&nbsp;<textarea name="DelWord" cols="50" rows="20" class="TextArea" id="DelWord"><%=DelWord%></textarea><br /><span style="color:#F00">（多个关键字用空格隔开）</span></td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:580px;">
        <input type="submit" onclick="$('#DelWord').text(escape($('#DelWord').val()));Sends('DelWordSet','DelWord.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="设 置" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：DelWordDo()
'作    用：设置过滤字符
'参    数：
'==========================================
Sub DelWordDo()
	DelWord=Request.Form("DelWord")
	Call FKFso.CreateFile("DelWord.dat",DelWord)
	Response.Write("过滤字符修改成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->