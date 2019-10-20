<!--#Include File="../../Include.asp"--><%
'==========================================
'文 件 名：Plugin/VoteR.asp
'文件用途：投票结果文件
'版权所有：
'插件版本：V1.0.0
'==========================================

Dim Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket,Fk_Vote_Count,TempArr2
Id=Clng(Request.QueryString("Id"))

Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id=" & Id
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	Fk_Vote_Name=Rs("Fk_Vote_Name")
	Fk_Vote_Content=Rs("Fk_Vote_Content")
	Fk_Vote_Ticket=Rs("Fk_Vote_Ticket")
	Fk_Vote_Count=Rs("Fk_Vote_Count")
	TempArr=Split(Fk_Vote_Content,"<br />")
	TempArr2=Split(Fk_Vote_Ticket,"|")
Else
	Call FKFun.AlertInfo("未找到投票项目！",SiteDir)
End If
Rs.Close
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Fk_Site_Name%>--投票结果</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style></head>

<body>
<table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" style="border-collapse:collapse;">
    <tr>
        <td height="25" colspan="2" align="center"><%=Fk_Vote_Name%>投票结果[共有<%=Fk_Vote_Count%>人次投票]</td>
    </tr>
<%
For i=0 To UBound(TempArr)
%>
    <tr>
        <td width="140" height="25" align="right"><%=TempArr(i)%>：</td>
        <td width="354">&nbsp;<%=TempArr2(i)%>票</td>
    </tr>
<%
Next
%>
</table>
</body>
</html>
<!--#Include File="../../Code.asp"-->
