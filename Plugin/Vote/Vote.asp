<!--#Include File="../../Include.asp"--><%
'==========================================
'文 件 名：Plugin/Vote.asp
'文件用途：在线投票
'版权所有：
'插件版本：V1.0.0
'==========================================

Dim Fk_Vote_Name,Fk_Vote_Content

Id=Clng(Request.QueryString("Id"))

Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id=" & Id
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	Fk_Vote_Name=Rs("Fk_Vote_Name")
	Fk_Vote_Content=Rs("Fk_Vote_Content")
	TempArr=Split(Fk_Vote_Content,"<br />")
%>
document.writeln("<form id=\"S\" name=\"S\" method=\"post\" action=\"Plugin/Vote/VoteDo.asp\">");
document.writeln("    <p><%=Fk_Vote_Name%></p>");
<%
	i=0
	For Each Temp In TempArr
		If Temp<>"" Then
%>
document.writeln("    <p><input type=\"checkbox\" name=\"V\" id=\"V\" value=\"<%=i%>\" /><%=Temp%></p>");
<%
			i=i+1
		End If
	Next
%>
document.writeln("    <p><input type=\"hidden\" name=\"Id\" value=\"<%=Id%>\" /><input type=\"submit\" name=\"button\" id=\"button\" value=\"投票\" />&nbsp;&nbsp;<input type=\"button\" onclick=\"window.open(\'<%=SiteDir%>Plugin/Vote/VoteR.asp?Id=<%=Id%>\',\'newwindow\',\'width=700,heigh=100,scrollbars=1\')\" name=\"button\" id=\"button\" value=\"查看结果\" /></p>");
document.writeln("</form>");

<%
End If
Rs.Close
%>
<!--#Include File="../../Code.asp"-->
