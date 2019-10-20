<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Click.asp
'文件用途：点击量显示
'版权所有：
'==========================================

Dim Fk_Down_File

Id=Clng(Request.QueryString("Id"))

Sqlstr="Select Fk_Down_Id,Fk_Down_Count,Fk_Down_File From [Fk_Down] Where Fk_Down_Show=1 And Fk_Down_Id=" & Id
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	Fk_Down_File=Rs("Fk_Down_File")
	Application.Lock()
	Rs("Fk_Down_Count")=Rs("Fk_Down_Count")+1
	Rs.Update()
	Application.UnLock()
End If
Rs.Close
Response.Redirect(Fk_Down_File)
%>
<!--#Include File="Code.asp"-->
