<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Click.asp
'文件用途：点击量显示
'版权所有：
'==========================================

'获取变量
Id=Clng(Request.QueryString("Id"))
Types=Trim(Request.QueryString("Type"))

If Types=1 Then
	Sqlstr="Select Fk_Article_Id,Fk_Article_Click From [Fk_Article] Where Fk_Article_Show=1 And Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
%>
if(document.getElementById("Click")){
    document.getElementById("Click").innerHTML=<%=Rs("Fk_Article_Click")+1%>;
}
<%
		Application.Lock()
		Rs("Fk_Article_Click")=Rs("Fk_Article_Click")+1
		Rs.Update()
		Application.UnLock()
	End If
	Rs.Close
ElseIf Types=2 Then
	Sqlstr="Select Fk_Product_Id,Fk_Product_Click From [Fk_Product] Where Fk_Product_Show=1 And Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
%>
if(document.getElementById("Click")){
    document.getElementById("Click").innerHTML=<%=Rs("Fk_Product_Click")+1%>;
}
<%
		Application.Lock()
		Rs("Fk_Product_Click")=Rs("Fk_Product_Click")+1
		Rs.Update()
		Application.UnLock()
	End If
	Rs.Close
ElseIf Types=3 Then
	Sqlstr="Select Fk_Down_Id,Fk_Down_Click,Fk_Down_Count From [Fk_Down] Where Fk_Down_Show=1 And Fk_Down_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
%>
if(document.getElementById("Click")){
    document.getElementById("Click").innerHTML=<%=Rs("Fk_Down_Click")+1%>;
}
if(document.getElementById("Count")){
    document.getElementById("Count").innerHTML=<%=Rs("Fk_Down_Count")%>;
}
<%
		Application.Lock()
		Rs("Fk_Down_Click")=Rs("Fk_Down_Click")+1
		Rs.Update()
		Application.UnLock()
	End If
	Rs.Close
End If
%>
<!--#Include File="Code.asp"-->
