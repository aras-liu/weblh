<!--#Include File="../Include.asp"--><%
'==========================================
'文 件 名：Subject/Index.asp
'文件用途：专题页
'版权所有：
'==========================================

'定义变量
Dim PageCode
Dim pTemp1,pTemp2,pTemp3,pTemp4,pTemp5
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode

'获取变量
Id=Clng(Request.QueryString("Id"))

'获取专题
Sqlstr="Select Fk_Subject_Template From [Fk_Subject] Where Fk_Subject_Id=" & Id
Rs.Open Sqlstr,Conn,1,1
If Not Rs.Eof Then
	pTemp1=Rs("Fk_Subject_Template")
Else
	Rs.Close
	Call FKFun.ShowErr("专题未找到！",0)
End If
Rs.Close

PageCode=FKPageCode.cSubject(Id,pTemp1)

Response.Write(PageCode)

%>
<!--#Include File="../Code.asp"-->
