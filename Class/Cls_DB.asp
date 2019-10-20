<%
'==========================================
'文 件 名：Class/Cls_DB.asp
'文件用途：数据库函数类
'版权所有：
'==========================================

Class Cls_DB
	Private ConnStr
	'==============================
	'函 数 名：DB_Open
	'作    用：创建读取对象
	'参    数：
	'==============================
	Public Sub DB_Open()
		On Error Resume Next
		Set Conn = Server.CreateObject("Adodb.Connection")
		Set Rs=Server.Createobject("Adodb.RecordSet")
		ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteData)
		Conn.Open ConnStr
		If Err Then
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
			Response.Write "<body style='font-size:12px'>"
			Response.Write "错 误 号：" & Err.Number & "<br />"
			Response.Write "错误描述：" & Err.Description & "<br />"
			Response.Write "错误来源：" & Err.Source & "<br />"
			Response.Write "如您出现此问题请查看用户手册安装部分！<br />"
			Response.Write "</body>"
			Response.End()
		End If
	End Sub
	
	'==============================
	'函 数 名：DB_Close
	'作    用：关闭读取对象
	'参    数：
	'==============================	
	Public Sub DB_Close()
		Set Rs=Nothing
		If IsObject(Conn) Then Conn.Close
		Set Conn = Nothing
	End Sub
End Class
%>
