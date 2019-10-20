<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
Response.Expires=-999
Session.Timeout=999
'==========================================
'文 件 名：Admin/Logout.asp
'文件用途：后台管理退出
'版权所有：方卡在线
'==========================================
Response.Cookies("FkAdminName")=""
Response.Cookies("FkAdminPass")=""
Response.Cookies("FkAdminIp")=""
Response.Cookies("FkAdminTime")=""
Response.Cookies("FkAdminName").Path="/"
Response.Cookies("FkAdminPass").Path="/"
Response.Cookies("FkAdminIp").Path="/"
Response.Cookies("FkAdminTime").Path="/"
Response.Redirect("Index.asp")
%>
<script type="text/javascript">
location.href='Index.asp';
</script>