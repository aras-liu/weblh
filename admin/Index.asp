<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Admin/Index.asp
'文件用途：后台管理首页
'版权所有：方卡在线
'==========================================
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>后台管理</title>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../Js/jquery.min.js"></script>
<script type="text/javascript" src="../Js/jquery.form.min.js"></script>
<%
If Fk_Site_Edit=0 Then
%>
<script type="text/javascript" src="Editor/xhEditor/xheditor-zh-cn.min.js"></script>
<%
ElseIf Fk_Site_Edit=1 Then
%>
<script type="text/javascript" charset="utf-8" src="Editor/kindeditor/kindeditor-min.js"></script>
<%
End If
%>
<script type="text/javascript" src="../Js/jquery.tooltip.js"></script>
<script type="text/javascript" src="../Js/jquery.tristate.js"></script>
<script type="text/javascript" src="../Js/jquery.colorselect.js"></script>
<script type="text/javascript" src="../Js/function.js"></script>
<script type="text/javascript" src="../Js/date.js"></script>
<script type="text/javascript">
$(document).ready(function(){
<%
If Login=False Then
	Response.Write("	ShowBox(""Login.asp?Type=1"");")
	If FKAdmin.GetAdminDir()="admin" Then
		Response.Write("	alert(""系统检测到您的管理目录是默认的admin，这样不利于系统安全！\n\n建议：目录名设为6位以上、尽量复杂一些！"");")
	End If
Else
	Response.Write("	SetRContent(""UserInfo"",""Get.asp?Type=4"");"&vbcrlf)
	Response.Write("	SetRContent(""Nav"",""Get.asp?Type=1"");"&vbcrlf)
	Response.Write("	SetRContent(""MainLeft"",""Get.asp?Type=2"");"&vbcrlf)
	Response.Write("	SetRContent(""MainRight"",""Get.asp?Type=3"");"&vbcrlf)
	Response.Write("	SetExt('"&Fk_Site_LinkExt&"','"&Fk_Site_PicExt&"','"&Fk_Site_FlashExt&"','"&Fk_Site_MediaExt&"');")
End If
%>
	PageReSize();
});
</script>
</head>

<body>
<div id="AllBox">
    <div id="Bodys" style="width:100%">
        <div id="PageTop">
            <div id="Top">
              <div id="UserInfo"><a href="javascript:void(0);" onClick="ShowBox('Login.asp?Type=1');" title="请您先登录！">请您先登录！</a></div>
                <div class="Cal"></div>
            </div>
            <div id="Nav">
            </div>
        </div>
        <div id="PageMain">
            <div id="MainLeft">
            </div>
            <div id="MainRight">
            </div>
            <div class="Cal"></div>
        </div>
        <div id="Boxs" style="display:none">
            <div id="BoxsContent">
                <div id="BoxContent">
                </div>
            </div>
            <div id="AlphaBox" onClick="CloseBox();"></div>
        </div>
        <div id="MsgBox"><div id="MsgContent"></div></div>
    </div>
</div>
</body>
</html>
<!--#Include File="../Code.asp"-->