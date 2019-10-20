<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>上传图片</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {
	font-size: 12px;
	color: #000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000;
}
a:hover {
	text-decoration: none;
	color: #333;
}
a:active {
	text-decoration: none;
	color: #000;
}
.Input,.Button {
	height:20px;
	line-height:20px;
	border:1px solid #999;
	border-bottom:1px solid #CCC;
	border-right:1px solid #CCC;
}
.Button {
	margin-left:0;
}
-->
</style></head>
<body>
<Form name=Form action="Upload.asp?Immediate=<%=Request.QueryString("Type")%>&Form=<%=Request.QueryString("Form")%>&Input=<%=Request.QueryString("Input")%>" method="post" enctype="multipart/Form-data">
    <input type="file" name="FileData" id="FileData" class="Input" Size="20">
    &nbsp;<input type="submit" name="submit" value="上传" class="Button">
</Form>
</body>
</html>