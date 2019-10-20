<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Map.asp
'文件用途：搜索引擎地图生成拉取页面
'版权所有：方卡在线
'==========================================
dim fenzhan
Call FKAdmin.AdminCheck(3,"System6",Request.Cookies("FkAdminLimit1"))

Set FKTemplate=New Cls_Template

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call MapBox() '读取搜索引擎地图生成
	Case 2
		Call MapGoogleDo() '生成google地图
    Case 3
        Call MapBaiduDo() '生成Baidu地图
End Select

'==========================================
'函 数 名：MapBox()
'作    用：读取搜索引擎地图生成
'参    数：
'==========================================
Sub MapBox()
%>
<div id="BoxTop" style="width:700px;">搜索引擎地图生成[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="25%" height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Map.asp?Type=2';">生成Google地图</a>&nbsp;&nbsp;<span class="qbox" title="<p>生成的Google地图请到谷歌网站管理员工具中提交。</p>"><img src="Images/help.jpg" /></span></td>
        <td width="25%" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='Map.asp?Type=3';">生成Baidu地图</a>&nbsp;&nbsp;<span class="qbox" title="<p>生成的Baidu地图请到百度站长平台中提交。</p>"><img src="Images/help.jpg" /></span></td>
        <td width="25%" align="center">&nbsp;</td>
        <td width="25%" align="center">&nbsp;</td>
        </tr>
    <tr>
        <td height="25" colspan="4" align="center">&nbsp;&nbsp;网站地图生成结果</td>
    </tr>
    <tr>
        <td height="25" colspan="4" id="Template" style="padding:10px; line-height:22px; font-size:14px;"><iframe src="" id="Gets" width="600px" height="200px"></iframe></td>
        </tr>
</table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==============================
'函 数 名：MapGoogleDo()
'作    用：生成google地图
'参    数：
'==============================
Sub MapGoogleDo()
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #ff0000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
	Dim ContentUrl,HtmlType
	If Fk_Site_Html=0 Then
		HtmlType="?"
	Else
		HtmlType=""
	End If
	Temp="<?xml version=""1.0"" encoding=""UTF-8""?>"&vbLf
	Temp=Temp&"<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"&vbLf
	Temp=Temp&"<url>"&vbLf
	Temp=Temp&"<loc>"&Fk_Site_Url&"</loc>"&vbLf
	Temp=Temp&"</url>"&vbLf
	Sqlstr="Select Fk_Article_Id,Fk_Article_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Show=1 And (Fk_Article_Url='' Or Fk_Article_Url Is Null) Order By Fk_Article_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ContentUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Article_Id"),Rs("Fk_Article_FileName"))
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ContentUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Product_Id,Fk_Product_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Show=1 And (Fk_Product_Url='' Or Fk_Product_Url Is Null) Order By Fk_Product_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ContentUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Product_Id"),Rs("Fk_Product_FileName"))
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ContentUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Down_Id,Fk_Down_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Show=1 And (Fk_Down_Url='' Or Fk_Down_Url Is Null) Order By Fk_Down_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ContentUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Down_Id"),Rs("Fk_Down_Id"))
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ContentUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Module_MUrl,Fk_Module_Type,Fk_Module_Id From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Type<>5 Order By Fk_Module_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ContentUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ContentUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Temp=Temp&"</urlset>"&vbLf
	Call FKFso.CreateFile("../google.xml",Temp)
	Response.Write("<p><a href=""../google.xml"" target=""_blank"">Google地图生成成功</a></p>")
End Sub

Sub MapBaiduDo()
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #ff0000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
    Dim ArticleUrl, ProductUrl, DownUrl, ModuleUrl
    Dim tempArticle, tempPorduct, tempDown, tempModule

    tempArticle="<?xml  version=""1.0"" encoding=""utf-8""?>"&vbLf
    tempArticle=tempArticle&"<urlset>"&vbLf
	Sqlstr="Select Top 50000 Fk_Article_Id,Fk_Article_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Show=1 And (Fk_Article_Url='' Or Fk_Article_Url Is Null) Order By Fk_Article_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ArticleUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Article_Id"),Rs("Fk_Article_FileName"))
		tempArticle=tempArticle&"<url>"&vbLf
		tempArticle=tempArticle&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ArticleUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		tempArticle=tempArticle&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
    tempArticle=tempArticle&"</urlset>"&vbLf
    Call FKFso.CreateFile("../baidu_article.xml",tempArticle)

    tempPorduct="<?xml  version=""1.0"" encoding=""utf-8""?>"&vbLf
    tempPorduct=tempPorduct&"<urlset>"&vbLf
	Sqlstr="Select Top 50000 Fk_Product_Id,Fk_Product_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Show=1 And (Fk_Product_Url='' Or Fk_Product_Url Is Null) Order By Fk_Product_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ProductUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Product_Id"),Rs("Fk_Product_FileName"))
		tempPorduct=tempPorduct&"<url>"&vbLf
		tempPorduct=tempPorduct&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ProductUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		tempPorduct=tempPorduct&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
    tempPorduct=tempPorduct&"</urlset>"&vbLf
    Call FKFso.CreateFile("../baidu_product.xml",tempPorduct)

    tempDown="<?xml  version=""1.0"" encoding=""utf-8""?>"&vbLf
    tempDown=tempDown&"<urlset>"&vbLf
	Sqlstr="Select Top 50000 Fk_Down_Id,Fk_Down_FileName,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Show=1 And (Fk_Down_Url='' Or Fk_Down_Url Is Null) Order By Fk_Down_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		DownUrl=FKTemplate.GetContentUrl(FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Down_Id"),Rs("Fk_Down_Id"))
		tempDown=tempDown&"<url>"&vbLf
		tempDown=tempDown&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&DownUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		tempDown=tempDown&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
    tempDown=tempDown&"</urlset>"&vbLf
    Call FKFso.CreateFile("../baidu_download.xml",tempDown)

    tempModule="<?xml  version=""1.0"" encoding=""utf-8""?>"&vbLf
    tempModule=tempModule&"<urlset>"&vbLf
    tempModule=tempModule&"<url>"&vbLf
	tempModule=tempModule&"<loc>"&Fk_Site_Url&"</loc>"&vbLf
	tempModule=tempModule&"</url>"&vbLf
	Sqlstr="Select Fk_Module_MUrl,Fk_Module_Type,Fk_Module_Id From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Type<>5 Order By Fk_Module_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ModuleUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
		tempModule=tempModule&"<url>"&vbLf
		tempModule=tempModule&"<loc>"&Server.HTMLEncode(Replace(Replace(Fk_Site_Url&ModuleUrl,"//","/"),"http:/","http://"))&"</loc>"&vbLf
		tempModule=tempModule&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
    tempModule=tempModule&"</urlset>"&vbLf
    Call FKFso.CreateFile("../baidu_module.xml",tempModule)

    Temp="<?xml  version=""1.0"" encoding=""utf-8""?>"&vbLf
    Temp=Temp&"<sitemapindex>"&vbLf
    Temp=Temp&"<sitemap>"&vbLf
    Temp=Temp&"<loc>"&Fk_Site_Url&"baidu_article.xml</loc>"&vbLf
    Temp=Temp&"<lastmod>"&FormatDateTime(Date())&"</lastmod>"&vbLf
    Temp=Temp&"</sitemap>"&vbLf
    Temp=Temp&"<sitemap>"&vbLf
    Temp=Temp&"<loc>"&Fk_Site_Url&"baidu_product.xml</loc>"&vbLf
    Temp=Temp&"<lastmod>"&FormatDateTime(Date())&"</lastmod>"&vbLf
    Temp=Temp&"</sitemap>"&vbLf
    Temp=Temp&"<sitemap>"&vbLf
    Temp=Temp&"<loc>"&Fk_Site_Url&"baidu_download.xml</loc>"&vbLf
    Temp=Temp&"<lastmod>"&FormatDateTime(Date())&"</lastmod>"&vbLf
    Temp=Temp&"</sitemap>"&vbLf
    Temp=Temp&"<sitemap>"&vbLf
    Temp=Temp&"<loc>"&Fk_Site_Url&"baidu_module.xml</loc>"&vbLf
    Temp=Temp&"<lastmod>"&FormatDateTime(Date())&"</lastmod>"&vbLf
    Temp=Temp&"</sitemap>"&vbLf
    Temp=Temp&"</sitemapindex>"&vbLf
    Call FKFso.CreateFile("../baidu.xml",Temp)

    Response.Write("<p><a href=""../baidu.xml"" target=""_blank"">Baidu地图生成成功</a></p>")
End Sub
%>
<!--#Include File="../Code.asp"-->