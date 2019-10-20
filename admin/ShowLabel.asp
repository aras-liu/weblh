<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/ShowLab.asp
'文件用途：模版标签拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System2",Request.Cookies("FkAdminLimit1"))

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call SiteLabel() '读取常规标签
	Case 2
		Call ForLabel() '读取For标签
	Case 3
		Call OtherLabel() '读取其他标签
End Select

'==========================================
'函 数 名：OtherLabel()
'作    用：读取其他标签
'参    数：
'==========================================
Sub OtherLabel()
	Dim s_For,s_Option
	s_For=Trim(Request.Form("For"))
	s_Option=Split(Request.Form("Label"),", ")
	Select Case s_For
		Case "Flash"
%>
<p><input type="text" class="Input" size="100" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Plugin/Flash/Index.asp?Type=<%=s_Option(2)%>&Menu=<%=s_Option(0)%>&Module=<%=s_Option(1)%>&Width=<%=s_Option(3)%>&Height=<%=s_Option(4)%>&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;Flash轮换代码</p>
<%
		Case "Info"
%>
<p><input type="text" class="Input" size="50" value="{$Info(<%=s_Option(0)%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;独立信息标签</p>
<%
		Case "Im"
%>
<p><input type="text" class="Input" size="80" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Plugin/Im/Index.asp&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;客服悬浮框代码</p>
<%
	End Select
End Sub

'==========================================
'函 数 名：SiteLabel()
'作    用：读取常规标签
'参    数：
'==========================================
Sub SiteLabel()
	Dim ArticleFieldList,ProductFieldList,DownFieldList,InfoFieldList,SiteFieldList
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Field_Tag,Fk_Field_Name,Fk_Field_Content From [Fk_Field] Where Fk_Field_Model=0 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Instr(Rs("Fk_Field_Content"),",Article,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeArticle,")>0 Then
			ArticleFieldList=ArticleFieldList&"<p><input type=""text"" class=""Input"" size=""70"" value=""{$Field_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Product,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeProduct,")>0 Then
			ProductFieldList=ProductFieldList&"<p><input type=""text"" class=""Input"" size=""70"" value=""{$Field_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Down,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeDown,")>0 Then
			DownFieldList=DownFieldList&"<p><input type=""text"" class=""Input"" size=""70"" value=""{$Field_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Info,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeInfo,")>0 Then
			InfoFieldList=InfoFieldList&"<p><input type=""text"" class=""Input"" size=""70"" value=""{$Field_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Field_Tag,Fk_Field_Name,Fk_Field_Content From [Fk_Field] Where Fk_Field_Model=1 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		SiteFieldList=SiteFieldList&"<p><input type=""text"" class=""Input"" size=""70"" value=""{$Site_Field_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		Rs.MoveNext
	Wend
	Rs.Close
	Select Case Id
		Case 1 '全站常规标签
%>
<p><input type="text" class="Input" size="70" value="{$SiteName$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点名称</p>
<p><input type="text" class="Input" size="70" value="{$SiteUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点链接</p>
<p><input type="text" class="Input" size="70" value="{$SiteKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点关键字</p>
<p><input type="text" class="Input" size="70" value="{$SiteDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点描述</p>
<p><input type="text" class="Input" size="70" value="{$SiteArticleCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章总数</p>
<p><input type="text" class="Input" size="70" value="{$SiteProductCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品总数</p>
<p><input type="text" class="Input" size="70" value="{$SiteDownCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载总数</p>
<p><input type="text" class="Input" size="70" value="{$SiteSkin$}" />&nbsp;&nbsp;&nbsp;&nbsp;模板路径</p>
<p><input type="text" class="Input" size="70" value="{$SiteDir$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点路径</p>
<p><input type="text" class="Input" size="70" value="{$PageCrumbs$}" />&nbsp;&nbsp;&nbsp;&nbsp;面包屑菜单</p>
<%
			Response.Write(SiteFieldList)
		Case 2 '静态页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<%
		Case 3 '信息页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p><input type="text" class="Input" size="70" value="{$ModuleContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块内容</p>
<%
			Response.Write(InfoFieldList)
		Case 4 '文章列表页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p><input type="text" class="Input" size="70" value="{$ModulePageCode(页码类型)$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块页码代码</p>
<p><input type="text" class="Input" size="70" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="70" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="70" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="70" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 5 '产品列表页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p><input type="text" class="Input" size="70" value="{$ModulePageCode(页码类型)$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块页码代码</p>
<p><input type="text" class="Input" size="70" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="70" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="70" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="70" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 12 '下载列表页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p><input type="text" class="Input" size="70" value="{$ModulePageCode(页码类型)$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块页码</p>
<p><input type="text" class="Input" size="70" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="70" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="70" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="70" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 6 '文章页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;栏目父模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ArticleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章编号</p>
<p><input type="text" class="Input" size="70" value="{$ArticleTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章标题</p>
<p><input type="text" class="Input" size="70" value="{$ArticleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章关键字</p>
<p><input type="text" class="Input" size="70" value="{$ArticleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章描述</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图小图</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图大图</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图列表（与题图输出FOR共用）</p>
<p><input type="text" class="Input" size="70" value="{$ArticleFrom$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章来源</p>
<p><input type="text" class="Input" size="70" value="{$ArticleContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章内容</p>
<p><input type="text" class="Input" size="70" value="{$ArticleClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章点击量</p>
<p><input type="text" class="Input" size="70" value="{$ArticleAdmin$}" />&nbsp;&nbsp;&nbsp;&nbsp;添加文章的管理员</p>
<p><input type="text" class="Input" size="70" value="{$ArticleTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章添加时间（格式见其他标签所述）</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$ArticlePrevPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇图片</p>
<p><input type="text" class="Input" size="70" value="{$ArticleNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$ArticleNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$ArticleNextPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇图片</p>
<p><input type="text" class="Input" size="70" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=1&Id={$ArticleId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;文章HTML点击JS，放置页面底部</p>
<%
			Response.Write(ArticleFieldList)
		Case 7 '产品页常规标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;栏目父模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ProductId$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品编号</p>
<p><input type="text" class="Input" size="70" value="{$ProductTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品名称</p>
<p><input type="text" class="Input" size="70" value="{$ProductKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品关键字</p>
<p><input type="text" class="Input" size="70" value="{$ProductDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品描述</p>
<p><input type="text" class="Input" size="70" value="{$ProductPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品图片小图</p>
<p><input type="text" class="Input" size="70" value="{$ProductPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品图片大图</p>
<p><input type="text" class="Input" size="70" value="{$ProductPicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品题图列表（与题图输出FOR共用）</p>
<p><input type="text" class="Input" size="70" value="{$ProductContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品简介</p>
<p><input type="text" class="Input" size="70" value="{$ProductClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品点击量</p>
<p><input type="text" class="Input" size="70" value="{$ProductAdmin$}" />&nbsp;&nbsp;&nbsp;&nbsp;添加产品的管理员</p>
<p><input type="text" class="Input" size="70" value="{$ProductTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品添加时间（格式见其他标签所述）</p>
<p><input type="text" class="Input" size="70" value="{$ProductPrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$ProductPrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$ProductPrevPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇图片</p>
<p><input type="text" class="Input" size="70" value="{$ProductNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$ProductNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$ProductNextPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇图片</p>
<p><input type="text" class="Input" size="70" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=2&Id={$ProductId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;产品HTML点击JS，放置页面底部</p>
<%
			Response.Write(ProductFieldList)
		Case 13 '下载页标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;栏目父模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$DownId$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载编号</p>
<p><input type="text" class="Input" size="70" value="{$DownTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载名称</p>
<p><input type="text" class="Input" size="70" value="{$DownKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载关键字</p>
<p><input type="text" class="Input" size="70" value="{$DownDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载描述</p>
<p><input type="text" class="Input" size="70" value="{$DownPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载图片小图</p>
<p><input type="text" class="Input" size="70" value="{$DownPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载图片大图</p>
<p><input type="text" class="Input" size="70" value="{$DownPicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图列表（与题图输出FOR共用）</p>
<p><input type="text" class="Input" size="70" value="{$DownContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载简介</p>
<p><input type="text" class="Input" size="70" value="{$DownSystem$}" />&nbsp;&nbsp;&nbsp;&nbsp;适用系统</p>
<p><input type="text" class="Input" size="70" value="{$DownLanguage$}" />&nbsp;&nbsp;&nbsp;&nbsp;语言</p>
<p><input type="text" class="Input" size="70" value="{$DownFile$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载链接</p>
<p><input type="text" class="Input" size="70" value="{$DownClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载点击量</p>
<p><input type="text" class="Input" size="70" value="{$DownAdmin$}" />&nbsp;&nbsp;&nbsp;&nbsp;添加下载的管理员</p>
<p><input type="text" class="Input" size="70" value="{$DownCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载量</p>
<p><input type="text" class="Input" size="70" value="{$DownTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加时间（格式见其他标签所述）</p>
<p><input type="text" class="Input" size="70" value="{$DownPrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$DownPrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$DownPrevPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇图片</p>
<p><input type="text" class="Input" size="70" value="{$DownNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="70" value="{$DownNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="70" value="{$DownNextPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇图片</p>
<p><input type="text" class="Input" size="70" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=3&Id={$DownId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;下载HTML点击JS，放置页面底部</p>
<%
			Response.Write(DownFieldList)
		Case 8 '留言页专用标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p><input type="text" class="Input" size="70" value="{$ModulePageCode(页码类型)$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块页码</p>
<p><input type="text" class="Input" size="70" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="70" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="70" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="70" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 9 '专题页标签
%>
<p><input type="text" class="Input" size="70" value="{$SubjectId$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题编号</p>
<p><input type="text" class="Input" size="70" value="{$SubjectName$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题名称</p>
<%
		Case 10 'IF标签使用方法
%>
<p><input type="text" class="Input" size="70" value="{$If(参数1,参数2,比较方式)$}" />
&nbsp;&nbsp;&nbsp;&nbsp;IF标签开始，比较方式支持&lt;/&gt;/=/&gt;=/&lt;=/&lt;&gt;/Mod</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="如果成立输出的HTML" /></p>
<p><input type="text" class="Input" size="70" value="{$Else$}" /></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="如果不成立输出的HTML" /></p>
<p><input type="text" class="Input" size="70" value="{$End If$}" />&nbsp;&nbsp;&nbsp;&nbsp;IF标签结束</p>
<%
		Case 11 '搜索页标签
%>
<p><input type="text" class="Input" size="70" value="{$SearchStr$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索关键字</p>
<p><input type="text" class="Input" size="70" value="{$SearchType$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索类型</p>
<p><input type="text" class="Input" size="70" value="{$SearchPageCode(页码类型)$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索页码代码</p>
<p><input type="text" class="Input" size="70" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="70" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="70" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="70" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="70" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 14 '招聘页标签
%>
<p><input type="text" class="Input" size="70" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleFName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="70" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="70" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="70" value="{$ModuleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块关键字</p>
<p><input type="text" class="Input" size="70" value="{$ModuleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p><input type="text" class="Input" size="70" value="{$ModulePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块图片</p>
<p><input type="text" class="Input" size="70" value="{$ModuleSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<%
	End Select
End Sub

'==========================================
'函 数 名：ForLabel()
'作    用：读取For标签
'参    数：
'==========================================
Sub ForLabel()
	Dim s_For,s_Option
	Dim ArticleFieldList,ProductFieldList,DownFieldList,JobFieldList
	s_For=Trim(Request.Form("For"))
	s_Option=Trim(Replace(Request.Form("Label"),", ","/"))
	Sqlstr="Select Fk_Field_Tag,Fk_Field_Name,Fk_Field_Content From [Fk_Field] Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Instr(Rs("Fk_Field_Content"),",Article,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeArticle,")>0 Then
			ArticleFieldList=ArticleFieldList&"<p style=""margin-left:20px;""><input type=""text"" class=""Input"" size=""70"" value=""{$FieldList_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Product,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeProduct,")>0 Then
			ProductFieldList=ProductFieldList&"<p style=""margin-left:20px;""><input type=""text"" class=""Input"" size=""70"" value=""{$FieldList_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Down,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeDown,")>0 Then
			DownFieldList=DownFieldList&"<p style=""margin-left:20px;""><input type=""text"" class=""Input"" size=""70"" value=""{$FieldList_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		If Instr(Rs("Fk_Field_Content"),",Job,")>0 Or Instr(Rs("Fk_Field_Content"),",SeeJob,")>0 Then
			JobFieldList=JobFieldList&"<p style=""margin-left:20px;""><input type=""text"" class=""Input"" size=""70"" value=""{$FieldList_"&Rs("Fk_Field_Tag")&"$}"" />&nbsp;&nbsp;&nbsp;&nbsp;"&Rs("Fk_Field_Name")&"</p>"
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If s_For="GBookList" Then
		TempArr=Split(s_Option,"/")
		If Not IsNumeric(TempArr(1)) Then
			Call FKFun.ShowErr("请选择留言模块！",2)
		End If
		Sqlstr="Select Fk_Module_GModel From [Fk_Module] Where Fk_Module_Type=4 And Fk_Module_Id=" & TempArr(1)
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Temp=Rs("Fk_Module_GModel")
		Else
			Rs.Close
			Call FKFun.ShowErr("请选择留言模块！",2)
		End If
		Rs.Close
		Sqlstr="Select Fk_GModel_Content From [Fk_GModel] Where Fk_GModel_Id=" & Temp
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempArr=Split(Rs("Fk_GModel_Content"),"|-_-|Fangka|-_-|")
		Else
			Rs.Close
			Call FKFun.ShowErr("留言模型不存在！",2)
		End If
		Rs.Close
	End If
	Response.Write("<p><input type=""text"" class=""Input"" size=""70"" value=""{$For("&s_For&","&s_Option&")$}"" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>")
	Select Case s_For
		Case "Nav"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;输出序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块描述</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块内容数量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavSubhead$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块副标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavType$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块类型</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块题图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$NavSub$}" />&nbsp;&nbsp;&nbsp;&nbsp;二级菜单标签</p>
<%
		Case "ArticleList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListPageNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章描述</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListContent(缩略读取字数)$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章内容缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListPicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章题图列表（与题图输出FOR共用）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章添加时间（格式见其他标签所述）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章发布时间差（输出单位为天）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ArticleListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;文章点击量</p>
<%
			Response.Write(ArticleFieldList)
		Case "ProductList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListPageNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品描述</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListContent(缩略读取字数)$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品内容缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListPicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品题图列表（与题图输出FOR共用）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品添加时间（格式见其他标签所述）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品发布时间差（输出单位为天）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ProductListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;产品点击量</p>
<%
			Response.Write(ProductFieldList)
		Case "DownList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListPageNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载描述</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListContent(缩略读取字数)$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载简介缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListPicList$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图列表（与题图输出FOR共用）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListSystem$}" />&nbsp;&nbsp;&nbsp;&nbsp;适用系统</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListLanguage$}" />&nbsp;&nbsp;&nbsp;&nbsp;语言</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListFile$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载地址</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListTime(yyyy-mm-dd)$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加时间（格式见其他标签所述）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载发布时间差（输出单位为天）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$DownListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载点击量</p>
<%
			Response.Write(DownFieldList)
		Case "GBookList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListPageNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<%
			For Each Temp In TempArr
				If Instr(Temp,"|-^-|Fangka|-^-|") Then
				
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$GBookList<%=Split(Temp,"|-^-|Fangka|-^-|")(3)%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Split(Temp,"|-^-|Fangka|-^-|")(0)%></p>
<%
				End If
			Next
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$GBookListTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言时间</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$GBookListReContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;回复内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$GBookListReTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;回复时间</p>
<%
		Case "FriendsList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$FriendsName$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$FriendsAbout$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点简介</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$FriendsUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$FriendsLogo$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点LOGO</p>
<%
		Case "JobList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobName$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘数量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobAbout$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘简介</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobArea$}" />&nbsp;&nbsp;&nbsp;&nbsp;工作地点</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;有效期限</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$JobTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;发布时间</p>
<%
			Response.Write(JobFieldList)
		Case "SubjectList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$SubjectListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$SubjectListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题图片</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$SubjectListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题链接</p>
<%
		Case "PicList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$PicListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$PicListSmall$}" />&nbsp;&nbsp;&nbsp;&nbsp;小图地址</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$PicListBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;大图地址</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$PicListText$}" />&nbsp;&nbsp;&nbsp;&nbsp;图片说明</p>
<%
		Case "NumList"
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="70" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<%
	End Select
	Response.Write("<p><input type=""text"" class=""Input"" size=""70"" value=""{$Next$}"" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>")
End Sub
%>
<!--#Include File="../Code.asp"-->