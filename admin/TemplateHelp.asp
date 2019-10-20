<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/TemplateHelp.asp
'文件用途：模版标签生成器拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System2",Request.Cookies("FkAdminLimit1"))

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call TemplateHelpBox() '读取标签生成器
	Case 3
		Call GetTemplate2() '读取标签2
End Select

'==========================================
'函 数 名：TemplateHelpBox()
'作    用：读取标签生成器
'参    数：
'==========================================
Sub TemplateHelpBox()
	Dim MenuSelect,RecommendSelect,SubjectSelect,FriendsTypeSelect,GetNumSelect,GetLenSelect,InfoSelect
	Sqlstr="Select Fk_Menu_Id,Fk_Menu_Name From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
        MenuSelect=MenuSelect&"<option value='"&Rs("Fk_Menu_Id")&"'>"&Rs("Fk_Menu_Name")&"</option>"
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
        RecommendSelect=RecommendSelect&"<option value='"&Rs("Fk_Recommend_Id")&"'>"&Rs("Fk_Recommend_Name")&"</option>"
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_Subject_Id,Fk_Subject_Name From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
        SubjectSelect=SubjectSelect&"<option value='"&Rs("Fk_Subject_Id")&"'>"&Rs("Fk_Subject_Name")&"</option>"
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select Fk_FriendsType_Id,Fk_FriendsType_Name From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
        FriendsTypeSelect=FriendsTypeSelect&"<option value='"&Rs("Fk_FriendsType_Id")&"'>"&Rs("Fk_FriendsType_Name")&"</option>"
		Rs.MoveNext
	Wend
	Rs.Close
	For i=1 To 20
		GetNumSelect=GetNumSelect&"<option value='"&i&"'>读取"&i&"条</option>"
	Next
	For i=1 To 20
		GetLenSelect=GetLenSelect&"<option value='"&i&"'>读取"&i&"字</option>"
	Next
	Sqlstr="Select Fk_Info_Id,Fk_Info_Name From [Fk_Info] Order By Fk_Info_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
        InfoSelect=InfoSelect&"<option value='"&Rs("Fk_Info_Id")&"'>"&Rs("Fk_Info_Name")&"</option>"
		Rs.MoveNext
	Wend
	Rs.Close
%>
<div id="BoxTop" style="width:900px;">标签生成器[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<ul class="BoxNav">
    	<li class="check bnr" id="s1" onclick="ClickBoxNav('1');$('#Template').text('');">常规标签列表</li>
    	<li class="bnr" id="s2" onclick="ClickBoxNav('2');$('#Template').text('');">FOR标签列表</li>
    	<li class="bnr" id="s3" onclick="ClickBoxNav('3');$('#Template').text('');">其他标签</li>
    	<li class="bnr" id="s4" onclick="ClickBoxNav('4');$('#Template').text('');">帮助说明</li>
        <div class="Cal"></div>
    </ul>
    <div class="Cal"></div>
<!--常规标签列表-->
<table width="95%" class="tnr" id="t1" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=1');">全站常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=2');">静态页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=3');">信息页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=4');">文章列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=6');">文章页常规标签</a></td>
        </tr>
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=5');">产品列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=7');">产品页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=12');">下载列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=13');">下载页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&Id=8');">留言列表页常规标签</a></td>
    </tr>
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&amp;Id=9');">专题页标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&amp;Id=11');">搜索页标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&amp;Id=14');">招聘页标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','ShowLabel.asp?Type=1&amp;Id=10');">IF标签使用方法</a></td>
        <td align="center">&nbsp;</td>
    </tr>
</table>
<!--FOR标签列表-->
<table width="95%" class="tnr" id="t2" style="display:none" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="30" align="right">菜单标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetNav" name="GetNav" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId1');">
            <option value="{$MenuId$}">当前菜单</option>
            <%=MenuSelect%>
            </select>
            <select name="Label" class="Input" id="ModuleId1">
                <option value="{$ModuleId$}">当前模块</option>
                </select>
            <select name="Label" class="Input" id="Label">
                <option value="1">读取1级</option>
                <option value="2">读取2级</option>
                <option value="3">读取3级</option>
                <option value="4">读取4级</option>
                <option value="5">读取5级</option>
                </select>
            <select name="Label" class="Input" id="Label">
                <option value="0">无需回溯</option>
                <option value="-1">回溯1级</option>
                <option value="-2">回溯2级</option>
                <option value="-3">回溯3级</option>
                <option value="-4">回溯4级</option>
                <option value="-5">回溯5级</option>
                </select>
                <input type="hidden" name="For" value="Nav" />
            <input type="submit" onclick="Sends_Div('GetNav','ShowLabel.asp?Type=2','Template');" class="Button" name="button2" id="button2" value="生 成" />
</form>
            </td>
        </tr>
    <tr>
        <td height="30" align="right">文章列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetArticleList" name="GetArticleList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId2');">
                <option value="{$MenuId$}">当前菜单</option>
                <%=MenuSelect%>
        </select>
        <select name="Label" class="Input" id="ModuleId2">
                <option value="{$ModuleId$}">当前模块</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">无需设置条数（分页模式）</option>
                <%=GetNumSelect%>
        </select><br />
         <select name="Label" class="Input" id="Label">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有文章</option>
                <option value="-1">非推荐文章</option>
                <%=RecommendSelect%>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有文章</option>
                <option value="-1">非专题文章</option>
                <%=SubjectSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">读取全部标题</option>
                <%=GetLenSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有记录</option>
                <option value="1">有题图的记录</option>
        </select>
        <input type="hidden" name="For" value="ArticleList" />
       <input type="submit" onclick="Sends_Div('GetArticleList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
        </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;产品列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetProductList" name="GetProductList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId3');">
                <option value="{$MenuId$}">当前菜单</option>
                <%=MenuSelect%>
        </select>
        <select name="Label" class="Input" id="ModuleId3">
                <option value="{$ModuleId$}">当前模块</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">无需设置条数（分页模式）</option>
                <%=GetNumSelect%>
        </select><br />
         <select name="Label" class="Input" id="Label">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有产品</option>
                <option value="-1">非推荐产品</option>
                <%=RecommendSelect%>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有产品</option>
                <option value="-1">非专题产品</option>
                <%=SubjectSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">读取全部标题</option>
                <%=GetLenSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有记录</option>
                <option value="1">有题图的记录</option>
        </select>
        <input type="hidden" name="For" value="ProductList" />
       <input type="submit" onclick="Sends_Div('GetProductList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;下载列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetDownList" name="GetDownList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId5');">
                <option value="{$MenuId$}">当前菜单</option>
                <%=MenuSelect%>
        </select>
        <select name="Label" class="Input" id="ModuleId5">
                <option value="{$ModuleId$}">当前模块</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">无需设置条数（分页模式）</option>
                <%=GetNumSelect%>
        </select><br />
         <select name="Label" class="Input" id="Label">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有下载</option>
                <option value="-1">非推荐下载</option>
                <%=RecommendSelect%>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">所有下载</option>
                <option value="-1">非专题下载</option>
                <%=SubjectSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">读取全部标题</option>
                <%=GetLenSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有记录</option>
                <option value="1">有题图的记录</option>
        </select>
        <input type="hidden" name="For" value="DownList" />
       <input type="submit" onclick="Sends_Div('GetDownList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
   <tr>
       <td height="30" align="right">留言列表标签生成：</td>
       <td height="30" colspan="4" style="padding:5px;">
<form id="GetGBookList" name="GetGBookList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId7');">
                <option value="{$MenuId$}">当前菜单</option>
                <%=MenuSelect%>
        </select>
        <select name="Label" class="Input" id="ModuleId7">
                <option value="{$ModuleId$}">当前模块</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">无需设置条数（分页模式）</option>
                <%=GetNumSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有记录</option>
                <option value="1">已回复记录</option>
        </select>
         <select name="Label" class="Input" id="Label">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
        <input type="hidden" name="For" value="GBookList" />
       <input type="submit" onclick="Sends_Div('GetGBookList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />            
</form>
       </td>
   </tr>
   <tr>
        <td height="30" align="right">&nbsp;&nbsp;友情链接标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetFriendsList" name="GetFriendsList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
         <select name="Label" class="Input" id="Label">
         <%=FriendsTypeSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="1">LOGO模式</option>
                <option value="2">文字模式</option>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有</option>
                <%=GetNumSelect%>
        </select>
        <input type="hidden" name="For" value="FriendsList" />
       <input type="submit" onclick="Sends_Div('GetFriendsList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
        </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;招聘列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetJobList" name="GetJobList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId8');">
            <option value="{$MenuId$}">当前菜单</option>
            <%=MenuSelect%>
            </select>
            <select name="Label" class="Input" id="ModuleId8">
                <option value="{$ModuleId$}">当前模块</option>
                </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">读取所有</option>
                <%=GetNumSelect%>
        </select>
        <select name="Label" class="Input" id="Label">
                <option value="0">所有招聘状态</option>
                <option value="1">有效</option>
                <option value="2">过期</option>
        </select>
        <input type="hidden" name="For" value="JobList" />
       <input type="submit" onclick="Sends_Div('GetJobList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;专题列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetSubjectList" name="GetSubjectList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <select name="Label" class="Input" id="Label">
                <option value="0">读取所有</option>
                <%=GetNumSelect%>
        </select>
        <input type="hidden" name="For" value="SubjectList" />
       <input type="submit" onclick="Sends_Div('GetSubjectList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;题图输出FOR标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetPicList" name="GetPicList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        <input type="hidden" name="Label" value="图片列表标签" />
        <input type="hidden" name="For" value="PicList" />
       <input type="submit" onclick="Sends_Div('GetPicList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;循环列表标签生成：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetNumList" name="GetNumList" method="post" action="ShowLabel.asp?Type=2" onsubmit="return false;">
        开始数字：<input type="text" name="Label" id="Label" value="1" size="1" class="Input" />
        结束数字：<input type="text" name="Label" id="Label" value="10" size="1" class="Input" />
        <input type="hidden" name="For" value="NumList" />
       <input type="submit" onclick="Sends_Div('GetNumList','ShowLabel.asp?Type=2','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
</table>
<!--其他标签-->
<table width="95%" class="tnr" id="t3" style="display:none" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;轮换代码JS：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetFlash" name="GetFlash" method="post" action="ShowLabel.asp?Type=3" onsubmit="return false;">
        <select name="Label" class="Input" id="Label" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId101');">
                <option value="">请选择菜单</option>
                <%=MenuSelect%>
        </select>
        <select name="Label" class="Input" id="ModuleId101">
                <option value="">请先选择菜单</option>
        </select>
        方案：<input type="text" name="Label" id="Label" size="1" value="1" class="Input" />
        宽度：<input type="text" name="Label" id="Label" size="1" class="Input" />
        高度：<input type="text" name="Label" id="Label" size="1" class="Input" />
        <input type="hidden" name="For" value="Flash" />
       <input type="submit" onclick="Sends_Div('GetFlash','ShowLabel.asp?Type=3','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;独立信息标签：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetInfo" name="GetInfo" method="post" action="ShowLabel.asp?Type=3" onsubmit="return false;">
        <select name="Label" class="Input" id="Label">
                <%=InfoSelect%>
        </select>
        <input type="hidden" name="For" value="Info" />
       <input type="submit" onclick="Sends_Div('GetInfo','ShowLabel.asp?Type=3','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
    <tr>
        <td height="30" align="right">&nbsp;&nbsp;客服浮窗JS：</td>
        <td height="30" colspan="4" style="padding:5px;">
<form id="GetIm" name="GetIm" method="post" action="ShowLabel.asp?Type=3" onsubmit="return false;">
        <input type="hidden" name="For" value="Im" />
       <input type="submit" onclick="Sends_Div('GetIm','ShowLabel.asp?Type=3','Template');" class="Button" name="button" id="button" value="生 成" />     
</form>
        </td>
    </tr>
</table>
<!--帮助说明-->
<table width="95%" class="tnr" id="t4" style="display:none;" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="21%" height="30" align="right">&nbsp;&nbsp;时间格式输出说明：</td>
        <td width="79%" height="30" colspan="4" style="padding:5px;">
        &nbsp;yyyy：4位数年份<br />
        &nbsp;yy：2位数年份<br />
        &nbsp;mm：月<br />
        &nbsp;dd：日<br />
        &nbsp;hh：小时<br />
        &nbsp;nn：分钟<br />
        &nbsp;ss：秒<br />
        &nbsp;可以自由组合格式
        </td>
    </tr>
    <tr>
        <td width="21%" height="30" align="right">&nbsp;&nbsp;搜索参数说明：</td>
        <td width="79%" height="30" colspan="4" style="padding:5px;">
        &nbsp;搜索请用get方式提交到Search/Index.asp<br />
        &nbsp;SearchType：搜索类型，必须是数字，0：文章，1：产品，2：下载<br />
        &nbsp;SearchStr：搜索字符串<br />
        &nbsp;SearchTemplate：为空调用默认模板，输入模板文件名（全部小写，不带后缀）调用指定模板，如search2<br />
        &nbsp;SearchField：搜索字段表单项，IsContent为检索具体内容，另外还可用自定义字段标识，多个用,隔开，例如“IsContent,自定义字段标识,自定义字段标识”<br />
        &nbsp;SearchFieldList：多自定义字段搜索表单项，可用多个select或其他元素组成，name全部命名为“SearchFieldList”,元素value为“自定义字段标识||搜索内容”<br />
        </td>
    </tr>
</table>
<br />
<table width="95%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" colspan="5" align="center">&nbsp;&nbsp;标签生成结果</td>
    </tr>
    <tr>
        <td height="25" colspan="5" id="Template" style="padding:10px; line-height:22px; font-size:14px;"></td>
        </tr>
</table>
</div>
<div id="BoxBottom" style="width:880px;">
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub
%>
<!--#Include File="../Code.asp"-->