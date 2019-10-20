<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/SiteSet.asp
'文件用途：站点设置拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System1",Request.Cookies("FkAdminLimit1"))

Dim Temp2

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call SiteSetBox() '读取站点信息
	Case 2
		Call SiteSetDo() '站点设置操作
	Case 4
		Call TestMail() '测试邮件
	Case 5
		Call SysHiddenBox() '读取功能隐藏信息
	Case 6
		Call SysHiddenDo() '功能隐藏操作
	Case 7
		Call CloseSysHidden() '对本次登录取消功能隐藏
End Select

'==========================================
'函 数 名：SiteSetBox()
'作    用：读取站点信息
'参    数：
'==========================================
Sub SiteSetBox()
	TempArr=Split(Fk_Site_Jpeg,"|-|Fangka|-|")
	Jpeg_Pic=Clng(TempArr(0))
	Jpeg_Pic_w=Clng(TempArr(1))
	Jpeg_Pic_h=Clng(TempArr(2))
	Jpeg_EditPic=Clng(TempArr(3))
	Jpeg_EditPic_w=Clng(TempArr(4))
	Jpeg_EditPic_h=Clng(TempArr(5))
	Jpeg_Water=Clng(TempArr(6))
	Jpeg_WaterText=TempArr(7)
	Jpeg_WaterFont=TempArr(8)
	Jpeg_WaterFontColor=TempArr(9)
	Jpeg_WaterFontWeight=Clng(TempArr(10))
	Jpeg_WaterPic=TempArr(11)
	Jpeg_WaterPicTransparence=Clng(TempArr(12))
	Jpeg_WaterPicBgColor=TempArr(13)
	Jpeg_WaterPosition=Clng(TempArr(14))
	Jpeg_Water_x=Clng(TempArr(15))
	Jpeg_Water_y=Clng(TempArr(16))
	Jpeg_WaterFontSize=Clng(TempArr(17))
%>
<OBJECT id=dlgHelper CLASSID="clsid:3050f819-98b5-11cf-bb82-00aa00bdce0b" WIDTH="0px" HEIGHT="0px"></OBJECT> 
<form id="SystemSet" name="SystemSet" method="post" action="SiteSet.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">站点设置[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<ul class="BoxNav">
    	<li class="check bnr" id="s1" onclick="ClickBoxNav('1');">常规选项</li>
<%If FkAdmin.AdminCheck(5,"1",Fk_Site_SysHidden) Then%>    	<li class="bnr" id="s2" onclick="ClickBoxNav('2');">生成选项</li><%End If%>
<%If FkAdmin.AdminCheck(5,"2",Fk_Site_SysHidden) Then%>    	<li class="bnr" id="s3" onclick="ClickBoxNav('3');">水印缩略选项</li><%End If%>
<%If FkAdmin.AdminCheck(5,"3",Fk_Site_SysHidden) Then%>    	<li class="bnr" id="s4" onclick="ClickBoxNav('4');">上传类型选项</li><%End If%>
<%If FkAdmin.AdminCheck(5,"4",Fk_Site_SysHidden) Then%>    	<li class="bnr" id="s6" onclick="ClickBoxNav('6');">邮件选项</li><%End If%>
<%If FkAdmin.AdminCheck(5,"5",Fk_Site_SysHidden) Then%>    	<li class="bnr" id="s5" onclick="ClickBoxNav('5');">调试选项</li><%End If%>
        <div class="Cal"></div>
    </ul>
    <div class="Cal"></div>
    <!--常规选项-->
	<table width="90%" border="1" class="tnr" id="t1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">站点名称：</td>
            <td width="78%">&nbsp;<input name="Fk_Site_Name" type="text" class="Input" id="Fk_Site_Name" value="<%=Fk_Site_Name%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>站点的名字，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">站点地址：</td>
            <td>&nbsp;<input name="Fk_Site_Url" type="text" class="Input" id="Fk_Site_Url" value="<%=Fk_Site_Url%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>站点链接地址，请输入正确的地址，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">站点关键字：</td>
            <td>&nbsp;<input name="Fk_Site_Keyword" type="text" class="Input" id="Fk_Site_Keyword" value="<%=Fk_Site_Keyword%>" size="50" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于前台首页meta的keywords，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">站点简介：</td>
            <td>&nbsp;<input name="Fk_Site_Description" type="text" class="Input" id="Fk_Site_Description" value="<%=Fk_Site_Description%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入站点的一段文字简介，用于前台首页meta的description，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
<%
	Call FKAdmin.ShowField(0,1,"",Fk_Site_Field,"")
	Call FKAdmin.ShowField(2,1,"",Fk_Site_Field,"")
	Call FKAdmin.ShowField(3,1,"",Fk_Site_Field,"")
	Call FKAdmin.ShowField(1,1,"",Fk_Site_Field,EditorClass)
%>
        <tr>
            <td height="30" align="right" class="MainTableTop">模板：</td>
            <td>&nbsp;<select class="Input" name="Fk_Site_Template" id="Fk_Site_Template">
<%
	Dim ObjFloders,ObjFloder
	Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	Set F=Fso.GetFolder(Server.MapPath("../Skin/"))
	Set ObjFloders=F.Subfolders
	For Each ObjFloder In ObjFloders
%>
                <option value="<%=ObjFloder.Name%>"<%=FKFun.BeSelect(Fk_Site_Template,ObjFloder.Name)%>><%=ObjFloder.Name%></option>
<%
	Next
	Set ObjFloders=Nothing
	Set F=Nothing
	Set Fso=Nothing
%>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择站点的模板目录，如果选择的模板与当前模板不同，设置时会自动重载模板缓存。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">每页条数：</td>
            <td>&nbsp;<select class="Input" name="Fk_Site_PageSize" id="Fk_Site_PageSize">
<%
	For i=1 To 50
%>
                    <option value="<%=i%>"<%=FKFun.BeSelect(Fk_Site_PageSize,i)%>><%=i%>条</option>
<%
	Next
%>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>默认的列表每页条数，用于全局，权重低于模块中的每页条数设置。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">文件名自动生成：</td>
            <td>&nbsp;<input type="radio" name="Fk_Site_ToPinyin" class="Input" id="Fk_Site_ToPinyin" value="0"<%=FKFun.BeCheck(Fk_Site_ToPinyin,0)%> />不自动生成
            <input name="Fk_Site_ToPinyin" class="Input" type="radio" id="Fk_Site_ToPinyin" value="1"<%=FKFun.BeCheck(Fk_Site_ToPinyin,1)%> />自动生成拼音文件名&nbsp;&nbsp;<span class="qbox" title="<p>如果开启此功能，文章、产品、下载添加的时候，会自动将标题转换为拼音并放入文件名输入框，本功能生成的文件名长度会很长，建议关闭。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">非法字符过滤：</td>

            <td>&nbsp;<input type="radio" name="Fk_Site_DelWord" class="Input" id="Fk_Site_DelWord" value="0"<%=FKFun.BeCheck(Fk_Site_DelWord,0)%> />关闭
            <input name="Fk_Site_DelWord" class="Input" type="radio" id="Fk_Site_DelWord" value="1"<%=FKFun.BeCheck(Fk_Site_DelWord,1)%> />开启&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，会对信息、文章、产品、留言等详细内容进行关键词过滤，以达到屏蔽非法字符的目的，非法字符在“过滤字符管理”中进行增减。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">站点开放：</td>
            <td>&nbsp;<input type="radio" name="Fk_Site_Open" class="Input" id="Fk_Site_Open" value="0"<%=FKFun.BeCheck(Fk_Site_Open,0)%> />关闭
            <input name="Fk_Site_Open" class="Input" type="radio" id="Fk_Site_Open" value="1"<%=FKFun.BeCheck(Fk_Site_Open,1)%> />开放&nbsp;&nbsp;<span class="qbox" title="<p>当站点开放选择“关闭”时，访问前台直接显示站点关闭提示，后台可正常操作，但本设置对纯静态生成无效。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">站点关闭提示：</td>
            <td>&nbsp;<textarea name="Fk_Site_CloseStr" cols="50" class="TextArea" id="Fk_Site_CloseStr"><%=Fk_Site_CloseStr%></textarea>&nbsp;&nbsp;<span class="qbox" title="<p>站点关闭时访问网站前台显示的提示。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
    </table>
    <!--生成选项-->
	<table width="90%" border="1" class="tnr" id="t2" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">系统模式：</td>
            <td width="78%">&nbsp;<input type="radio" name="Fk_Site_Html" class="Input" id="Fk_Site_Html" value="0"<%=FKFun.BeCheck(Fk_Site_Html,0)%> onclick="$('.mode0').show();$('.mode1').show();" />仿静态模式（动态）
            <input type="radio" name="Fk_Site_Html" class="Input" id="Fk_Site_Html" value="1"<%=FKFun.BeCheck(Fk_Site_Html,1)%> onclick="$('.mode0').hide();$('.mode1').hide();" />ID模式（动态）
            <input name="Fk_Site_Html" class="Input" type="radio" id="Fk_Site_Html" value="2"<%=FKFun.BeCheck(Fk_Site_Html,2)%> onclick="$('.mode0').hide();$('.mode1').show();" />纯静态模式&nbsp;&nbsp;<span class="qbox" title="<p>动态模式直接通过ASP进行访问，纯静态模式在更新内容后需要手工进行生成操作。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="mode0"<%If Fk_Site_Html>0 Then%> style="display:none"<%End If%>>
            <td height="30" align="right" class="MainTableTop">目录分隔符：</td>
            <td>&nbsp;<input name="Fk_Site_Sign" type="text" class="Input" id="Fk_Site_Sign" value="<%=Fk_Site_Sign%>" size="10" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>动态模式下级联的目录分隔符，可留空。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="mode0"<%If Fk_Site_Html>0 Then%> style="display:none"<%End If%>>
            <td height="30" align="right" class="MainTableTop">分页分隔符：</td>
            <td>&nbsp;<input name="Fk_Site_PageSign" type="text" class="Input" id="Fk_Site_PageSign" value="<%=Fk_Site_PageSign%>" size="10" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>动态模式下级联的分页分隔符，可留空，当设置了目录分隔符必须设置分页分隔符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="mode1"<%If Fk_Site_Html=1 Then%> style="display:none"<%End If%>>
            <td width="22%" height="30" align="right" class="MainTableTop">目录模式：</td>
            <td width="78%">&nbsp;<input type="radio" name="Fk_Site_HtmlType" class="Input" id="Fk_Site_HtmlType" value="0"<%=FKFun.BeCheck(Fk_Site_HtmlType,0)%> />不级联
            <input name="Fk_Site_HtmlType" class="Input" type="radio" id="Fk_Site_HtmlType" value="1"<%=FKFun.BeCheck(Fk_Site_HtmlType,1)%> />级联&nbsp;&nbsp;<span class="qbox" title="<p>目录模式不级联时，多级栏目都单独存在于根目录下，开启级联时，多级栏目会按级别进行目录嵌套。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="mode1"<%If Fk_Site_Html=1 Then%> style="display:none"<%End If%>>
            <td width="22%" height="30" align="right" class="MainTableTop">生成后缀：</td>
            <td width="78%">&nbsp;<select class="Input" name="Fk_Site_HtmlSuffix" id="Fk_Site_HtmlSuffix">
                <option value="0"<%=FKFun.BeSelect(Fk_Site_HtmlSuffix,0)%>>.html</option>
                <option value="1"<%=FKFun.BeSelect(Fk_Site_HtmlSuffix,1)%>>.htm</option>
                <option value="2"<%=FKFun.BeSelect(Fk_Site_HtmlSuffix,2)%>>.shtml</option>
                <option value="3"<%=FKFun.BeSelect(Fk_Site_HtmlSuffix,3)%>>.xml</option>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>设置生成文件的后缀，在动态或者伪静态模式下，也带有生成后缀。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
    </table>
    <!--水印缩略选项-->
	<table width="90%" border="1" class="tnr" id="t3" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
<%
	If FKFun.IsObjInstalled("Persits.Jpeg") Then
%>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">题图缩略：</td>
            <td width="78%">&nbsp;<input type="radio" name="Jpeg_Pic" class="Input" id="Jpeg_Pic" value="0" onclick="$('#s_Pic').hide();"<%=FKFun.BeCheck(Jpeg_Pic,0)%> /> 关闭
            <input name="Jpeg_Pic" class="Input" type="radio" id="Jpeg_Pic" value="1" onclick="$('#s_Pic').show();"<%=FKFun.BeCheck(Jpeg_Pic,1)%> /> 开启&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，会对上传的题图进行缩略，缩略会根据相应比例缩小。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="s_Pic"<%If Jpeg_Pic=0 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">题图缩略配置：</td>
            <td>&nbsp;限宽：<input name="Jpeg_Pic_w" type="text" class="Input" id="Jpeg_Pic_w" value="<%=Jpeg_Pic_w%>" size="10" />
            	&nbsp;限高：<input name="Jpeg_Pic_h" type="text" class="Input" id="Jpeg_Pic_h" value="<%=Jpeg_Pic_h%>" size="10" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">编辑器图缩略：</td>
            <td>&nbsp;<input type="radio" name="Jpeg_EditPic" class="Input" id="Jpeg_EditPic" value="0" onclick="$('#s_EditPic').hide();"<%=FKFun.BeCheck(Jpeg_EditPic,0)%> /> 关闭
            <input name="Jpeg_EditPic" class="Input" type="radio" id="Jpeg_EditPic" value="1" onclick="$('#s_EditPic').show();"<%=FKFun.BeCheck(Jpeg_EditPic,1)%> /> 开启&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，会对上传的题图进行缩略，缩略会根据相应比例缩小。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="s_EditPic"<%If Jpeg_EditPic=0 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">编辑器图缩略配置：</td>
            <td>&nbsp;限宽：<input name="Jpeg_EditPic_w" type="text" class="Input" id="Jpeg_EditPic_w" value="<%=Jpeg_EditPic_w%>" size="10" />
            	&nbsp;限高：<input name="Jpeg_EditPic_h" type="text" class="Input" id="Jpeg_EditPic_h" value="<%=Jpeg_EditPic_h%>" size="10" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">水印：</td>
            <td>&nbsp;<input type="radio" name="Jpeg_Water" class="Input" id="Jpeg_Water" value="0" onclick="$('.s_Water').hide();"<%=FKFun.BeCheck(Jpeg_Water,0)%> /> 关闭
            <input name="Jpeg_Water" class="Input" type="radio" id="Jpeg_Water" value="1" onclick="$('.s_Water').hide();$('.s_WaterText').show();"<%=FKFun.BeCheck(Jpeg_Water,1)%> /> 文字水印
            <input name="Jpeg_Water" class="Input" type="radio" id="Jpeg_Water" value="2" onclick="$('.s_Water').hide();$('.s_WaterPic').show();"<%=FKFun.BeCheck(Jpeg_Water,2)%> /> 图片水印&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，会对所有上传的图片打水印。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText"<%If Jpeg_Water=0 Or Jpeg_Water=2 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">水印文字：</td>
            <td>&nbsp;<input name="Jpeg_WaterText" type="text" class="Input" id="Jpeg_WaterText" value="<%=Jpeg_WaterText%>" />&nbsp;&nbsp;<span class="qbox" title="<p>水印的文字，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText"<%If Jpeg_Water=0 Or Jpeg_Water=2 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">水印字体：</td>
            <td>&nbsp;<select class="Input" name="Jpeg_WaterFont" id="Jpeg_WaterFont">
                    <option value="宋体"<%=FKFun.BeSelect(Jpeg_WaterFont,"宋体")%>>宋体</option>
                    <option value="楷体_GB2312"<%=FKFun.BeSelect(Jpeg_WaterFont,"楷体_GB2312")%>>楷体</option>
                    <option value="仿宋_GB2312"<%=FKFun.BeSelect(Jpeg_WaterFont,"仿宋_GB2312")%>>新宋体</option>
                    <option value="黑体"<%=FKFun.BeSelect(Jpeg_WaterFont,"黑体")%>>黑体</option>
                    <option value="隶书"<%=FKFun.BeSelect(Jpeg_WaterFont,"隶书")%>>隶书</option>
                    <option value="幼圆"<%=FKFun.BeSelect(Jpeg_WaterFont,"幼圆")%>>幼圆</option>
                    <option value="Andale Mono"<%=FKFun.BeSelect(Jpeg_WaterFont,"Andale Mono")%>>Andale Mono</OPTION> 
                    <option value="Arial"<%=FKFun.BeSelect(Jpeg_WaterFont,"Arial")%>>Arial</OPTION> 
                    <option value="Arial Black"<%=FKFun.BeSelect(Jpeg_WaterFont,"Arial Black")%>>Arial Black</OPTION> 
                    <option value="Book Antiqua"<%=FKFun.BeSelect(Jpeg_WaterFont,"Book Antiqua")%>>Book Antiqua</OPTION>
                    <option value="Century Gothic"<%=FKFun.BeSelect(Jpeg_WaterFont,"Century Gothic")%>>Century Gothic</OPTION> 
                    <option value="Comic Sans MS"<%=FKFun.BeSelect(Jpeg_WaterFont,"Comic Sans MS")%>>Comic Sans MS</OPTION>
                    <option value="Courier New"<%=FKFun.BeSelect(Jpeg_WaterFont,"Courier New")%>>Courier New</OPTION>
                    <option value="Georgia"<%=FKFun.BeSelect(Jpeg_WaterFont,"Georgia")%>>Georgia</OPTION>
                    <option value="Impact"<%=FKFun.BeSelect(Jpeg_WaterFont,"Impact")%>>Impact</OPTION>
                    <option value="Tahoma"<%=FKFun.BeSelect(Jpeg_WaterFont,"Tahoma")%>>Tahoma</OPTION>
                    <option value="Times New Roman"<%=FKFun.BeSelect(Jpeg_WaterFont,"Times New Roman")%>>Times New Roman</OPTION>
                    <option value="Trebuchet MS"<%=FKFun.BeSelect(Jpeg_WaterFont,"Trebuchet MS")%>>Trebuchet MS</OPTION>
                    <option value="Script MT Bold"<%=FKFun.BeSelect(Jpeg_WaterFont,"Script MT Bold")%>>Script MT Bold</OPTION>
                    <option value="Stencil"<%=FKFun.BeSelect(Jpeg_WaterFont,"Stencil")%>>Stencil</OPTION>
                    <option value="Verdana"<%=FKFun.BeSelect(Jpeg_WaterFont,"Verdana")%>>Verdana</OPTION>
                    <option value="Lucida Console"<%=FKFun.BeSelect(Jpeg_WaterFont,"Lucida Console")%>>Lucida Console</OPTION>
             </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择水印所用字体。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText"<%If Jpeg_Water=0 Or Jpeg_Water=2 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">水印文字大小：</td>
            <td>&nbsp;<input name="Jpeg_WaterFontSize" type="text" class="Input" id="Jpeg_WaterFontSize" value="<%=Jpeg_WaterFontSize%>" /> px&nbsp;&nbsp;<span class="qbox" title="<p>水印的文字大小，必须是数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText"<%If Jpeg_Water=0 Or Jpeg_Water=2 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">水印文字颜色：</td>
            <td>&nbsp;<input name="Jpeg_WaterFontColor" type="text" class="Input" id="Jpeg_WaterFontColor" value="<%=Jpeg_WaterFontColor%>" onclick="ColorPicker(this);" />&nbsp;&nbsp;<span class="qbox" title="<p>选择水印文字颜色。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText"<%If Jpeg_Water=0 Or Jpeg_Water=2 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">是否粗体：</td>
            <td>&nbsp;<select class="Input" name="Jpeg_WaterFontWeight" id="Jpeg_WaterFontWeight">
                    <option value="0"<%=FKFun.BeSelect(Jpeg_WaterFontWeight,0)%>>不加粗</option>
                    <option value="1"<%=FKFun.BeSelect(Jpeg_WaterFontWeight,1)%>>加粗</option>
             </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择水印文字是否加粗。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterPic"<%If Jpeg_Water=0 Or Jpeg_Water=1 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">水印图片：</td>
            <td>&nbsp;<input name="Jpeg_WaterPic" type="text" class="Input" id="Jpeg_WaterPic" size="60" value="<%=Jpeg_WaterPic%>" />&nbsp;&nbsp;<span class="qbox" title="<p>水印图片，上传或者输入图片相对路径（用“/”开始）。</p>"><img src="Images/help.jpg" /></span><br />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Friends_Logos" name="Fk_Friends_Logos" src="PicUpLoad.asp?Type=2&Form=SystemSet&Input=Jpeg_WaterPic"></iframe></td>
        </tr>
        <tr class="s_Water s_WaterPic"<%If Jpeg_Water=0 Or Jpeg_Water=1 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">图片透明度：</td>
            <td>&nbsp;<input name="Jpeg_WaterPicTransparence" type="text" class="Input" id="Jpeg_WaterPicTransparence" value="<%=Jpeg_WaterPicTransparence%>" size="5" /> %&nbsp;&nbsp;<span class="qbox" title="<p>水印图片透明度，请输入1-100，100为不透明。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterPic"<%If Jpeg_Water=0 Or Jpeg_Water=1 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">图片底色：</td>
            <td>&nbsp;<input name="Jpeg_WaterPicBgColor" type="text" class="Input" id="Jpeg_WaterPicBgColor" value="<%=Jpeg_WaterPicBgColor%>" />&nbsp;&nbsp;<span class="qbox" title="<p>如果需要去掉底色，请输入底色的RGB值。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText s_WaterPic"<%If Jpeg_Water=0 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">坐标起点：</td>
            <td>&nbsp;<select class="Input" name="Jpeg_WaterPosition" id="Jpeg_WaterPosition">
                    <option value="0"<%=FKFun.BeSelect(Jpeg_WaterPosition,0)%>>左上</option>
                    <option value="1"<%=FKFun.BeSelect(Jpeg_WaterPosition,1)%>>左下</option>
                    <option value="2"<%=FKFun.BeSelect(Jpeg_WaterPosition,2)%>>居中</option>
                    <option value="3"<%=FKFun.BeSelect(Jpeg_WaterPosition,3)%>>右上</option>
                    <option value="4"<%=FKFun.BeSelect(Jpeg_WaterPosition,4)%>>右下</option>
             </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择水印坐标起点。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr class="s_Water s_WaterText s_WaterPic"<%If Jpeg_Water=0 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right" class="MainTableTop">坐标位置：</td>
            <td>&nbsp;X：<input name="Jpeg_Water_x" type="text" class="Input" id="Jpeg_Water_x" value="<%=Jpeg_Water_x%>" size="10" />
            	&nbsp;Y：<input name="Jpeg_Water_y" type="text" class="Input" id="Jpeg_Water_y" value="<%=Jpeg_Water_y%>" size="10" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入数字。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
<%
Else
%>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop" style="color:#F00;">说明：</td>
            <td width="78%" style="color:#F00;">&nbsp;您的空间不支持ASPjpeg组件，缩略和水印功能自动关闭！</td>
        </tr>
<%
End If
%>
    </table>
    <!--上传类型选项-->
	<table width="90%" border="1" class="tnr" id="t4" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" colspan="2" align="left" class="MainTableTop" style="color:#F00">&nbsp;&nbsp;&nbsp;注：设置上传类型后，需要刷新页面后才会生效！</td>
            </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">附件上传类型：</td>
            <td width="78%">&nbsp;<input name="Fk_Site_LinkExt" type="text" class="Input" id="Fk_Site_LinkExt" value="<%=Fk_Site_LinkExt%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>多种格式请用英文逗号隔开，请输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">图片上传类型：</td>
            <td width="78%">&nbsp;<input name="Fk_Site_PicExt" type="text" class="Input" id="Fk_Site_PicExt" value="<%=Fk_Site_PicExt%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>多种格式请用英文逗号隔开，请输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">FLASH上传类型：</td>
            <td width="78%">&nbsp;<input name="Fk_Site_FlashExt" type="text" class="Input" id="Fk_Site_FlashExt" value="<%=Fk_Site_FlashExt%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>多种格式请用英文逗号隔开，请输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">流媒体上传类型：</td>
            <td width="78%">&nbsp;<input name="Fk_Site_MediaExt" type="text" class="Input" id="Fk_Site_MediaExt" value="<%=Fk_Site_MediaExt%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>多种格式请用英文逗号隔开，请输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
    </table>
    <!--邮件选项-->
	<table width="90%" border="1" class="tnr" id="t6" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
<%
	If FKFun.IsObjInstalled("JMail.Message") Then
%>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">邮件发送：</td>
            <td width="78%">&nbsp;<input type="radio" name="Fk_Site_Mail" class="Input" id="Fk_Site_Mail" value="0"<%=FKFun.BeCheck(Fk_Site_Mail,0)%> />关闭
            <input name="Fk_Site_Mail" class="Input" type="radio" id="Fk_Site_Mail" value="1"<%=FKFun.BeCheck(Fk_Site_Mail,1)%> />开启&nbsp;&nbsp;<span class="qbox" title="<p>是否开启邮件发送，如果开启，则有新的留言会发送一份邮件到邮箱。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">邮件地址：</td>
            <td width="78%">&nbsp;<input name="Mail_Address" type="text" class="Input" id="Mail_Address" value="<%=Mail_Address%>" size="30" />&nbsp;&nbsp;<span class="qbox" title="<p>发送邮件到邮件地址，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span>&nbsp;&nbsp;<a href="javascript:void(0);" onclick="SetRContent('MailTest','SiteSet.asp?Type=4&Mail='+escape(document.all.Mail_Address.value)+'&Name='+escape(document.all.Mail_Name.value)+'&Pass='+escape(document.all.Mail_Pass.value)+'&Smtp='+escape(document.all.Mail_Smtp.value))">测试邮件</a>&nbsp;&nbsp;&nbsp;<span id="MailTest" style="color:#F00;"></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">帐号：</td>
            <td width="78%">&nbsp;<input name="Mail_Name" type="text" class="Input" id="Mail_Name" value="<%=Mail_Name%>" size="30" />&nbsp;&nbsp;<span class="qbox" title="<p>帐号，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">密码：</td>
            <td width="78%">&nbsp;<input name="Mail_Pass" type="password" class="Input" id="Mail_Pass" value="<%=Mail_Pass%>" size="30" />&nbsp;&nbsp;<span class="qbox" title="<p>密码，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">SMTP服务器：</td>
            <td width="78%">&nbsp;<input name="Mail_Smtp" type="text" class="Input" id="Mail_Smtp" value="<%=Mail_Smtp%>" size="30" />&nbsp;&nbsp;<span class="qbox" title="<p>SMTP服务器，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
<%
Else
%>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop" style="color:#F00;">说明：</td>
            <td width="78%" style="color:#F00;">&nbsp;您的空间不支持JMail组件，邮件发送功能自动关闭！</td>
        </tr>
<%
End If
%>
    </table>
    <!--调试选项-->
	<table width="90%" border="1" class="tnr" id="t5" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">模板调试：</td>
            <td width="78%">&nbsp;<input type="radio" name="Fk_Site_SkinTest" class="Input" id="Fk_Site_SkinTest" value="0"<%=FKFun.BeCheck(Fk_Site_SkinTest,0)%> />关闭
            <input name="Fk_Site_SkinTest" class="Input" type="radio" id="Fk_Site_SkinTest" value="1"<%=FKFun.BeCheck(Fk_Site_SkinTest,1)%> />开启&nbsp;&nbsp;<span class="qbox" title="<p>开启模板调试后，修改模板无需进行重载模板缓存，系统会直接读取页面内容。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">地址是否带Index.asp：</td>
            <td>&nbsp;<input type="radio" name="Fk_Site_Index" class="Input" id="Fk_Site_Index" value="0"<%=FKFun.BeCheck(Fk_Site_Index,0)%> />关闭
            <input name="Fk_Site_Index" class="Input" type="radio" id="Fk_Site_Index" value="1"<%=FKFun.BeCheck(Fk_Site_Index,1)%> />开启&nbsp;&nbsp;<span class="qbox" title="<p>设置动态模式下，链接中是否带index.asp，否则用?的方式调用。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">显示路径：</td>
            <td>&nbsp;<input name="Fk_Site_Dir" type="text" class="Input" id="Fk_Site_Dir" value="<%=Fk_Site_Dir%>" size="20" />&nbsp;&nbsp;<span class="qbox" title="<p>此功能用于SiteDir值与实际路径不同时，在此设置实际路径，从而使系统能正常的使用。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td width="22%" height="30" align="right" class="MainTableTop">编辑器：</td>
            <td width="78%">&nbsp;<select class="Input" name="Fk_Site_Edit" id="Fk_Site_Edit">
                <option value="0"<%=FKFun.BeSelect(Fk_Site_Edit,0)%>>xhEditor </option>
                <option value="1"<%=FKFun.BeSelect(Fk_Site_Edit,1)%>>KindEditor</option>
                <option value="2"<%=FKFun.BeSelect(Fk_Site_Edit,2)%>>eWebEditor免费版</option>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择系统所试用的编辑器。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('SystemSet','SiteSet.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="设 置" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：SiteSetDo()
'作    用：站点设置操作
'参    数：
'==========================================
Sub SiteSetDo()
	Dim OldTemplate,OldHtmlType,OldHtmlSuffix
	OldTemplate=Fk_Site_Template
	OldHtmlType=Fk_Site_HtmlType
	OldHtmlSuffix=Fk_Site_HtmlSuffix
	Fk_Site_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Name")))
	Fk_Site_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Url")))
	Fk_Site_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Keyword")))
	Fk_Site_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Description")))
	Fk_Site_CloseStr=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_CloseStr")))
	Fk_Site_Template=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Template")))
	Fk_Site_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Dir")))
	Mail_Address=FKFun.HTMLEncode(Trim(Request.Form("Mail_Address")))
	Mail_Name=FKFun.HTMLEncode(Trim(Request.Form("Mail_Name")))
	Mail_Pass=FKFun.HTMLEncode(Trim(Request.Form("Mail_Pass")))
	Mail_Smtp=FKFun.HTMLEncode(Trim(Request.Form("Mail_Smtp")))
	Fk_Site_Sign=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_Sign")))
	Fk_Site_PageSign=FKFun.HTMLEncode(Trim(Request.Form("Fk_Site_PageSign")))
	Fk_Site_Open=Trim(Request.Form("Fk_Site_Open"))
	Fk_Site_Html=Trim(Request.Form("Fk_Site_Html"))
	Fk_Site_HtmlType=Trim(Request.Form("Fk_Site_HtmlType"))
	Fk_Site_HtmlSuffix=Trim(Request.Form("Fk_Site_HtmlSuffix"))
	Fk_Site_PageSize=Trim(Request.Form("Fk_Site_PageSize"))
	Fk_Site_ToPinyin=Trim(Request.Form("Fk_Site_ToPinyin"))
	Fk_Site_DelWord=Trim(Request.Form("Fk_Site_DelWord"))
	Fk_Site_SkinTest=Trim(Request.Form("Fk_Site_SkinTest"))
	Fk_Site_Index=Trim(Request.Form("Fk_Site_Index"))
	Fk_Site_Edit=Trim(Request.Form("Fk_Site_Edit"))
	Fk_Site_Mail=Trim(Request.Form("Fk_Site_Mail"))
	Fk_Site_LinkExt=LCase(FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Site_LinkExt"))," ","")))
	Fk_Site_PicExt=LCase(FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Site_PicExt"))," ","")))
	Fk_Site_FlashExt=LCase(FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Site_FlashExt"))," ","")))
	Fk_Site_MediaExt=LCase(FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Site_MediaExt"))," ","")))
	Fk_Site_Field=FKAdmin.GetFieldData(1,"")
	Fk_Site_LinkExt=FKFun.ReplaceExt(Fk_Site_LinkExt)
	Fk_Site_PicExt=FKFun.ReplaceExt(Fk_Site_PicExt)
	Fk_Site_FlashExt=FKFun.ReplaceExt(Fk_Site_FlashExt)
	Fk_Site_MediaExt=FKFun.ReplaceExt(Fk_Site_MediaExt)
	If Not FKFun.IsObjInstalled("JMail.Message") Then
		Fk_Site_Mail=0
	Else
		If Fk_Site_Mail=1 Then
			Call FKFun.ShowString(Mail_Address,1,50,0,"请输入邮件地址！","邮件地址不能大于50个字符！")
			Call FKFun.ShowString(Mail_Name,1,50,0,"请输入帐号！","帐号不能大于50个字符！")
			Call FKFun.ShowString(Mail_Pass,1,50,0,"请输入密码！","密码不能大于50个字符！")
			Call FKFun.ShowString(Mail_Smtp,1,50,0,"请输入SMTP服务器！","SMTP服务器不能大于50个字符！")
		End If
	End If
	Fk_Site_MailStr=Mail_Address&"||"&Mail_Name&"||"&Mail_Pass&"||"&Mail_Smtp
	If FKFun.IsObjInstalled("Persits.Jpeg") Then
		Jpeg_Pic=Trim(Request.Form("Jpeg_Pic"))
		Jpeg_Pic_w=Trim(Request.Form("Jpeg_Pic_w"))
		Jpeg_Pic_h=Trim(Request.Form("Jpeg_Pic_h"))
		Jpeg_EditPic=Trim(Request.Form("Jpeg_EditPic"))
		Jpeg_EditPic_w=Trim(Request.Form("Jpeg_EditPic_w"))
		Jpeg_EditPic_h=Trim(Request.Form("Jpeg_EditPic_h"))
		Jpeg_Water=Trim(Request.Form("Jpeg_Water"))
		Jpeg_WaterText=Trim(Request.Form("Jpeg_WaterText"))
		Jpeg_WaterFont=Trim(Request.Form("Jpeg_WaterFont"))
		Jpeg_WaterFontColor=FKFun.HTMLEncode(Trim(Request.Form("Jpeg_WaterFontColor")))
		Jpeg_WaterFontWeight=Trim(Request.Form("Jpeg_WaterFontWeight"))
		Jpeg_WaterPic=FKFun.HTMLEncode(Trim(Request.Form("Jpeg_WaterPic")))
		Jpeg_WaterPicTransparence=Trim(Request.Form("Jpeg_WaterPicTransparence"))
		Jpeg_WaterPicBgColor=FKFun.HTMLEncode(Trim(Request.Form("Jpeg_WaterPicBgColor")))
		Jpeg_WaterPosition=Trim(Request.Form("Jpeg_WaterPosition"))
		Jpeg_Water_x=Trim(Request.Form("Jpeg_Water_x"))
		Jpeg_Water_y=Trim(Request.Form("Jpeg_Water_y"))
		Jpeg_WaterFontSize=Trim(Request.Form("Jpeg_WaterFontSize"))
		Call FKFun.ShowNum(Jpeg_Pic,"请选择是否开启题图缩略！")
		If Jpeg_Pic=1 Then
			Call FKFun.ShowNum(Jpeg_Pic_w,"请输入题图缩略宽度且必须是数字！")
			Call FKFun.ShowNum(Jpeg_Pic_h,"请输入题图缩略高度且必须是数字！")
			If Jpeg_Pic_w<80 Or Jpeg_Pic_h<80 Then
				Call FKFun.ShowErr("题图缩略高度和宽度都必须大于80！",2)
			End If
		End If
		Call FKFun.ShowNum(Jpeg_Pic,"请选择是否开启编辑器图缩略！")
		If Jpeg_EditPic=1 Then
			Call FKFun.ShowNum(Jpeg_EditPic_w,"请输入编辑器图缩略宽度且必须是数字！")
			Call FKFun.ShowNum(Jpeg_EditPic_h,"请输入编辑器图缩略高度且必须是数字！")
			If Jpeg_EditPic_w<80 Or Jpeg_EditPic_h<80 Then
				Call FKFun.ShowErr("编辑器图缩略高度和宽度都必须大于80！",2)
			End If
		End If
		Call FKFun.ShowNum(Jpeg_Water,"请选择是否开启图片水印！")
		If Jpeg_Water=1 Then
			Call FKFun.ShowString(Jpeg_WaterText,1,50,0,"请输入水印文字！","水印文字不能大于50个字符！")
			Call FKFun.ShowString(Jpeg_WaterFont,1,50,0,"请选择水印字体！","水印字体不能大于50个字符！")
			Call FKFun.ShowString(Jpeg_WaterFontColor,1,50,0,"请选择水印文字颜色！","水印文字颜色不能大于50个字符！")
			Call FKFun.ShowNum(Jpeg_WaterFontWeight,"请选择水印文字是否加粗！")
			Call FKFun.ShowNum(Jpeg_WaterFontSize,"请输入字体大小！")
		ElseIf Jpeg_Water=2 Then
			Call FKFun.ShowString(Jpeg_WaterPic,1,255,0,"请上传或输入水印图片地址！","水印图片地址不能大于255个字符！")
			Call FKFun.ShowString(Jpeg_WaterPicBgColor,0,50,0,"请选择水印图片底色！","水印图片底色不能大于50个字符！")
			Call FKFun.ShowNum(Jpeg_WaterPicTransparence,"请选择水印图片透明度！")
		End If
		If Jpeg_Water>0 Then
			Call FKFun.ShowNum(Jpeg_WaterPosition,"请选择水印坐标起点！")
			Call FKFun.ShowNum(Jpeg_Water_x,"请选择水印坐标X位置！")
			Call FKFun.ShowNum(Jpeg_Water_y,"请选择水印坐标Y位置！")
		End If
		Fk_Site_Jpeg=Jpeg_Pic&"|-|Fangka|-|"&Jpeg_Pic_w&"|-|Fangka|-|"&Jpeg_Pic_h&"|-|Fangka|-|"&Jpeg_EditPic&"|-|Fangka|-|"&Jpeg_EditPic_w&"|-|Fangka|-|"&Jpeg_EditPic_h&"|-|Fangka|-|"&Jpeg_Water&"|-|Fangka|-|"&Jpeg_WaterText&"|-|Fangka|-|"&Jpeg_WaterFont&"|-|Fangka|-|"&Jpeg_WaterFontColor&"|-|Fangka|-|"&Jpeg_WaterFontWeight&"|-|Fangka|-|"&Jpeg_WaterPic&"|-|Fangka|-|"&Jpeg_WaterPicTransparence&"|-|Fangka|-|"&Jpeg_WaterPicBgColor&"|-|Fangka|-|"&Jpeg_WaterPosition&"|-|Fangka|-|"&Jpeg_Water_x&"|-|Fangka|-|"&Jpeg_Water_y&"|-|Fangka|-|"&Jpeg_WaterFontSize
	End If
	Call FKFun.ShowString(Fk_Site_Name,1,50,0,"请输入站点名称！","站点名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Site_Url,1,255,0,"请输入站点地址！","站点地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_Keyword,1,255,0,"请输入站点关键字！","站点关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_Description,1,255,0,"请输入站点介绍！","站点介绍不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_Template,1,50,0,"兄弟啊，你怎么能把模板删完啊！","模板文件夹名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Site_CloseStr,1,1000,0,"请输入站点关闭提示！","请输入站点关闭提示不能大于1000个字符！")
	Call FKFun.ShowString(Fk_Site_Dir,0,50,0,"！","显示路径不能大于50个字符！")
	Call FKFun.ShowString(Fk_Site_LinkExt,1,255,0,"请输入附件上传后缀！","附件上传后缀不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_PicExt,1,255,0,"请输入图片上传后缀！","图片上传后缀不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_FlashExt,1,255,0,"请输入FLASH上传后缀！","FLASH上传后缀不能大于255个字符！")
	Call FKFun.ShowString(Fk_Site_MediaExt,1,255,0,"请输入流媒体上传后缀！","流媒体上传后缀不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Site_Html,"请选择系统模式！")
	Call FKFun.ShowNum(Fk_Site_HtmlType,"请选择目录级联模式！")
	Call FKFun.ShowNum(Fk_Site_HtmlSuffix,"请选择生成文件名后缀！")
	Call FKFun.ShowNum(Fk_Site_Open,"请选择站点是否开放！")
	Call FKFun.ShowNum(Fk_Site_PageSize,"请选择每页条数！")
	Call FKFun.ShowNum(Fk_Site_ToPinyin,"请选是否自动生成拼音文件名！")
	Call FKFun.ShowNum(Fk_Site_DelWord,"请选择是否进行非法字符过滤！")
	Call FKFun.ShowNum(Fk_Site_SkinTest,"请选择是否开启模板调试模式！")
	Call FKFun.ShowNum(Fk_Site_Index,"请选择是否开启动态地址带Index.asp！")
	Call FKFun.ShowNum(Fk_Site_Edit,"请选择编辑器类型！")
	Call FKFun.ShowNum(Fk_Site_Mail,"请选择是否开启邮件发送！")
	If Right(Fk_Site_Url,1)<>"/" Then
		Call FKFun.ShowErr("站点地址请用“/”结束！",2)
	End If
	If Fk_Site_Sign<>"" Or Fk_Site_PageSign<>"" Then
		Call FKFun.ShowString(Fk_Site_Sign,1,50,0,"请输入目录分隔符！","目录分隔符不能大于50个字符！")
		Call FKFun.ShowString(Fk_Site_PageSign,1,50,0,"请输入分页分隔符！","分页分隔符不能大于50个字符！")
		If Fk_Site_Sign=Fk_Site_PageSign Then
			Call FKFun.ShowErr("目录分隔符和分页分隔符不能相同！",2)
		End If
	End If
	Sqlstr="Select * From [Fk_Site]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Site_Name")=Fk_Site_Name
		Rs("Fk_Site_Url")=Fk_Site_Url
		Rs("Fk_Site_Keyword")=Fk_Site_Keyword
		Rs("Fk_Site_Description")=Fk_Site_Description
		Rs("Fk_Site_Open")=Fk_Site_Open
		Rs("Fk_Site_CloseStr")=Fk_Site_CloseStr
		Rs("Fk_Site_Template")=Fk_Site_Template
		Rs("Fk_Site_Html")=Fk_Site_Html
		Rs("Fk_Site_HtmlType")=Fk_Site_HtmlType
		Rs("Fk_Site_HtmlSuffix")=Fk_Site_HtmlSuffix
		Rs("Fk_Site_PageSize")=Fk_Site_PageSize
		Rs("Fk_Site_ToPinyin")=Fk_Site_ToPinyin
		Rs("Fk_Site_DelWord")=Fk_Site_DelWord
		Rs("Fk_Site_SkinTest")=Fk_Site_SkinTest
		Rs("Fk_Site_Dir")=Fk_Site_Dir
		Rs("Fk_Site_Index")=Fk_Site_Index
		Rs("Fk_Site_Edit")=Fk_Site_Edit
		Rs("Fk_Site_Mail")=Fk_Site_Mail
		Rs("Fk_Site_MailStr")=Fk_Site_MailStr
		Rs("Fk_Site_LinkExt")=Fk_Site_LinkExt
		Rs("Fk_Site_PicExt")=Fk_Site_PicExt
		Rs("Fk_Site_FlashExt")=Fk_Site_FlashExt
		Rs("Fk_Site_MediaExt")=Fk_Site_MediaExt
		Rs("Fk_Site_Field")=Fk_Site_Field
		Rs("Fk_Site_Sign")=Fk_Site_Sign
		Rs("Fk_Site_PageSign")=Fk_Site_PageSign
		If FKFun.IsObjInstalled("Persits.Jpeg") Then
			Rs("Fk_Site_Jpeg")=Fk_Site_Jpeg
		End If
		Rs.Update()
		Application.UnLock()
	End If
	Rs.Close
	If OldTemplate<>Fk_Site_Template Then
		Call FKAdmin.ReLoadTemplate(Fk_Site_Template)
	End If
	If OldHtmlType<>Fk_Site_HtmlType Or OldHtmlSuffix<>Fk_Site_HtmlSuffix Then
		Sqlstr="Select Fk_Menu_Id From [Fk_Menu] Order By Fk_Menu_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
			Call FKAdmin.ReLoadModuleUrl(Rs("Fk_Menu_Id"),Fk_Site_HtmlType,Fk_Site_HtmlSuffix)
			Rs.MoveNext
		Wend
		Rs.Close
	End If
	Response.Write("站点设置成功！")
End Sub

'==========================================
'函 数 名：TestMail()
'作    用：测试邮件
'参    数：
'==========================================
Sub TestMail()
	Mail_Address=FKFun.HTMLEncode(Trim(Request.QueryString("Mail")))
	Mail_Name=FKFun.HTMLEncode(Trim(Request.QueryString("Name")))
	Mail_Pass=FKFun.HTMLEncode(Trim(Request.QueryString("Pass")))
	Mail_Smtp=FKFun.HTMLEncode(Trim(Request.QueryString("Smtp")))
	If Mail_Address<>"" And Mail_Name<>"" And Mail_Pass<>"" And Mail_Smtp<>"" Then
		Temp=FKFun.Jmail(Mail_Address,"测试邮件发送","这是一封 <br/> 测试邮件！","GB2312","text/html")
		Response.Write(Temp)
	Else
		Response.Write("请先设置邮件信息！")
	End If
End Sub


'==========================================
'函 数 名：SysHiddenBox()
'作    用：读取功能隐藏信息
'参    数：
'==========================================
Sub SysHiddenBox()
%>
<form id="SysHiddenSet" name="SysHiddenSet" method="post" action="SiteSet.asp?Type=6" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">功能隐藏[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<p style="text-align:center; color:#F00;">设置后刷新页面生效！</p>
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="14%" height="25" align="right">功能隐藏&nbsp;&nbsp;<span class="qbox" title="<p>勾选的功能将会隐藏。</p>"><img src="Images/help.jpg" /></span>：</td>
	        <td width="86%">
                <ul class="triState">
                    <li><span class="title">隐藏全部</span>
                        <ul>
                            <li><span class="title">站点设置</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="1"<%If Instr(Fk_Site_SysHidden,",1,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">生成选项</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="2"<%If Instr(Fk_Site_SysHidden,",2,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">水印缩略选项</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="3"<%If Instr(Fk_Site_SysHidden,",3,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">上传类型选项</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="4"<%If Instr(Fk_Site_SysHidden,",4,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">邮件选项</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="5"<%If Instr(Fk_Site_SysHidden,",5,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">调试选项</label></li>
                                </ul>
                            </li>
                            <li><span class="title">管理首页</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="6"<%If Instr(Fk_Site_SysHidden,",6,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">站点设置</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="7"<%If Instr(Fk_Site_SysHidden,",7,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">管理员管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="8"<%If Instr(Fk_Site_SysHidden,",8,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">权限管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="9"<%If Instr(Fk_Site_SysHidden,",9,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">模板管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="10"<%If Instr(Fk_Site_SysHidden,",10,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">模板标签生成器</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="11"<%If Instr(Fk_Site_SysHidden,",11,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">静态生成</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="12"<%If Instr(Fk_Site_SysHidden,",12,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">数据库操作</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="13"<%If Instr(Fk_Site_SysHidden,",13,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">上传文件管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="14"<%If Instr(Fk_Site_SysHidden,",14,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">过滤字符管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="15"<%If Instr(Fk_Site_SysHidden,",15,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="16"<%If Instr(Fk_Site_SysHidden,",16,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">搜索引擎地图生成</label></li>
                                </ul>
                            </li>
                            <li><span class="title">内容管理</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="17"<%If Instr(Fk_Site_SysHidden,",17,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">菜单管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="18"<%If Instr(Fk_Site_SysHidden,",18,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">菜单模块管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="19"<%If Instr(Fk_Site_SysHidden,",19,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">友情链接管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="20"<%If Instr(Fk_Site_SysHidden,",20,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">友情链接类型管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="21"<%If Instr(Fk_Site_SysHidden,",21,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">广告管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="22"<%If Instr(Fk_Site_SysHidden,",22,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">站内关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="23"<%If Instr(Fk_Site_SysHidden,",23,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">推荐类型管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="24"<%If Instr(Fk_Site_SysHidden,",24,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">专题管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="25"<%If Instr(Fk_Site_SysHidden,",25,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">在线投票管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="26"<%If Instr(Fk_Site_SysHidden,",26,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">独立信息管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="27"<%If Instr(Fk_Site_SysHidden,",27,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">自定义字段管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="28"<%If Instr(Fk_Site_SysHidden,",28,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">客服悬窗代码管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Site_SysHidden" value="29"<%If Instr(Fk_Site_SysHidden,",29,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">留言模型管理</label></li>
                                </ul>
                            </li>
                        </ul>
                    </li>
                </ul>
                <div class="Cal"></div>
            </td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('SysHiddenSet','SiteSet.asp?Type=6',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="设 置" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：SysHiddenDo()
'作    用：功能隐藏操作
'参    数：
'==========================================
Sub SysHiddenDo()
	Fk_Site_SysHidden=",,"&FKFun.HTMLEncode(Trim(Replace(Request.Form("Fk_Site_SysHidden")," ","")))&",,"
	Sqlstr="Update [Fk_Site] Set Fk_Site_SysHidden='"&Fk_Site_SysHidden&"'"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("功能隐藏设置成功！")
End Sub

'==========================================
'函 数 名：CloseSysHidden()
'作    用：对本次登录取消功能隐藏
'参    数：
'==========================================
Sub CloseSysHidden()
	Response.Cookies("CloseSysHidden")="1"
	If Fk_Site_Dir<>"" Then
		Response.Cookies("CloseSysHidden").Path="/"
	End If
%>
<div id="BoxTop" style="width:400px;">对本次登录取消功能隐藏[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<p>&nbsp;</p>
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="center">操作成功，刷新页面后生效！<span class="qbox" title="<p>请输入1-50个字符（两个字符为一个汉字）。</p>"></span></td>
	        </tr>
	    </table>
	<p>&nbsp;</p>
</div>
<div id="BoxBottom" style="width:380px;">
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub
%>
<!--#Include File="../Code.asp"-->