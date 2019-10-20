<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Product.asp
'文件用途：产品管理拉取页面
'版权所有：方卡在线
'==========================================

'定义页面变量
Dim Fk_Product_Title,Fk_Product_Color,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Content,Fk_Product_MUrl,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_PicList,Fk_Product_Menu,Fk_Product_Module,Fk_Product_Click,Fk_Product_FileName,Fk_Product_Recommend,Fk_Product_Field,Fk_Product_Subject,Fk_Product_Template,Fk_Product_Show,Fk_Product_Time,Fk_Product_Url
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Type,Fk_Product_Order,Fk_Module_MName
Dim Temp2
Set FKHtml=New Cls_Html
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode
Fk_Module_MName="产品"

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ProductList() '产品列表
	Case 2
		Call ProductAddForm() '添加产品表单
	Case 3
		Call ProductAddDo() '执行添加产品
	Case 4
		Call ProductEditForm() '修改产品表单
	Case 5
		Call ProductEditDo() '执行修改产品
	Case 6
		Call ProductDelDo() '执行删除产品
	Case 7
		Call ListDelDo() '执行批量删除产品
	Case 8
		Call ProductMove() '执行批量移动产品
	Case 9
		Call ProductOrderSet() '执行产品排序
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：ProductList()
'作    用：产品列表
'参    数：
'==========================================
Sub ProductList()
	Session("NowPage")=FkFun.GetNowUrl()
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限！",2)
	End If
	PageNow=FKFun.GetNumeric("Page",1)
	Sqlstr="Select Fk_Module_Name,Fk_Module_Menu,Fk_Module_MName From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		If Rs("Fk_Module_MName")<>"" Then
			Fk_Module_MName=Rs("Fk_Module_MName")
		End If
	Else
		Rs.Close
		Call FKFun.ShowErr("模块不存在！",2)
	End If
	Rs.Close
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Product.asp?Type=2&ModuleId=<%=Fk_Module_Id%>');">添加新<%=Fk_Module_MName%></a></li>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    “<%=Fk_Module_Name%>”模块&nbsp;&nbsp;<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" />&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','Product.asp?Type=1&ModuleId=<%=Fk_Module_Id%>&SearchStr='+escape($('#SearchStr').val()));" name="S" Id="S" value="  查询  " />
    &nbsp;&nbsp;快速通道：<select name="D1" id="D1" onChange="eval(this.options[this.selectedIndex].value);" class="Input">
      <option value="alert('请选择模块');">请选择模块</option>
<%
Call FKAdmin.GetModuleList(0,Fk_Module_Menu,0,Fk_Module_Id,"")
%>
</select>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Product.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop"><%=Fk_Module_MName%>名称</td>
            <td align="center" class="ListTdTop">文件名</td>
            <td align="center" class="ListTdTop"><%=Fk_Module_MName%>参数</td>
            <td align="center" class="ListTdTop">点击量</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">添加时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_Color,Fk_Product_Url,Fk_Product_FileName,Fk_Product_Show,Fk_Product_Pic,Fk_Product_Template,Fk_Product_Recommend,Fk_Product_Subject,Fk_Product_Click,Fk_Product_Order,Fk_Product_Time From [Fk_Product] Where Fk_Product_Module="&Fk_Module_Id&""
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Fk_Product_Title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By Fk_Product_Order Asc,Fk_Product_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		j=(PageNow-1)*PageSizes+1
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
			If Rs("Fk_Product_Template")>0 Then
				Sqlstr="Select Fk_Template_Name From [Fk_Template] Where Fk_Template_Id=" & Rs("Fk_Product_Template")
				Rs2.Open Sqlstr,Conn,1,1
				If Not Rs2.Eof Then
					Fk_Product_Template=Rs2("Fk_Template_Name")
				Else
					Fk_Product_Template="未知模板"
				End If
				Rs2.Close
			Else
				Fk_Product_Template="默认模板"
			End If
			Fk_Product_Recommend=""
			If Rs("Fk_Product_Recommend")<>"" Then
				TempArr=Split(Rs("Fk_Product_Recommend"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select Fk_Recommend_Name From [Fk_Recommend] Where Fk_Recommend_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Product_Recommend=Fk_Product_Recommend&"<p>"&Rs2("Fk_Recommend_Name")&"</p>"
						End If
						Rs2.Close
					End If
				Next
			End If
			Fk_Product_Subject=""
			If Rs("Fk_Product_Subject")<>"" Then
				TempArr=Split(Rs("Fk_Product_Subject"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select Fk_Subject_Name From [Fk_Subject] Where Fk_Subject_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Product_Subject=Fk_Product_Subject&"<p>"&Rs2("Fk_Subject_Name")&"</p>"
						End If
						Rs2.Close
					End If
				Next
			End If
%>
        <tr>
            <td height="20" align="center"><%=j%></td>
            <td align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("Fk_Product_Id")%>" id="List<%=Rs("Fk_Product_Id")%>" /></td>
            <td>&nbsp;&nbsp;<%=Rs("Fk_Product_Title")%><%If Rs("Fk_Product_Color")<>"" Then%><span style="color:<%=Rs("Fk_Product_Color")%>">■</span><%End If%><%If Rs("Fk_Product_Url")<>"" Then%>[转向链接]<%End If%></td>
            <td align="center"><%=Rs("Fk_Product_FileName")%></td>
            <td align="center"><%If Rs("Fk_Product_Show")=1 Then%><span style="color:#000">[显]</span><%Else%><span style="color:#CCC">[隐]</span><%End If%><%If Rs("Fk_Product_Pic")<>"" Then%><span style="color:#F00">[图]</span><%End If%><span class="qbox qbox2" title="<p><%=Fk_Product_Template%></p>">[模]</span><%If Fk_Product_Recommend<>"" Then%><span class="qbox" title="<p><%=Fk_Product_Recommend%></p>">[推]</span><%End If%><%If Fk_Product_Subject<>"" Then%><span class="qbox" title="<p><%=Fk_Product_Subject%></p>">[专]</span><%End If%></td>
            <td align="center"><%=Rs("Fk_Product_Click")%></td>
            <td align="center"><div><input name="OrderCount<%=Rs("Fk_Product_Id")%>" value="<%=Rs("Fk_Product_Order")%>" type="text" class="orderinput" id="OrderCount<%=Rs("Fk_Product_Id")%>" /><span id="order<%=Rs("Fk_Product_Id")%>"><input type="button" class="orderbtn" onclick="AjaxSet('Product.asp?Type=9&Id=<%=Rs("Fk_Product_Id")%>&Order='+$('#OrderCount<%=Rs("Fk_Product_Id")%>').val());" name="S" Id="S" value="设置" /></span></div></td>
            <td align="center"><%=Rs("Fk_Product_Time")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Product.asp?Type=4&Id=<%=Rs("Fk_Product_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Product_Title")%>”？此操作不可逆！','Product.asp?Type=6&Id=<%=Rs("Fk_Product_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
			j=j+1
		Wend
%>
        <tr>
            <td height="30" colspan="9">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)"> 全选
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的<%=Fk_Module_MName%>吗？')){Sends('DelList','Product.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}">
<select name="ProductMove" id="ProductMove" class="Input" onchange="DelIt('确实要移动这部分<%=Fk_Module_MName%>？','Product.asp?Type=8&Id='+this.options[this.options.selectedIndex].value+'&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');">
      <option value="">转移到</option>
<%
Call FKAdmin.GetModuleList(1,Fk_Module_Menu,0,0,"")
%>
</select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("Product.asp?Type=1&ModuleId="&Fk_Module_Id&"&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="9" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
    </form>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：ProductAddForm()
'作    用：添加产品表单
'参    数：
'==========================================
Sub ProductAddForm()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select Fk_Module_Name,Fk_Module_Menu,Fk_Module_MName From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		If Rs("Fk_Module_MName")<>"" Then
			Fk_Module_MName=Rs("Fk_Module_MName")
		End If
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="ProductAdd" name="ProductAdd" method="post" action="Product.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">添加“<%=Fk_Module_Name%>”模块新<%=Fk_Module_MName%>[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<ul class="BoxNav">
    	<li class="check bnr" id="s1" onclick="ClickBoxNav('1');">常规项目</li>
    	<li class="bnr" id="s2" onclick="ClickBoxNav('2');">专题/推荐</li>
    	<li class="bnr" id="s3" onclick="ClickBoxNav('3');">其他选项</li>
        <div class="Cal"></div>
    </ul>
    <div class="Cal"></div>
<!--常规项目-->
<table width="95%" border="1" class="tnr" id="t1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="10%" height="28" align="right"><%=Fk_Module_MName%>名称：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Title"<%If Fk_Site_ToPinyin=1 Then%> onchange="GetPinyin('Fk_Product_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> type="text" class="Input" id="Fk_Product_Title" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p><%=Fk_Module_MName%>名称，不能为空，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span>
        颜色：<input name="Fk_Product_Color" type="text" id="Fk_Product_Color" class="Input" size="10" onclick="$(this).selectColor();">&nbsp;&nbsp;<span class="qbox" title="<p>点击输入框可以选择颜色，如需取消颜色，点下，在弹出颜色框中直接点取消即可。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">关键字：</td>
        <td width="42%">&nbsp;<input name="Fk_Product_Keyword" type="text" class="Input" id="Fk_Product_Keyword" size="30" />&nbsp;<input type="submit" onclick="SendGet('ProductAdd','Ajax.asp?Type=2&Id=2','Fk_Product_Keyword');" class="Button" name="button" id="button" value="提 取" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于页面meta的keywords，请输入1-255个字符（两个字符为一个汉字），提取按钮是根据“关键字管理”中的关键字来提取，如果在“关键字管理”中没有的关键字是不会被提取的。2%≦提取关键字密度≦8%，不在这个密度里的不提取！</p>"><img src="Images/help.jpg" /></span></td>
        <td width="8%" align="right">描述：</td>
        <td width="40%">&nbsp;<input name="Fk_Product_Description" type="text" class="Input" id="Fk_Product_Description" size="30" />&nbsp;<input type="submit" onclick="SendGet('ProductAdd','Ajax.asp?Type=3&Id=2','Fk_Product_Description');" class="Button" name="button" id="button" value="提 取" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入<%=Fk_Module_MName%>的一段文字简介，用于页面meta的description，请输入1-255个字符（两个字符为一个汉字），点提取则自动提取<%=Fk_Module_MName%>内容的一部分。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">题图：</td>
        <td colspan="3">
        <div id="st" class="Cal"></div>
        <input name="Fk_Product_Pic" type="hidden" class="Input" id="Fk_Pic" />
        <input name="Fk_Product_PicBig" type="hidden" class="Input" id="Fk_PicBig" />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Product_Pics" name="Fk_Product_Pics" src="PicUpLoad.asp?Type=3&Form=ProductAdd&Input=Fk_Pic-Fk_PicBig"></iframe>&nbsp;&nbsp;<span class="qbox" title="<p>题图可以上传多个，黄色边框的为封面题图，点击非封面题图的图片可以设置该图片为封面题图。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">文件名：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_FileName" type="text" class="Input" id="Fk_Product_FileName" value="" size="50" />&nbsp;*一旦确立不可修改&nbsp;&nbsp;<span class="qbox" title="<p>文件名，请输入1-50个字符，必须是字母或数字。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	Call FKAdmin.ShowField(0,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
	Call FKAdmin.ShowField(2,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
	Call FKAdmin.ShowField(3,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,"")
%>
    <tr>
        <td height="28" align="right">模板：</td>
        <td colspan="3">&nbsp;<select name="Fk_Product_Template" class="Input" id="Fk_Product_Template">
            <option value="0">默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本页面使用的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right"><%=Fk_Module_MName%>显示：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Show" class="Input" type="radio" id="Fk_Product_Show" value="1" checked="checked" />显示
        <input type="radio" class="Input" name="Fk_Product_Show" id="Fk_Product_Show" value="0" />不显示&nbsp;&nbsp;<span class="qbox" title="<p>选择该<%=Fk_Module_MName%>是否显示。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right"><%=Fk_Module_MName%>内容&nbsp;&nbsp;<span class="qbox" title="<p><%=Fk_Module_MName%>内容，请输入10个字符以上（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span>：</td>
        <td colspan="3"><textarea name="Fk_Product_Content" style="width:100%;" class="<%=EditorClass%>" rows="15" id="Fk_Product_Content"></textarea></td>
    </tr>
<%
	Call FKAdmin.ShowField(1,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Null,EditorClass)
%>
</table>
<!--专题/推荐-->
<table width="95%" border="1" class="tnr" id="t2" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">推荐：</td>
        <td>&nbsp;<select name="Fk_Product_Recommend" size="10" multiple="multiple" class="TextArea" id="Fk_Product_Recommend">
            <option value="0">无推荐</option>
<%
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Recommend_Id")%>"><%=Rs("Fk_Recommend_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择<%=Fk_Module_MName%>推荐类型，可按住CTRL键用鼠标左键多选。</p>"><img src="Images/help.jpg" /></span></td>
        <td align="right">专题：</td>
        <td>&nbsp;<select name="Fk_Product_Subject" class="TextArea" size="10" multiple="multiple" id="Fk_Product_Subject">
            <option value="0">无专题</option>
<%
	Sqlstr="Select Fk_Subject_Id,Fk_Subject_Name From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Subject_Id")%>"><%=Rs("Fk_Subject_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择<%=Fk_Module_MName%>归属专题，可按住CTRL键用鼠标左键多选。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
</table>
<!--其他选项-->
<table width="95%" border="1" class="tnr" id="t3" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">添加时间：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Time" value="<%=Now()%>" type="text" class="Input" id="Fk_Product_Time" size="50" onClick="javascript:ShowCalendar(this.id)" />&nbsp;&nbsp;<span class="qbox" title="<p>添加时间，可自主修改，但必须按格式书写。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">转向链接：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Url" type="text" class="Input" id="Fk_Product_Url" size="50" />&nbsp;*正常<%=Fk_Module_MName%>请留空&nbsp;&nbsp;<span class="qbox" title="<p>转向链接，如果在这里输入网址，则列表上自动指向该网址。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:880px;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('ProductAdd','Product.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ProductAddDo
'作    用：执行添加产品
'参    数：
'==============================
Sub ProductAddDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("您无此权限！",2)
	End If
	Sqlstr="Select Fk_Module_MName,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		If Rs("Fk_Module_MName")<>"" Then
			Fk_Module_MName=Rs("Fk_Module_MName")
		End If
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",2)
	End If
	Rs.Close
	Fk_Product_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Title")))
	Fk_Product_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Color")))
	Fk_Product_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Keyword")))
	Fk_Product_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Description")))
	Fk_Product_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Url")))
	Fk_Product_Content=Request.Form("Fk_Product_Content")
	Fk_Product_PicList=FKFun.HTMLEncode(Trim(Request.Form("Fk_PicList")))
	Fk_Product_PicList=Replace(Fk_Product_PicList,", ","||")
	Fk_Product_PicList=Replace(Fk_Product_PicList,"|||-_-|","|-_-|")
	Fk_Product_PicList=Replace(Fk_Product_PicList,"|-_-|||","|-_-|")
	If Right(Fk_Product_PicList,5)="|-_-|" Then
		Fk_Product_PicList=Left(Fk_Product_PicList,Len(Fk_Product_PicList)-5)
	End If
	Fk_Product_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Pic")))
	Fk_Product_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_PicBig")))
	Fk_Product_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_FileName")))
	Fk_Product_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Recommend"))," ",""))&","
	Fk_Product_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Subject"))," ",""))&","
	Fk_Product_Template=Trim(Request.Form("Fk_Product_Template"))
	Fk_Product_Show=Trim(Request.Form("Fk_Product_Show"))
	Fk_Product_Time=Trim(Request.Form("Fk_Product_Time"))
	If Fk_Product_PicList<>"" And (Fk_Product_Pic="" Or Fk_Product_PicBig="") Then
		Call FKFun.ShowErr("请在上传的题图列表中选择封面图片！",2)
	End If
	Call FKFun.ShowString(Fk_Product_Title,1,255,0,"请输入"&Fk_Module_MName&"名称！",""&Fk_Module_MName&"名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Keyword,0,255,2,"请输入"&Fk_Module_MName&"关键字！",""&Fk_Module_MName&"关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Description,0,255,2,"请输入"&Fk_Module_MName&"描述！",""&Fk_Module_MName&"描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Url,0,255,2,"请输入"&Fk_Module_MName&"转向链接！",""&Fk_Module_MName&"转向链接不能大于255个字符！")
	If Fk_Product_Url="" Then
		Call FKFun.ShowString(Fk_Product_Content,10,1,1,"请输入"&Fk_Module_MName&"内容，不少于10个字符！",""&Fk_Module_MName&"内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Product_Pic,0,255,2,"请输入"&Fk_Module_MName&"题图路径！",""&Fk_Module_MName&"题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_PicBig,0,255,2,"请输入"&Fk_Module_MName&"题图路径！",""&Fk_Module_MName&"题图大图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_FileName,0,50,2,"请输入"&Fk_Module_MName&"文件名！",""&Fk_Module_MName&"文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Product_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Product_Show,"请选择"&Fk_Module_MName&"是否显示！")
	If Fk_Product_Time="" Then
		Fk_Product_Time=Now()
	End If
	Fk_Product_Field=FKAdmin.GetFieldData(0,"(Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')")
	Call FKAdmin.CheckDandF("file",Fk_Product_FileName,0)
	If Fk_Site_DelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Product_Content=Replace(Fk_Product_Content,Temp,"**")
				Fk_Product_Title=Replace(Fk_Product_Title,Temp,"**")
				Fk_Product_Keyword=Replace(Fk_Product_Keyword,Temp,"**")
				Fk_Product_Description=Replace(Fk_Product_Description,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_Color,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Content,Fk_Product_MUrl,Fk_Product_PicList,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_Menu,Fk_Product_Module,Fk_Product_Click,Fk_Product_FileName,Fk_Product_Recommend,Fk_Product_Subject,Fk_Product_Field,Fk_Product_Template,Fk_Product_Show,Fk_Product_Url,Fk_Product_Time From [Fk_Product] Where Fk_Product_Module="&Fk_Module_Id&" And (Fk_Product_Title='"&Fk_Product_Title&"'"
	If Fk_Product_FileName<>"" Then
		Sqlstr=Sqlstr&" Or Fk_Product_FileName='"&Fk_Product_FileName&"'"
	End If
	Sqlstr=Sqlstr&")"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Product_Title")=Fk_Product_Title
		Rs("Fk_Product_Color")=Fk_Product_Color
		Rs("Fk_Product_Keyword")=Fk_Product_Keyword
		Rs("Fk_Product_Field")=Fk_Product_Field
		Rs("Fk_Product_Description")=Fk_Product_Description
		Rs("Fk_Product_Url")=Fk_Product_Url
		Rs("Fk_Product_Show")=Fk_Product_Show
		Rs("Fk_Product_Pic")=Fk_Product_Pic
		Rs("Fk_Product_PicBig")=Fk_Product_PicBig
		Rs("Fk_Product_PicList")=Fk_Product_PicList
		Rs("Fk_Product_MUrl")=""
		Rs("Fk_Product_Content")=Fk_Product_Content
		Rs("Fk_Product_Recommend")=Fk_Product_Recommend
		Rs("Fk_Product_Subject")=Fk_Product_Subject
		Rs("Fk_Product_Module")=Fk_Module_Id
		Rs("Fk_Product_Menu")=Fk_Module_Menu
		Rs("Fk_Product_FileName")=Fk_Product_FileName
		Rs("Fk_Product_Template")=Fk_Product_Template
		Rs("Fk_Product_Time")=Fk_Product_Time
		Rs.Update()
		Application.UnLock()
		Response.Write("新"&Fk_Module_MName&"添加成功！")
	Else
		Response.Write("该"&Fk_Module_MName&"名称已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：ProductEditForm()
'作    用：修改产品表单
'参    数：
'==========================================
Sub ProductEditForm()
	Dim picNameTemp
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Product_Title,Fk_Product_Color,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Module,Fk_Product_Content,Fk_Product_MUrl,Fk_Product_PicList,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_Click,Fk_Product_FileName,Fk_Product_Recommend,Fk_Product_Subject,Fk_Product_Field,Fk_Product_Template,Fk_Product_Show,Fk_Product_Url,Fk_Product_Time From [Fk_Product] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Product_Title=Rs("Fk_Product_Title")
		Fk_Product_Color=Rs("Fk_Product_Color")
		Fk_Product_Keyword=Rs("Fk_Product_Keyword")
		Fk_Product_Description=Rs("Fk_Product_Description")
		Fk_Product_Content=Rs("Fk_Product_Content")
		Fk_Product_Module=Rs("Fk_Product_Module")
		Fk_Product_Click=Rs("Fk_Product_Click")
		Fk_Product_Url=Rs("Fk_Product_Url")
		Fk_Product_Pic=Rs("Fk_Product_Pic")
		Fk_Product_PicBig=Rs("Fk_Product_PicBig")
		Fk_Product_PicList=Rs("Fk_Product_PicList")
		Fk_Product_Recommend=Rs("Fk_Product_Recommend")
		Fk_Product_Subject=Rs("Fk_Product_Subject")
		Fk_Product_Show=Rs("Fk_Product_Show")
		Fk_Product_Template=Rs("Fk_Product_Template")
		Fk_Product_FileName=Rs("Fk_Product_FileName")
		Fk_Product_Time=Rs("Fk_Product_Time")
		If IsNull(Rs("Fk_Product_Field")) Or Rs("Fk_Product_Field")="" Then
			Fk_Product_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Product_Field=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
		End If
		If Fk_Product_PicList<>"" Then
			TempArr=Split(Fk_Product_PicList,"|-_-|")
		End If
	Else
		Call FKFun.ShowErr("未找到产品，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Product_Module,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select Fk_Module_MName From [Fk_Module] Where Fk_Module_Id=" & Fk_Product_Module
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		If Rs("Fk_Module_MName")<>"" Then
			Fk_Module_MName=Rs("Fk_Module_MName")
		End If
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="ProductEdit" name="ProductEdit" method="post" action="Product.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">修改<%=Fk_Module_MName%>[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<ul class="BoxNav">
    	<li class="check bnr" id="s1" onclick="ClickBoxNav('1');">常规项目</li>
    	<li class="bnr" id="s2" onclick="ClickBoxNav('2');">专题/推荐</li>
    	<li class="bnr" id="s3" onclick="ClickBoxNav('3');">其他选项</li>
        <div class="Cal"></div>
    </ul>
    <div class="Cal"></div>
<!--常规项目-->
<table width="95%" border="1" class="tnr" id="t1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="10%" height="28" align="right"><%=Fk_Module_MName%>名称：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Title" value="<%=Fk_Product_Title%>"<%If Fk_Site_ToPinyin=1 Then%> onchange="GetPinyin('Fk_Product_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> type="text" class="Input" id="Fk_Product_Title" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p><%=Fk_Module_MName%>名称，不能为空，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span>
        颜色：<input name="Fk_Product_Color" value="<%=Fk_Product_Color%>" type="text" id="Fk_Product_Color" class="Input" size="10" onclick="$(this).selectColor();">&nbsp;&nbsp;<span class="qbox" title="<p>点击输入框可以选择颜色，如需取消颜色，点下，在弹出颜色框中直接点取消即可。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">关键字：</td>
        <td width="42%">&nbsp;<input name="Fk_Product_Keyword" value="<%=Fk_Product_Keyword%>" type="text" class="Input" id="Fk_Product_Keyword" size="30" />&nbsp;<input type="submit" onclick="SendGet('ProductEdit','Ajax.asp?Type=2&Id=2','Fk_Product_Keyword');" class="Button" name="button" id="button" value="提 取" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于页面meta的keywords，请输入1-255个字符（两个字符为一个汉字），提取按钮是根据“关键字管理”中的关键字来提取，如果在“关键字管理”中没有的关键字是不会被提取的。2%≦提取关键字密度≦8%，不在这个密度里的不提取！</p>"><img src="Images/help.jpg" /></span></td>
        <td width="8%" align="right">描述：</td>
        <td width="40%">&nbsp;<input name="Fk_Product_Description" value="<%=Fk_Product_Description%>" type="text" class="Input" id="Fk_Product_Description" size="30" />&nbsp;<input type="submit" onclick="SendGet('ProductEdit','Ajax.asp?Type=3&Id=2','Fk_Product_Description');" class="Button" name="button" id="button" value="提 取" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入<%=Fk_Module_MName%>的一段文字简介，用于页面meta的description，请输入1-255个字符（两个字符为一个汉字），点提取则自动提取<%=Fk_Module_MName%>内容的一部分。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">题图：</td>
        <td colspan="3">
<%
	i=100000
	If Fk_Product_PicList<>"" Then
		For Each Temp In TempArr
			If UBound(Split(Temp,"||"))>=2 Then
				picNameTemp=Split(Temp,"||")(2)
			Else
				picNameTemp=""
			End If
%>
        <div id="p<%=i%>" class="picList<%If Fk_Product_Pic=Split(Temp,"||")(0) Then%> picCheck<%End If%>"><img src="<%=Split(Temp,"||")(0)%>" width="30" height="30" class="qbox qbox2" title="<img src='<%=Split(Temp,"||")(0)%>' width=190 height=180 />" onclick="clickPic(<%=i%>,'Fk_Pic','Fk_PicBig')" /><br /><input name="Fk_PicList" type="hidden" class="Input" id="Fk_PicList<%=i%>a" value="<%=Split(Temp,"||")(0)%>" /><input name="Fk_PicList" type="hidden" class="Input" id="Fk_PicList<%=i%>b" value="<%=Split(Temp,"||")(1)%>" /><input name="Fk_PicList" type="text" class="Input" id="Fk_PicList<%=i%>t" value="<%=picNameTemp%>" style="width:60px;" /><input name="Fk_PicList" value="|-_-|" type="hidden" class="Input" id="Fk_PicList" /><br /><a href="javascript:void(0);" onclick="unPic(<%=i%>)" title="删除">删除</a></div>
<%
			i=i+1
		Next
	End If
%>
        <div id="st" class="Cal"></div>
        <input name="Fk_Product_Pic" value="<%=Fk_Product_Pic%>" type="hidden" class="Input" id="Fk_Pic" />
        <input name="Fk_Product_PicBig" value="<%=Fk_Product_PicBig%>" type="hidden" class="Input" id="Fk_PicBig" />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Product_Pics" name="Fk_Product_Pics" src="PicUpLoad.asp?Type=3&Form=ProductEdit&Input=Fk_Pic-Fk_PicBig"></iframe>&nbsp;&nbsp;<span class="qbox" title="<p>题图可以上传多个，黄色边框的为封面题图，点击非封面题图的图片可以设置该图片为封面题图。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">文件名：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_FileName" value="<%=Fk_Product_FileName%>" type="text" class="Input" id="Fk_Product_FileName" size="50"<%If Fk_Product_FileName<>"" Then%> disabled="disabled"<%End If%> />&nbsp;*一旦确立不可修改&nbsp;&nbsp;<span class="qbox" title="<p>文件名，请输入1-50个字符，必须是字母或数字。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	Call FKAdmin.ShowField(0,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Product_Module&",%%')",Fk_Product_Field,"")
	Call FKAdmin.ShowField(2,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Product_Module&",%%')",Fk_Product_Field,"")
	Call FKAdmin.ShowField(3,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Product_Module&",%%')",Fk_Product_Field,"")
%>
    <tr>
        <td height="28" align="right">模板：</td>
        <td colspan="3">&nbsp;<select name="Fk_Product_Template" class="Input" id="Fk_Product_Template">
            <option value="0"<%=FKFun.BeSelect(Fk_Product_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Product_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本页面使用的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right"><%=Fk_Module_MName%>显示：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Show" class="Input" type="radio" id="Fk_Product_Show" value="1"<%=FKFun.BeCheck(Fk_Product_Show,1)%> />显示
        <input type="radio" class="Input" name="Fk_Product_Show" id="Fk_Product_Show" value="0"<%=FKFun.BeCheck(Fk_Product_Show,0)%> />不显示&nbsp;&nbsp;<span class="qbox" title="<p>选择该<%=Fk_Module_MName%>是否显示。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right"><%=Fk_Module_MName%>内容&nbsp;&nbsp;<span class="qbox" title="<p><%=Fk_Module_MName%>内容，请输入10个字符以上（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span>：</td>
        <td colspan="3"><textarea name="Fk_Product_Content" style="width:100%;" class="<%=EditorClass%>" rows="15" id="Fk_Product_Content"><%=Fk_Product_Content%></textarea></td>
    </tr>
<%
	Call FKAdmin.ShowField(1,0," And (Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Product_Module&",%%')",Fk_Product_Field,EditorClass)
%>
</table>
<!--专题/推荐-->
<table width="95%" border="1" class="tnr" id="t2" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">推荐：</td>
        <td>&nbsp;<select name="Fk_Product_Recommend" size="10" multiple="multiple" class="TextArea" id="Fk_Product_Recommend">
            <option value="0">无推荐</option>
<%
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Recommend_Id")%>"<%If Instr(Fk_Product_Recommend,","&Rs("Fk_Recommend_Id")&",")>0 Then%> selected="selected"<%End If%>><%=Rs("Fk_Recommend_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择<%=Fk_Module_MName%>推荐类型，可按住CTRL键用鼠标左键多选。</p>"><img src="Images/help.jpg" /></span></td>
        <td align="right">专题：</td>
        <td>&nbsp;<select name="Fk_Product_Subject" class="TextArea" size="10" multiple="multiple" id="Fk_Product_Subject">
            <option value="0">无专题</option>
<%
	Sqlstr="Select Fk_Subject_Id,Fk_Subject_Name From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Subject_Id")%>"<%If Instr(Fk_Product_Subject,","&Rs("Fk_Subject_Id")&",")>0 Then%> selected="selected"<%End If%>><%=Rs("Fk_Subject_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择<%=Fk_Module_MName%>归属专题，可按住CTRL键用鼠标左键多选。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
</table>
<!--其他选项-->
<table width="95%" border="1" class="tnr" id="t3" style="display:none;" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">添加时间：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Time" value="<%=Fk_Product_Time%>" type="text" class="Input" id="Fk_Product_Time" size="50" onClick="javascript:ShowCalendar(this.id)" />&nbsp;&nbsp;<span class="qbox" title="<p>添加时间，可自主修改，但必须按格式书写。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">转向链接：</td>
        <td colspan="3">&nbsp;<input name="Fk_Product_Url" value="<%=Fk_Product_Url%>" type="text" class="Input" id="Fk_Product_Url" size="50" />&nbsp;*正常<%=Fk_Module_MName%>请留空&nbsp;&nbsp;<span class="qbox" title="<p>转向链接，如果在这里输入网址，则列表上自动指向该网址。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:880px;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Product_Module%>" />
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('ProductEdit','Product.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ProductEditDo
'作    用：执行修改产品
'参    数：
'==============================
Sub ProductEditDo()
	Dim TempModuleId,TempMUrl,TempLowTemplate,TempFileName
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("您无此权限！",2)
	End If
	Sqlstr="Select Fk_Module_MName From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		If Rs("Fk_Module_MName")<>"" Then
			Fk_Module_MName=Rs("Fk_Module_MName")
		End If
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",2)
	End If
	Rs.Close
	Fk_Product_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Title")))
	Fk_Product_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Color")))
	Fk_Product_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Keyword")))
	Fk_Product_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Description")))
	Fk_Product_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Url")))
	Fk_Product_Content=Request.Form("Fk_Product_Content")
	Fk_Product_PicList=FKFun.HTMLEncode(Trim(Request.Form("Fk_PicList")))
	Fk_Product_PicList=Replace(Fk_Product_PicList,", ","||")
	Fk_Product_PicList=Replace(Fk_Product_PicList,"|||-_-|","|-_-|")
	Fk_Product_PicList=Replace(Fk_Product_PicList,"|-_-|||","|-_-|")
	If Right(Fk_Product_PicList,5)="|-_-|" Then
		Fk_Product_PicList=Left(Fk_Product_PicList,Len(Fk_Product_PicList)-5)
	End If
	Fk_Product_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Pic")))
	Fk_Product_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_PicBig")))
	Fk_Product_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_FileName")))
	Fk_Product_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Recommend"))," ",""))&","
	Fk_Product_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Subject"))," ",""))&","
	Fk_Product_Template=Trim(Request.Form("Fk_Product_Template"))
	Fk_Product_Show=Trim(Request.Form("Fk_Product_Show"))
	Fk_Product_Time=Trim(Request.Form("Fk_Product_Time"))
	Id=Trim(Request.Form("Id"))
	If Fk_Product_PicList<>"" And (Fk_Product_Pic="" Or Fk_Product_PicBig="") Then
		Call FKFun.ShowErr("请在上传的题图列表中选择封面图片！",2)
	End If
	Call FKFun.ShowString(Fk_Product_Title,1,255,0,"请输入"&Fk_Module_MName&"名称！",""&Fk_Module_MName&"名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Keyword,0,255,2,"请输入"&Fk_Module_MName&"关键字！",""&Fk_Module_MName&"关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Description,0,255,2,"请输入"&Fk_Module_MName&"描述！",""&Fk_Module_MName&"描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Url,0,255,2,"请输入"&Fk_Module_MName&"转向链接！",""&Fk_Module_MName&"转向链接不能大于255个字符！")
	If Fk_Product_Url="" Then
		Call FKFun.ShowString(Fk_Product_Content,10,1,1,"请输入"&Fk_Module_MName&"内容，不少于10个字符！",""&Fk_Module_MName&"内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Product_Pic,0,255,2,"请输入"&Fk_Module_MName&"题图路径！",""&Fk_Module_MName&"题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_PicBig,0,255,2,"请输入"&Fk_Module_MName&"题图路径！",""&Fk_Module_MName&"题图大图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_FileName,0,50,2,"请输入"&Fk_Module_MName&"文件名！",""&Fk_Module_MName&"文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Product_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Product_Show,"请选择"&Fk_Module_MName&"是否显示！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	If Fk_Product_Time="" Then
		Fk_Product_Time=Now()
	End If
	Fk_Product_Field=FKAdmin.GetFieldData(0,"(Fk_Field_Content Like '%%,Product,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')")
	Call FKAdmin.CheckDandF("file",Fk_Product_FileName,0)
	If Fk_Site_DelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Product_Content=Replace(Fk_Product_Content,Temp,"**")
				Fk_Product_Title=Replace(Fk_Product_Title,Temp,"**")
				Fk_Product_Keyword=Replace(Fk_Product_Keyword,Temp,"**")
				Fk_Product_Description=Replace(Fk_Product_Description,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select Fk_Product_Id,Fk_Product_Title,Fk_Product_Color,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Content,Fk_Product_MUrl,Fk_Product_PicList,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_Module,Fk_Product_Click,Fk_Product_FileName,Fk_Product_Recommend,Fk_Product_Subject,Fk_Product_Field,Fk_Product_Template,Fk_Product_Show,Fk_Product_Url,Fk_Product_Time From [Fk_Product] Where Fk_Product_Id="&Id&""
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Product_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		TempModuleId=Rs("Fk_Product_Module")
		If Fk_Product_FileName<>"" Then
			TempFileName=Fk_Product_FileName
		Else
			TempFileName=Rs("Fk_Product_FileName")
		End If
		Application.Lock()
		Rs("Fk_Product_Title")=Fk_Product_Title
		Rs("Fk_Product_Color")=Fk_Product_Color
		Rs("Fk_Product_Keyword")=Fk_Product_Keyword
		Rs("Fk_Product_Field")=Fk_Product_Field
		Rs("Fk_Product_Description")=Fk_Product_Description
		Rs("Fk_Product_Url")=Fk_Product_Url
		Rs("Fk_Product_Show")=Fk_Product_Show
		Rs("Fk_Product_Pic")=Fk_Product_Pic
		Rs("Fk_Product_PicBig")=Fk_Product_PicBig
		Rs("Fk_Product_PicList")=Fk_Product_PicList
		Rs("Fk_Product_Content")=Fk_Product_Content
		Rs("Fk_Product_Recommend")=Fk_Product_Recommend
		Rs("Fk_Product_Subject")=Fk_Product_Subject
		If Rs("Fk_Product_FileName")="" Or IsNull(Rs("Fk_Product_FileName")) Then
			Rs("Fk_Product_FileName")=Fk_Product_FileName
		End If
		Rs("Fk_Product_Template")=Fk_Product_Template
		Rs("Fk_Product_Time")=Fk_Product_Time
		Rs.Update()
		Application.UnLock()
		Response.Write(""&Fk_Module_MName&"修改成功！")
		Rs.Close
		If Fk_Site_Html=2 And Fk_Product_Show=1 Then
			Sqlstr="Select Fk_Module_Id,Fk_Module_Type,Fk_Module_LowTemplate,Fk_Module_MUrl From [Fk_Module] Where Fk_Module_Id="&TempModuleId&" And Fk_Module_Show=1"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
				TempLowTemplate=Rs("Fk_Module_LowTemplate")
				Rs.Close
				If TempFileName<>"" Then
					TempMUrl=TempMUrl&TempFileName&FKTemplate.GetHtmlSuffix()
				Else
					TempMUrl=TempMUrl&Id&FKTemplate.GetHtmlSuffix()
				End If
				Response.Write("<span style='display:none;'>")
				Call FKHtml.CreatPage(Id,TempModuleId,2,TempMUrl,Fk_Product_Title)
				Response.Write("</span>")
			Else
				Rs.Close
			End If
		End If
	Else
		Rs.Close
		Response.Write(""&Fk_Module_MName&"不存在！")
	End If
End Sub

'==============================
'函 数 名：ProductDelDo
'作    用：执行删除产品
'参    数：
'==============================
Sub ProductDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Product_Id,Fk_Product_Module From [Fk_Product] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Product_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("删除成功！")
	Else
		Rs.Close
		Call FKFun.ShowErr("不存在！",2)
	End If
End Sub

'==============================
'函 数 名：ListDelDo
'作    用：执行批量删除产品
'参    数：
'==============================
Sub ListDelDo()
	Id=Replace(Trim(Request.Form("ListId"))," ","")
	If Id="" Then
		Call FKFun.ShowErr("请选择要删除的内容！",2)
	End If
	Sqlstr="Select Top 1 Fk_Product_Module From [Fk_Product] Where Fk_Product_Id In ("&Id&")"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Product_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
	End If
	Sqlstr="Delete From [Fk_Product] Where Fk_Product_Id In ("&Id&")"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("批量删除成功！")
End Sub

'==============================
'函 数 名：ProductMove
'作    用：执行批量移动产品
'参    数：
'==============================
Sub ProductMove()
	Id=Replace(Trim(Request.QueryString("ListId"))," ","")
	Fk_Module_Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Fk_Module_Id,"请选择转移到的模块！")
	If Id="" Then
		Call FKFun.ShowErr("请选择要移动的内容！",2)
	End If
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Rs.Close
		Call FKFun.ShowErr("转移到的模块没有权限！",2)
	End If
	Sqlstr="Select Fk_Module_Type From [Fk_Module] Where Fk_Module_Id="&Fk_Module_Id&""
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Type=Rs("Fk_Module_Type")
	Else
		Rs.Close
		Call FKFun.ShowErr("要移到的模块不存在！",2)
	End If
	Rs.Close
	If Fk_Module_Type<>2 Then
		Call FKFun.ShowErr("只能移动到同类型的模块！",2)
	End If
	Sqlstr="Update [Fk_Product] Set Fk_Product_Module="&Fk_Module_Id&" Where Fk_Product_Id In ("&Id&")"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("批量移动成功！")
End Sub

'==============================
'函 数 名：ProductOrderSet
'作    用：执行产品排序
'参    数：
'==============================
Sub ProductOrderSet()
	Id=Trim(Request.QueryString("Id"))
	Fk_Product_Order=Trim(Request.QueryString("Order"))
	Call FKFun.ShowNum(Id,"参数错误！")
	Call FKFun.ShowNum(Fk_Product_Order,"排序序号必须是数字！")
	Sqlstr="Update [Fk_Product] Set Fk_Product_Order="&Fk_Product_Order&" Where Fk_Product_Id="&Id&""
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("排序序号设置成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->