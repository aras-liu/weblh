<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Module.asp
'文件用途：模块管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System1",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim MenuId
Dim Fk_Menu_Name,Fk_Menu_Dir,Fk_Menu_IsIndex
Dim Fk_Module_Name,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Type,Fk_Module_Dir,Fk_Module_MUrl,Fk_Module_PageCount,Fk_Module_Menu,Fk_Module_Level,Fk_Module_Order,Fk_Module_LevelList,Fk_Module_Template,Fk_Module_LowTemplate,Fk_Module_UrlType,Fk_Module_Url,Fk_Module_Show,Fk_Module_MenuShow,Fk_Module_GModel,Fk_Module_MName,Fk_Module_Pic,Fk_Module_Subhead,Fk_Module_IsIndex,Fk_Module_Limit,Fk_Module_GUrl

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ModuleList() '模块列表
	Case 2
		Call ModuleAddForm() '添加模块表单
	Case 3
		Call ModuleAddDo() '执行添加模块
	Case 4
		Call ModuleEditForm() '修改模块表单
	Case 5
		Call ModuleEditDo() '执行修改模块
	Case 6
		Call ModuleDelDo() '执行删除模块
	Case 7
		Call ModuleOrderForm() '模块排序表单
	Case 8
		Call ModuleOrderDo() '执行模块排序
	Case 9
		Call ReUrlDo() '重置模块链接
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：ModuleList()
'作    用：模块列表
'参    数：
'==========================================
Sub ModuleList()
	Session("NowPage")=FkFun.GetNowUrl()
	MenuId=Clng(Request.QueryString("MenuId"))
	Sqlstr="Select Fk_Menu_Name From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	Else
		Call FKFun.ShowErr("菜单不存在！",2)
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Module.asp?Type=2&MenuId=<%=MenuId%>');">添加新模块</a></li>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Module.asp?Type=7&MenuId=<%=MenuId%>');">模块排序</a></li>
        <li><a href="javascript:void(0);" onclick="DelIt('需要重置“<%=Fk_Menu_Name%>”菜单下模块链接么？','Module.asp?Type=9&MenuId=<%=MenuId%>','MainRight','<%=Session("NowPage")%>');">重置模块链接</a></li>
    </ul>
</div>
<div id="ListTop">
    “<%=Fk_Menu_Name%>”菜单模块管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">模块名称</td>
            <td align="center" class="ListTdTop">前台链接</td>
            <td align="center" class="ListTdTop">模块类型</td>
            <td align="center" class="ListTdTop">文件名/目录</td>
            <td align="center" class="ListTdTop">模块模板</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Call FKAdmin.GetModuleList(4,MenuId,0,0,"")
%>
        <tr>
            <td height="30" colspan="8">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：ModuleAddForm()
'作    用：添加模块表单
'参    数：
'==========================================
Sub ModuleAddForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Sqlstr="Select Fk_Menu_Name,Fk_Menu_Dir From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
		Fk_Menu_Dir=Rs("Fk_Menu_Dir")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到菜单，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
	If Fk_Menu_Dir<>"" Then
		Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_IsIndex=1 And Fk_Module_Menu=" & MenuId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Fk_Menu_IsIndex=1
		Else
			Fk_Menu_IsIndex=0
		End If
		Rs.Close
	Else
		Fk_Menu_IsIndex=0
	End If
%>
<form id="ModuleAdd" name="ModuleAdd" method="post" action="Module.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">添加新“<%=Fk_Menu_Name%>”模块[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<ul class="BoxNav">
    	<li class="check bnr" id="s1" onclick="ClickBoxNav('1');">常规选项</li>
    	<li class="bnr" id="s2" onclick="ClickBoxNav('2');">权限分配</li>
        <div class="Cal"></div>
    </ul>
    <div class="Cal"></div>
<!--常规标签列表-->
<table width="90%" class="tnr" id="t1" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">模块名称：</td>
        <td>&nbsp;<input name="Fk_Module_Name" type="text" class="Input" id="Fk_Module_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>模块名称，不能为空，同一级中不能有重复，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">副标题：</td>
        <td>&nbsp;<input name="Fk_Module_Subhead" type="text" class="Input" id="Fk_Module_Subhead" />&nbsp;&nbsp;<span class="qbox" title="<p>副标题，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7">
        <td height="28" align="right">关键字：</td>
        <td>&nbsp;<input name="Fk_Module_Keyword" type="text" class="Input" id="Fk_Module_Keyword" size="40" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于页面meta的keywords，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7">
        <td height="28" align="right">描述：</td>
        <td>&nbsp;<input name="Fk_Module_Description" type="text" class="Input" id="Fk_Module_Description" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入站点的一段文字简介，用于页面meta的description，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块类型：</td>
        <td>&nbsp;<select name="Fk_Module_Type" class="Input" id="Fk_Module_Type" onchange="ModuleTypeChange(this.options[this.options.selectedIndex].value);">
<%
	For i=0 To UBound(FKModuleId)
%>
                <option value="<%=FKModuleId(i)%>"><%=FKModuleName(i)%></option>
<%
	Next
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>模块类型，不同的模块有不同的用途。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块目录/文件名：</td>
        <td>&nbsp;<input name="Fk_Module_Dir" type="text" class="Input" id="Fk_Module_Dir" />*一旦确立不可修改&nbsp;&nbsp;<span class="qbox" title="<p>本选项在列表栏目中，此设置为目录，如果在单页类型模块，则为文件名。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module5" style="display:none;">
        <td height="28" align="right">跳转模式：</td>
        <td>&nbsp;<input name="Fk_Module_UrlType" type="radio" class="Input" id="Fk_Module_UrlType" value="0" checked="checked" />直接跳转
        <input type="radio" name="Fk_Module_UrlType" class="Input" id="Fk_Module_UrlType" value="1" />页面跳转&nbsp;&nbsp;<span class="qbox" title="<p>设置跳转链接的方式，如果选择“直接跳转”，则会直接输出链接到页面，可以做外链；如选择“页面跳转”，会生成一个页面，并在通过页面内代码跳转。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr  class="moduleList module5" style="display:none;">
        <td height="28" align="right">模块转向链接：</td>
        <td>&nbsp;<input name="Fk_Module_Url" type="text" class="Input" id="Fk_Module_Url" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>设置一个链接，链接请输入1-255个字符，支持相对链接。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块分级：</td>
        <td>&nbsp;<select name="Fk_Module_Level" class="Input" id="Fk_Module_Level">
            <option value="0">一级模块</option>
<%
	Call FKAdmin.GetModuleList(6,MenuId,0,0,"")
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择模块所属分级。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module4 module7" style="display:none;">
        <td height="28" align="right">每页条数：</td>
        <td>&nbsp;<select name="Fk_Module_PageCount" class="Input" id="Fk_Module_PageCount">
            <option value="0">系统默认</option>
<%
	For i=1 To 50
%>
            <option value="<%=i%>"><%=i%>条</option>
<%
	Next
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>列表每页条数，如果选择系统默认，则使用站点设置中的“每页条数”设置值。</p>"><img src="Images/help.jpg" /></span></td>
       </tr>
    <tr class="moduleList module4" style="display:none;">
        <td height="28" align="right">留言后跳转到：</td>
        <td>&nbsp;<input name="Fk_Module_GUrl" type="text" class="Input" id="Fk_Module_GUrl" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功后跳转到的页面链接，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module4" style="display:none;">
        <td height="28" align="right">留言模型：</td>
        <td>&nbsp;<select name="Fk_Module_GModel" class="Input" id="Fk_Module_GModel">
<%
	Sqlstr="Select Fk_GModel_Id,Fk_GModel_Name From [Fk_GModel] Order By Fk_GModel_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_GModel_Id")%>"><%=Rs("Fk_GModel_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>请选择留言模型。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7">
        <td height="28" align="right">显示模板：</td>
        <td>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
            <option value="0">默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本模块页面使用的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module4 module7" style="display:none;">
        <td height="28" align="right">子内容模板：</td>
        <td>&nbsp;<select name="Fk_Module_LowTemplate" class="Input" id="Fk_Module_LowTemplate">
            <option value="0">默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本模块子内容页面的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module7" style="display:none;">
        <td height="28" align="right">管理名称：</td>
        <td>&nbsp;<input name="Fk_Module_MName" type="text" class="Input" id="Fk_Module_MName" />&nbsp;&nbsp;<span class="qbox" title="<p>管理名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="25" align="right">模块图片：</td>
        <td>&nbsp;<input name="Fk_Module_Pic" type="text" class="Input" id="Fk_Module_Pic" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>模块图片，可留空，请输入0-255个字符，可以输入链接，也可以上传到空间。</p>"><img src="Images/help.jpg" /></span><br />
    &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Module_Pics" name="Fk_Module_Pics" src="PicUpLoad.asp?Type=2&Form=ModuleAdd&Input=Fk_Module_Pic"></iframe>
        </td>
    </tr>
    <tr>
        <td height="28" align="right">是否使用：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_Show" class="Input" id="Fk_Module_Show" value="0" />禁用
        <input name="Fk_Module_Show" type="radio" class="Input" id="Fk_Module_Show" value="1" checked="checked" />使用&nbsp;&nbsp;<span class="qbox" title="<p>如果选择禁用，则前台不可访问，生成亦不会涉及本模块，但本设置对已经生成的文件没有作用，如果已经生成页面，请手工进入FTP删除。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">是否菜单显示：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_MenuShow" class="Input" id="Fk_Module_MenuShow" value="0" />不显示
        <input name="Fk_Module_MenuShow" type="radio" class="Input" id="Fk_Module_MenuShow" value="1" checked="checked" />显示&nbsp;&nbsp;<span class="qbox" title="<p>本设置仅涉及菜单显示，前台访问、生成等一切照常，只是不在菜单中输出。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	If Fk_Menu_IsIndex=0 Then
%>
    <tr class="moduleList module0">
        <td height="28" align="right">菜单首页：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_IsIndex" class="Input" id="Fk_Module_IsIndex" value="0" checked="checked" />非首页
        <input name="Fk_Module_IsIndex" type="radio" class="Input" id="Fk_Module_IsIndex" value="1" />首页&nbsp;&nbsp;<span class="qbox" title="<p>当菜单有菜单目录时，可以设置一个静态模块为菜单首页，菜单首页链接为“菜单目录/index.html”，归入一键更新的自动首页生成。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	End If
%>
</table>
<!--权限分配-->
<table width="90%" class="tnr" id="t2" style="display:none;" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right" width="20%">权限分配<span class="qbox" title="<p>可以在此设置新模块可以管理内容的普通人员。</p>"><img src="Images/help.jpg" /></span>：</td>
        <td>
<%
	Sqlstr="Select Fk_Admin_Id,Fk_Admin_LoginName,Fk_Admin_Name From [Fk_Admin] Where Fk_Admin_Limit>0 Order By Fk_Admin_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
        &nbsp;<input type="checkbox" name="Fk_Module_Limit" class="Checks" value="<%=Rs("Fk_Admin_Id")%>" id="Fk_Module_Limit" /> <%=Rs("Fk_Admin_Name")%>[<%=Rs("Fk_Admin_LoginName")%>]<br />
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:580px;">
		<input type="hidden" name="MenuId" value="<%=MenuId%>" />
        <input type="submit" onclick="Sends('ModuleAdd','Module.asp?Type=3',0,'',0,1,'MainRight','Module.asp?Type=1&MenuId=<%=MenuId%>');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ModuleAddDo
'作    用：执行添加模块
'参    数：
'==============================
Sub ModuleAddDo()
	Dim TempArr2,Temp2
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Subhead=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Subhead")))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
	Fk_Module_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Url")))
	Fk_Module_MName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_MName")))
	Fk_Module_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Pic")))
	Fk_Module_GUrl=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_GUrl")))
	Fk_Module_Type=Trim(Request.Form("Fk_Module_Type"))
	Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
	Fk_Module_Menu=Trim(Request.Form("MenuId"))
	Fk_Module_Level=Trim(Request.Form("Fk_Module_Level"))
	Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
	Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
	Fk_Module_UrlType=Trim(Request.Form("Fk_Module_UrlType"))
	Fk_Module_Show=Trim(Request.Form("Fk_Module_Show"))
	Fk_Module_MenuShow=Trim(Request.Form("Fk_Module_MenuShow"))
	Fk_Module_GModel=Trim(Request.Form("Fk_Module_GModel"))
	Fk_Module_IsIndex=Trim(Request.Form("Fk_Module_IsIndex"))
	Fk_Module_Limit=FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Module_Limit"))," ",""))
	Call FKFun.ShowString(Fk_Module_Name,1,50,0,"请输入模块名称！","模块名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Subhead,0,50,0,"请输入模块副标题！","副标题不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入关键字！","关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入描述！","描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入模块目录名/文件名！","模块目录名/文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Module_Type,"请选择模块类型！")
	Call FKFun.ShowNum(Fk_Module_Level,"请选择模块分级！")
	Call FKFun.ShowNum(Fk_Module_Show,"请选择模块是否菜单显示！")
	Call FKFun.ShowNum(Fk_Module_Template,"请选择模块模板！")
	Call FKFun.ShowNum(Fk_Module_Menu,"系统参数错误，请刷新页面！")
	If Fk_Module_Type=0 Then
		If Fk_Module_IsIndex<>"" Then
			Call FKFun.ShowNum(Fk_Module_IsIndex,"请选择是否模块首页！")
		Else
			Fk_Module_IsIndex=0
		End If
		' ===== 此处代码移动到下方 ===== 
	Else
		Fk_Module_IsIndex=0
	End If

	' ===== Fk_Menu_Dir的代码先移动出来 ===== 
	Sqlstr="Select Fk_Menu_Dir From [Fk_Menu] Where Fk_Menu_Id=" & Fk_Module_Menu
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Dir=Rs("Fk_Menu_Dir")
	End If
	Rs.Close
	' ===== End ===== 

	
	Select Case Fk_Module_Type
		Case 0
		Case 1
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
		Case 2
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
		Case 3
		Case 4
			Call FKFun.ShowString(Fk_Module_GUrl,0,255,0,"请输入留言后跳转链接！","留言后跳转链接不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_GModel,"请选择留言模型！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
		Case 5
			Call FKFun.ShowString(Fk_Module_Url,1,255,0,"请输入转向链接！","转向链接不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_UrlType,"请选择跳转模式！")
		Case 6
		Case 7
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
	End Select
	If Fk_Module_Level>0 Then
		Fk_Module_LevelList=FKAdmin.GetModuleLevelList(Fk_Module_Level)
	Else
		Fk_Module_LevelList=""
	End If
	If Fk_Module_Dir<>"" Then
		Call FKAdmin.CheckDandF("dir",Fk_Module_Dir,Fk_Module_Level)
		If Fk_Site_HtmlType=0 Then
			' ===== 增加Fk_Module_Menu判断，即同一menu下目录名称不能相同 ===== 
			Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"' AND Fk_Module_Menu="&Fk_Module_Menu&""
		ElseIf Fk_Site_HtmlType=1 Then
			' ===== 增加Fk_Module_Menu判断，即同一menu下目录名称不能相同 ===== 
			Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"' And Fk_Module_Level="&Fk_Module_Level&" AND Fk_Module_Menu="&Fk_Module_Menu&""
		End If
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Rs.Close
			Call FKFun.ShowErr("模块文件名/目录已经被使用，请重新输入！",2)
		End If
		Rs.Close
	End If
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Type,Fk_Module_Dir,Fk_Module_MUrl,Fk_Module_PageCount,Fk_Module_Menu,Fk_Module_Level,Fk_Module_LevelList,Fk_Module_Template,Fk_Module_LowTemplate,Fk_Module_UrlType,Fk_Module_Url,Fk_Module_Show,Fk_Module_MenuShow,Fk_Module_MName,Fk_Module_Admin,Fk_Module_Ip,Fk_Module_GModel,Fk_Module_Pic,Fk_Module_Subhead,Fk_Module_IsIndex,Fk_Module_GUrl From [Fk_Module] Where Fk_Module_Name='"&Fk_Module_Name&"' And Fk_Module_Level="&Fk_Module_Level&" AND Fk_Module_Menu="&Fk_Module_Menu&""
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Subhead")=Fk_Module_Subhead
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		Rs("Fk_Module_Dir")=Fk_Module_Dir
		Rs("Fk_Module_Type")=Fk_Module_Type
		Rs("Fk_Module_Menu")=Fk_Module_Menu
		Rs("Fk_Module_Level")=Fk_Module_Level
		Rs("Fk_Module_LevelList")=Fk_Module_LevelList
		Rs("Fk_Module_Template")=Fk_Module_Template
		Rs("Fk_Module_Show")=Fk_Module_Show
		Rs("Fk_Module_MenuShow")=Fk_Module_MenuShow
		Rs("Fk_Module_MUrl")=""
		Rs("Fk_Module_Pic")=Fk_Module_Pic
		Rs("Fk_Module_IsIndex")=Fk_Module_IsIndex
		Select Case Fk_Module_Type
			Case 0
			Case 1
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
			Case 2
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
			Case 3
			Case 4
				Rs("Fk_Module_GUrl")=Fk_Module_GUrl
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_GModel")=Fk_Module_GModel
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 5
				Rs("Fk_Module_Url")=Fk_Module_Url
				Rs("Fk_Module_UrlType")=Fk_Module_UrlType
			Case 6
			Case 7
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
		End Select
		Rs("Fk_Module_Admin")=Request.Cookies("FkAdminId")
		Rs("Fk_Module_Ip")=Request.ServerVariables("REMOTE_ADDR")
		Rs.Update()
		Dim newid
		newid = Rs("Fk_Module_Id")
		Rs.Close

		' ===== 加上Fk_Module_Menu菜单判断 =====
		Sqlstr="Select Fk_Module_Id,Fk_Module_MUrl,Fk_Module_IsIndex,Fk_Module_Type,Fk_Module_LevelList,Fk_Module_Dir,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Name='"&Fk_Module_Name&"' And Fk_Module_Level="&Fk_Module_Level&" AND Fk_Module_Menu="&Fk_Module_Menu&""
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Id=Rs("Fk_Module_Id")
			If Fk_Module_IsIndex=0 Then
				Rs("Fk_Module_MUrl")=FKAdmin.ModuleMUrl(Rs("Fk_Module_Id"),Rs("Fk_Module_Type"),Rs("Fk_Module_LevelList"),Rs("Fk_Module_UrlType"),Rs("Fk_Module_Url"),Rs("Fk_Module_Dir"),Fk_Menu_Dir,Fk_Site_HtmlType,Fk_Site_HtmlSuffix)
			Else
				Rs("Fk_Module_MUrl")=Fk_Menu_Dir&"/"
			End If
			Rs.Update()
		End If
		Application.UnLock()
		Response.Write("新模块添加成功！")
	Else
		Response.Write("该模块标题已经被占用，请重新输入！")
	End If
	Rs.Close
	Temp=""
	If Fk_Module_Limit<>"" Then
		Sqlstr="Select Fk_Admin_Id,Fk_Admin_Limit From [Fk_Admin] Where Fk_Admin_Id In ("&Fk_Module_Limit&") Order By Fk_Admin_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
			If Temp="" Then
				Temp=Rs("Fk_Admin_Limit")
			Else
				Temp=Temp&","&Rs("Fk_Admin_Limit")
			End If
			Rs.MoveNext
		Wend
		Rs.Close
		If Temp<>"" Then
			Fk_Module_Limit="Menu"&Fk_Module_Menu&",Module"&Id
			If Fk_Module_LevelList<>"" Then
				TempArr2=Split(Fk_Module_LevelList,",")
				For Each Temp2 In TempArr2
					If Temp2<>"" Then
						Fk_Module_Limit=Fk_Module_Limit&","&"See"&Temp2
					End If
				Next
			End If
			Sqlstr="Select Fk_Limit_Id,Fk_Limit_Content From [Fk_Limit] Order By Fk_Limit_Id Asc"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
				If Rs("Fk_Limit_Content")<>"" Then
					TempArr2=Split(Fk_Module_Limit,",")
					For Each Temp2 In TempArr2
						If Instr(Rs("Fk_Limit_Content"),","&Temp2&",")=0 Then
							Rs("Fk_Limit_Content")=Rs("Fk_Limit_Content")&Temp2&","
						End If
					Next
				Else
					Rs("Fk_Limit_Content")=Fk_Module_Limit
				End If
				Rs.Update
				Rs.MoveNext
			Wend
			Rs.Close
		End If
	End If
End Sub

'==========================================
'函 数 名：ModuleEditForm()
'作    用：修改模块表单
'参    数：
'==========================================
Sub ModuleEditForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Module_Name,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Type,Fk_Module_Dir,Fk_Module_PageCount,Fk_Module_Menu,Fk_Module_Level,Fk_Module_Template,Fk_Module_LowTemplate,Fk_Module_UrlType,Fk_Module_Url,Fk_Module_Show,Fk_Module_MenuShow,Fk_Module_GModel,Fk_Module_MName,Fk_Module_Pic,Fk_Module_Subhead,Fk_Module_IsIndex,Fk_Module_GUrl From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Subhead=Rs("Fk_Module_Subhead")
		Fk_Module_Keyword=Rs("Fk_Module_Keyword")
		Fk_Module_Description=Rs("Fk_Module_Description")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Module_Url=Rs("Fk_Module_Url")
		Fk_Module_Type=Rs("Fk_Module_Type")
		Fk_Module_PageCount=Rs("Fk_Module_PageCount")
		Fk_Module_Level=Rs("Fk_Module_Level")
		Fk_Module_Template=Rs("Fk_Module_Template")
		Fk_Module_LowTemplate=Rs("Fk_Module_LowTemplate")
		Fk_Module_UrlType=Rs("Fk_Module_UrlType")
		Fk_Module_Show=Rs("Fk_Module_Show")
		Fk_Module_MenuShow=Rs("Fk_Module_MenuShow")
		Fk_Module_GModel=Rs("Fk_Module_GModel")
		Fk_Module_MName=Rs("Fk_Module_MName")
		Fk_Module_Pic=Rs("Fk_Module_Pic")
		Fk_Module_IsIndex=Rs("Fk_Module_IsIndex")
		Fk_Module_GUrl=Rs("Fk_Module_GUrl")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
	Sqlstr="Select Fk_Menu_Dir From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Dir=Rs("Fk_Menu_Dir")
	End If
	Rs.Close
	If Fk_Menu_Dir<>"" Then
		Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_IsIndex=1 And Fk_Module_Menu=" & MenuId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Fk_Menu_IsIndex=1
		Else
			Fk_Menu_IsIndex=0
		End If
		Rs.Close
	Else
		Fk_Menu_IsIndex=0
	End If
%>
<form id="ModuleEdit" name="ModuleEdit" method="post" action="Module.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">修改“<%=Fk_Module_Name%>”模块[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">模块名称：</td>
        <td>&nbsp;<input name="Fk_Module_Name" type="text" class="Input" id="Fk_Module_Name" value="<%=Fk_Module_Name%>" />&nbsp;&nbsp;<span class="qbox" title="<p>模块名称，不能为空，同一级中不能有重复，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">副标题：</td>
        <td>&nbsp;<input name="Fk_Module_Subhead" type="text" class="Input" id="Fk_Module_Subhead" value="<%=Fk_Module_Subhead%>" />&nbsp;&nbsp;<span class="qbox" title="<p>副标题，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7"<%If Instr(",5,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">关键字：</td>
        <td>&nbsp;<input name="Fk_Module_Keyword" type="text" class="Input" id="Fk_Module_Keyword" value="<%=Fk_Module_Keyword%>" size="40" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于页面meta的keywords，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7"<%If Instr(",5,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">描述：</td>
        <td>&nbsp;<input name="Fk_Module_Description" type="text" class="Input" id="Fk_Module_Description" value="<%=Fk_Module_Description%>" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入站点的一段文字简介，用于页面meta的description，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块类型：</td>
        <td>&nbsp;<select name="Fk_Module_Type" class="Input" id="Fk_Module_Type" onchange="ModuleTypeChange(this.options[this.options.selectedIndex].value);">
<%
	For i=0 To UBound(FKModuleId)
%>
                <option value="<%=FKModuleId(i)%>"<%=FKFun.BeSelect(Fk_Module_Type,Clng(FKModuleId(i)))%>><%=FKModuleName(i)%></option>
<%
	Next
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>模块类型，不同的模块有不同的用途。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块目录/文件名：</td>
        <td>&nbsp;<input name="Fk_Module_Dir" type="text" class="Input" id="Fk_Module_Dir" value="<%=Fk_Module_Dir%>" <%If Fk_Module_Dir<>"" Then%> disabled="disabled"<%End If%> />*一旦确立不可修改&nbsp;&nbsp;<span class="qbox" title="<p>本选项在列表栏目中，此设置为目录，如果在单页类型模块，则为文件名。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module5"<%If Instr(",0,1,2,3,4,6,7,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">跳转模式：</td>
        <td>&nbsp;<input name="Fk_Module_UrlType" type="radio" class="Input" id="Fk_Module_UrlType" value="0"<%=FKFun.BeCheck(Fk_Module_UrlType,0)%> />直接跳转
        <input type="radio" name="Fk_Module_UrlType" class="Input" id="Fk_Module_UrlType" value="1"<%=FKFun.BeCheck(Fk_Module_UrlType,1)%> />页面跳转&nbsp;&nbsp;<span class="qbox" title="<p>设置跳转链接的方式，如果选择“直接跳转”，则会直接输出链接到页面，可以做外链；如选择“页面跳转”，会生成一个页面，并在通过页面内代码跳转。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr  class="moduleList module5"<%If Instr(",0,1,2,3,4,6,7,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">模块转向链接：</td>
        <td>&nbsp;<input name="Fk_Module_Url" type="text" class="Input" id="Fk_Module_Url" value="<%=Fk_Module_Url%>" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>设置一个链接，链接请输入1-255个字符，支持相对链接。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">模块分级：</td>
        <td>&nbsp;<select name="Fk_Module_Level" class="Input" id="Fk_Module_Level">
            <option value="0">一级模块</option>
<%
	Call FKAdmin.GetModuleList(6,MenuId,0,Fk_Module_Level,"")
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择模块所属分级。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module4 module7"<%If Instr(",0,3,5,6,7,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">每页条数：</td>
        <td>&nbsp;<select name="Fk_Module_PageCount" class="Input" id="Fk_Module_PageCount">
            <option value="0">系统默认</option>
<%
	For i=1 To 50
%>
            <option value="<%=i%>"<%=FKFun.BeSelect(Fk_Module_PageCount,i)%>><%=i%>条</option>
<%
	Next
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>列表每页条数，如果选择系统默认，则使用站点设置中的“每页条数”设置值。</p>"><img src="Images/help.jpg" /></span></td>
       </tr>
    <tr class="moduleList module4"<%If Instr(",0,1,2,3,5,6,7,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">留言后跳转到：</td>
        <td>&nbsp;<input name="Fk_Module_GUrl" value="<%=Fk_Module_GUrl%>" type="text" class="Input" id="Fk_Module_GUrl" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功后跳转到的页面链接，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module4"<%If Instr(",0,1,2,3,5,6,7,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">留言模型：</td>
        <td>&nbsp;<select name="Fk_Module_GModel" class="Input" id="Fk_Module_GModel">
<%
	Sqlstr="Select Fk_GModel_Id,Fk_GModel_Name From [Fk_GModel] Order By Fk_GModel_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_GModel_Id")%>"<%=FKFun.BeSelect(Fk_Module_GModel,Rs("Fk_GModel_Id"))%>><%=Rs("Fk_GModel_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>请选择留言模型。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module0 module1 module2 module3 module4 module6 module7"<%If Instr(",5,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">显示模板：</td>
        <td>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
            <option value="0">默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本模块页面使用的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module4 module7"<%If Instr(",0,3,5,6,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">子内容模板：</td>
        <td>&nbsp;<select name="Fk_Module_LowTemplate" class="Input" id="Fk_Module_LowTemplate">
            <option value="0">默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_LowTemplate,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本模块子内容页面的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr class="moduleList module1 module2 module7"<%If Instr(",0,3,4,5,6,8,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">管理名称：</td>
        <td>&nbsp;<input name="Fk_Module_MName" value="<%=Fk_Module_MName%>" type="text" class="Input" id="Fk_Module_MName" />&nbsp;&nbsp;<span class="qbox" title="<p>管理名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="25" align="right">模块图片：</td>
        <td>&nbsp;<input name="Fk_Module_Pic" value="<%=Fk_Module_Pic%>" type="text" class="Input" id="Fk_Module_Pic" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>模块图片，可留空，请输入0-255个字符，可以输入链接，也可以上传到空间。</p>"><img src="Images/help.jpg" /></span><br />
    &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Module_Pics" name="Fk_Module_Pics" src="PicUpLoad.asp?Type=2&Form=ModuleEdit&Input=Fk_Module_Pic"></iframe>
        </td>
    </tr>
    <tr>
        <td height="28" align="right">是否使用：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_Show" class="Input" id="Fk_Module_Show" value="0"<%=FKFun.BeCheck(Fk_Module_Show,0)%> />禁用
        <input name="Fk_Module_Show" type="radio" class="Input" id="Fk_Module_Show" value="1"<%=FKFun.BeCheck(Fk_Module_Show,1)%> />使用&nbsp;&nbsp;<span class="qbox" title="<p>如果选择禁用，则前台不可访问，生成亦不会涉及本模块，但本设置对已经生成的文件没有作用，如果已经生成页面，请手工进入FTP删除。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">是否菜单显示：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_MenuShow" class="Input" id="Fk_Module_MenuShow" value="0"<%=FKFun.BeCheck(Fk_Module_MenuShow,0)%> />不显示
        <input name="Fk_Module_MenuShow" type="radio" class="Input" id="Fk_Module_MenuShow" value="1"<%=FKFun.BeCheck(Fk_Module_MenuShow,1)%> />显示&nbsp;&nbsp;<span class="qbox" title="<p>本设置仅涉及菜单显示，前台访问、生成等一切照常，只是不在菜单中输出。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	If Fk_Menu_IsIndex=0 Or Fk_Module_IsIndex=1 Then
%>
    <tr class="moduleList module0"<%If Instr(",0,",Fk_Module_Type)=0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">菜单首页：</td>
        <td>&nbsp;<input type="radio" name="Fk_Module_IsIndex" class="Input" id="Fk_Module_IsIndex" value="0"<%=FKFun.BeCheck(Fk_Module_IsIndex,0)%> />非首页
        <input name="Fk_Module_IsIndex" type="radio" class="Input" id="Fk_Module_IsIndex" value="1"<%=FKFun.BeCheck(Fk_Module_IsIndex,1)%> />首页&nbsp;&nbsp;<span class="qbox" title="<p>当菜单有菜单目录时，可以设置一个静态模块为菜单首页，菜单首页链接为“菜单目录/index.html”，归入一键更新的自动首页生成。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	End If
%>
</table>
</div>
<div id="BoxBottom" style="width:580px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="MenuId" value="<%=MenuId%>" />
        <input type="submit" onclick="Sends('ModuleEdit','Module.asp?Type=5',0,'',0,1,'MainRight','Module.asp?Type=1&MenuId=<%=MenuId%>');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ModuleEditDo
'作    用：执行修改模块
'参    数：
'==============================
Sub ModuleEditDo()
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Subhead=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Subhead")))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
	Fk_Module_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Url")))
	Fk_Module_MName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_MName")))
	Fk_Module_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Pic")))
	Fk_Module_GUrl=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_GUrl")))
	Fk_Module_Type=Trim(Request.Form("Fk_Module_Type"))
	Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
	Fk_Module_Menu=Trim(Request.Form("MenuId"))
	Fk_Module_Level=Trim(Request.Form("Fk_Module_Level"))
	Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
	Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
	Fk_Module_UrlType=Trim(Request.Form("Fk_Module_UrlType"))
	Fk_Module_Show=Trim(Request.Form("Fk_Module_Show"))
	Fk_Module_MenuShow=Trim(Request.Form("Fk_Module_MenuShow"))
	Fk_Module_GModel=Trim(Request.Form("Fk_Module_GModel"))
	Fk_Module_IsIndex=Trim(Request.Form("Fk_Module_IsIndex"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Module_Name,1,50,0,"请输入模块名称！","模块名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Subhead,0,50,0,"请输入模块副标题！","副标题不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入关键字！","关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入描述！","描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入模块目录名/文件名！","模块目录名/文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Module_Type,"请选择模块类型！")
	Call FKFun.ShowNum(Fk_Module_Level,"请选择模块分级！")
	Call FKFun.ShowNum(Fk_Module_Show,"请选择模块是否菜单显示！")
	Call FKFun.ShowNum(Fk_Module_Template,"请选择模块模板！")
	Call FKFun.ShowNum(Fk_Module_Menu,"系统参数错误，请刷新页面！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	If Fk_Module_Type=0 Then
		If Fk_Module_IsIndex<>"" Then
			Call FKFun.ShowNum(Fk_Module_IsIndex,"请选择是否模块首页！")
		Else
			Fk_Module_IsIndex=0
		End If
		' ===== 此处代码移动到下方 ===== 
	Else
		Fk_Module_IsIndex=0
	End If

	' ===== Fk_Menu_Dir的代码先移动出来 ===== 
	Sqlstr="Select Fk_Menu_Dir From [Fk_Menu] Where Fk_Menu_Id=" & Fk_Module_Menu
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Dir=Rs("Fk_Menu_Dir")
	End If
	Rs.Close
	' ===== End ===== 
	
	Select Case Fk_Module_Type
		Case 0
		Case 1
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
		Case 2
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
		Case 3
		Case 4
			Call FKFun.ShowString(Fk_Module_GUrl,0,255,0,"请输入留言后跳转链接！","留言后跳转链接不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_GModel,"请选择留言模型！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
		Case 5
			Call FKFun.ShowString(Fk_Module_Url,1,255,0,"请输入转向链接！","转向链接不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_UrlType,"请选择跳转模式！")
		Case 6
		Case 7
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择模块子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			Call FKFun.ShowString(Fk_Module_MName,0,50,0,"请输入管理名称！","管理名称不能大于50个字符！")
	End Select
	If Id=Fk_Module_Level Then
		Call FKFun.ShowErr("自己不能成为自己的分类哦！",2)
	End If
	Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%' And Fk_Module_Id=" & Fk_Module_Level
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("不能转移到子分类旗下！如需要请先中转一个分类！",2)
	End If
	Rs.Close
	If Fk_Module_Dir<>"" Then
		Call FKAdmin.CheckDandF("dir",Fk_Module_Dir,Fk_Module_Level)
		If Fk_Site_HtmlType=0 Then
			' ===== 增加Fk_Module_Menu判断，即同一menu下目录名称不能相同 =====
			Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"' And Fk_Module_Id<>"&Id&" AND Fk_Module_Menu="&Fk_Module_Menu&""
		ElseIf Fk_Site_HtmlType=1 Then
			' ===== 增加Fk_Module_Menu判断，即同一menu下目录名称不能相同 ===== 
			Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"' And Fk_Module_Level="&Fk_Module_Level&"  And Fk_Module_Id<>"&Id&" AND Fk_Module_Menu="&Fk_Module_Menu&""
		End If
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Rs.Close
			Call FKFun.ShowErr("模块文件名/目录已经被使用，请重新输入！",2)
		End If
		Rs.Close
	End If
	If Fk_Module_Level>0 Then
		Fk_Module_LevelList=FKAdmin.GetModuleLevelList(Fk_Module_Level)
	Else
		Fk_Module_LevelList=""
	End If
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Type,Fk_Module_Dir,Fk_Module_MUrl,Fk_Module_PageCount,Fk_Module_Menu,Fk_Module_Level,Fk_Module_LevelList,Fk_Module_Template,Fk_Module_LowTemplate,Fk_Module_UrlType,Fk_Module_Url,Fk_Module_Show,Fk_Module_MenuShow,Fk_Module_GModel,Fk_Module_MName,Fk_Module_Pic,Fk_Module_Subhead,Fk_Module_IsIndex,Fk_Module_GUrl From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Dim OldLevelList
		OldLevelList=Rs("Fk_Module_LevelList")
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Subhead")=Fk_Module_Subhead
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		If Rs("Fk_Module_Dir")="" Or IsNull(Rs("Fk_Module_Dir")) Then
			Rs("Fk_Module_Dir")=Fk_Module_Dir
		End If
		Rs("Fk_Module_Type")=Fk_Module_Type
		Rs("Fk_Module_Menu")=Fk_Module_Menu
		Rs("Fk_Module_Level")=Fk_Module_Level
		Rs("Fk_Module_LevelList")=Fk_Module_LevelList
		Rs("Fk_Module_Template")=Fk_Module_Template
		Rs("Fk_Module_Show")=Fk_Module_Show
		Rs("Fk_Module_MenuShow")=Fk_Module_MenuShow
		Rs("Fk_Module_MUrl")=""
		Rs("Fk_Module_Pic")=Fk_Module_Pic
		Rs("Fk_Module_IsIndex")=Fk_Module_IsIndex
		Select Case Fk_Module_Type
			Case 0
			Case 1
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
			Case 2
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
			Case 3
			Case 4
				Rs("Fk_Module_GUrl")=Fk_Module_GUrl
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_GModel")=Fk_Module_GModel
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 5
				Rs("Fk_Module_Url")=Fk_Module_Url
				Rs("Fk_Module_UrlType")=Fk_Module_UrlType
			Case 6
			Case 7
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
				Rs("Fk_Module_MName")=Fk_Module_MName
		End Select
		Rs.Update()
		If Fk_Module_LevelList<>OldLevelList Then
			Rs.Close
			Sqlstr="Select Fk_Module_Id,Fk_Module_LevelList From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%'"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
				Rs("Fk_Module_LevelList")=Replace(Rs("Fk_Module_LevelList"),OldLevelList,Fk_Module_LevelList)
				Rs.Update
				Rs.MoveNext
			Wend
		End If
		Rs.Close

		
		'Sqlstr="Select Fk_Module_Id,Fk_Module_MUrl,Fk_Module_IsIndex,Fk_Module_Type,Fk_Module_LevelList,Fk_Module_Dir,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Name='"&Fk_Module_Name&"' And Fk_Module_Level="&Fk_Module_Level&""
		' ===== 此处使用ID能避免错误 ===== 
		Sqlstr="Select Fk_Module_Id,Fk_Module_MUrl,Fk_Module_IsIndex,Fk_Module_Type,Fk_Module_LevelList,Fk_Module_Dir,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Id="&Id&""
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			If Fk_Module_IsIndex=0 Then
				Rs("Fk_Module_MUrl")=FKAdmin.ModuleMUrl(Rs("Fk_Module_Id"),Rs("Fk_Module_Type"),Rs("Fk_Module_LevelList"),Rs("Fk_Module_UrlType"),Rs("Fk_Module_Url"),Rs("Fk_Module_Dir"),Fk_Menu_Dir,Fk_Site_HtmlType,Fk_Site_HtmlSuffix)
			Else
				Rs("Fk_Module_MUrl")=Fk_Menu_Dir&"/"
			End If
			Rs.Update()
		End If
		Application.UnLock()
		Response.Write("模块修改成功！")
	Else
		Response.Write("模块不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ModuleDelDo
'作    用：执行删除模块
'参    数：
'==============================
Sub ModuleDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("此模块有子模块，暂无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("模块删除成功！")
	Else
		Response.Write("模块不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ModuleOrderForm
'作    用：模块排序表单
'参    数：
'==============================
Sub ModuleOrderForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Sqlstr="Select Fk_Menu_Name From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	Else
		Call FKFun.ShowErr("菜单不存在！",2)
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Module.asp?Type=1&MenuId=<%=MenuId%>');">返回模块列表</a></li>
    </ul>
</div>
<div id="ListTop">
    “<%=Fk_Menu_Name%>”菜单模块排序
</div>
<div id="ListContent">
	<form id="ModuleOrderSet" name="ModuleOrderSet" method="post" action="Module.asp?Type=8" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">模块名称</td>
            <td align="center" class="ListTdTop">模块类型</td>
            <td align="center" class="ListTdTop">模块排序</td>
        </tr>
<%
	Call FKAdmin.GetModuleList(5,MenuId,0,0,"")
%>
        <tr>
            <td height="30" colspan="4" style="text-indent:32px;">
            <input type="hidden" name="MenuId" value="<%=MenuId%>" />
            <input type="submit" name="Enter" class="Button" id="Enter" value="设置" onclick="Sends('ModuleOrderSet','Module.asp?Type=8',0,'',0,1,'MainRight','Module.asp?Type=1&MenuId=<%=MenuId%>');" />&nbsp;&nbsp;
            <input type="reset" name="ReSet" class="Button" id="ReSet" value="重置" />
            </td>
        </tr>
    </table>
    </form>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==============================
'函 数 名：ModuleOrderDo
'作    用：执行模块排序
'参    数：
'==============================
Sub ModuleOrderDo()
	MenuId=Trim(Request.Form("MenuId"))
	Call FKFun.ShowNum(MenuId,"MenuId系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Order From [Fk_Module] Where Fk_Module_Menu="&MenuId&" Order By Fk_Module_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	While Not Rs.Eof
		Fk_Module_Order=Trim(Request.Form("Fk_Module_Order"&Rs("Fk_Module_Id")))
		Call FKFun.ShowNum(Fk_Module_Order,Rs("Fk_Module_Name")&"模块的序号不是数字，排序序号必须是有效数字！")
		Rs("Fk_Module_Order")=Fk_Module_Order
		Rs.Update()
		Rs.MoveNext
	Wend
	Application.UnLock()
	Rs.Close
	Response.Write("模块排序成功！")
End Sub

'==============================
'函 数 名：ReUrlDo
'作    用：重置模块链接
'参    数：
'==============================
Sub ReUrlDo()
	MenuId=Trim(Request.QueryString("MenuId"))
	Call FKFun.ShowNum(MenuId,"MenuId系统参数错误，请刷新页面！")
	Call FKAdmin.ReLoadModuleUrl(MenuId,Fk_Site_HtmlType,Fk_Site_HtmlSuffix)
	Response.Write("模块链接重置成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->