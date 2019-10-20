<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Info.asp
'文件用途：信息管理拉取页面
'版权所有：方卡在线
'==========================================

'定义页面变量
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Content,Fk_Module_Template,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Field,fenzhan
Set FKHtml=New Cls_Html
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call InfoForm() '修改信息表单
	Case 2
		Call InfoDo() '执行修改信息
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：InfoForm()
'作    用：修改信息表单
'参    数：
'==========================================
Sub InfoForm()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select Fk_Module_Name,Fk_Module_Field,Fk_Module_Content,Fk_Module_Template,Fk_Module_Keyword,Fk_Module_Description From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Content=Rs("Fk_Module_Content")
		Fk_Module_Template=Rs("Fk_Module_Template")
		Fk_Module_Keyword=Rs("Fk_Module_Keyword")
		Fk_Module_Description=Rs("Fk_Module_Description")
		If IsNull(Rs("Fk_Module_Field")) Or Rs("Fk_Module_Field")="" Then
			Fk_Module_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Module_Field=Split(Rs("Fk_Module_Field"),"[-Fangka_Field-]")
		End If
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="InfoEdit" name="InfoEdit" method="post" action="Info.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">修改“<%=Fk_Module_Name%>”[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
<table width="95%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="11%" height="28" align="right">信息标题：</td>
        <td width="89%">&nbsp;<input name="Fk_Module_Name" value="<%=Fk_Module_Name%>" type="text" class="Input" id="Fk_Module_Name" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>信息标题，不能为空，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">关键字：</td>
        <td>&nbsp;<input name="Fk_Module_Keyword" value="<%=Fk_Module_Keyword%>" type="text" class="Input" id="Fk_Module_Keyword" size="50" />&nbsp;&nbsp;<span id="th3" class="qbox" title="<p>多个关键字用英文逗号隔开，用于页面meta的keywords，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">描述：</td>
        <td>&nbsp;<input name="Fk_Module_Description" value="<%=Fk_Module_Description%>" type="text" class="Input" id="Fk_Module_Description" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入站点的一段文字简介，用于页面meta的description，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
	Call FKAdmin.ShowField(0,0," And (Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Fk_Module_Field,"")
	Call FKAdmin.ShowField(2,0," And (Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Fk_Module_Field,"")
	Call FKAdmin.ShowField(3,0," And (Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Fk_Module_Field,"")
%>
    <tr>
        <td height="28" align="right">模板：</td>
        <td>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
            <option value="0"<%=FKFun.BeSelect(Fk_Module_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select Fk_Template_Id,Fk_Template_Name From [Fk_Template] Where "&NoDirStr&""
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择本页面使用的模板。</p>"><img src="Images/help.jpg" /></span></td>
    </tr>
    <tr>
        <td height="28" align="right">信息内容&nbsp;&nbsp;<span class="qbox" title="<p>信息内容，不能为空，请输入5-1000000个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span>：</td>
        <td><textarea name="Fk_Module_Content" id="Fk_Module_Content" class="<%=EditorClass%>" rows="20" style="width:100%;"><%=Fk_Module_Content%></textarea></td>
    </tr>
<%
	Call FKAdmin.ShowField(1,0," And (Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')",Fk_Module_Field,EditorClass)
%>
</table>
</div>
<div id="BoxBottom" style="width:880px;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('InfoEdit','Info.asp?Type=2',0,'',0,0,'','');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：InfoDo()
'作    用：执行修改信息
'参    数：
'==========================================
Sub InfoDo()
	Dim TempMUrl,TempShow
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("您无此权限！",2)
	End If
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Content=Request.Form("Fk_Module_Content")
	Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Call FKFun.ShowString(Fk_Module_Name,1,255,0,"请输入信息标题！","信息标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Content,5,1,1,"请输入信息内容，不少于5个字符！","信息内容不能大于1个字符！")
	Call FKFun.ShowNum(Fk_Module_Template,"请选择模块模板！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入关键字！","关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入描述！","描述不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Fk_Module_Field=FKAdmin.GetFieldData(0,"(Fk_Field_Content Like '%%,Info,%%' Or Fk_Field_Content Like '%%,Module"&Fk_Module_Id&",%%')")
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Field,Fk_Module_Content,Fk_Module_Template,Fk_Module_MUrl,Fk_Module_Show From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		TempMUrl=FKTemplate.GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
		TempShow=Rs("Fk_Module_Show")
		Application.Lock()
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		Rs("Fk_Module_Field")=Fk_Module_Field
		Rs("Fk_Module_Content")=Fk_Module_Content
		Rs("Fk_Module_Template")=Fk_Module_Template
		Rs.Update()
		Application.UnLock()
		Response.Write("“"&Fk_Module_Name&"”修改成功！")
		Rs.Close
		If Fk_Site_Html=2 And TempShow=1 Then
			Id=Fk_Module_Id
			Response.Write("<span style='display:none;'>")
			Call FKHtml.CreatModule(Fk_Module_Id,3,TempMUrl,Fk_Module_Name,"",0,1)
			Response.Write("</span>")
		End If
	Else
		Rs.Close
		Response.Write("信息不存在！")
	End If
End Sub
%>
<!--#Include File="../Code.asp"-->