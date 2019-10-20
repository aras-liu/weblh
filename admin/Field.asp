<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Field.asp
'文件用途：自定义字段管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System9",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Field_Name,Fk_Field_Tag,Fk_Field_Content,Fk_Field_Type,Fk_Field_Help,Fk_Field_Model,Fk_Field_Option,Fk_Field_Remark

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FieldList() '自定义字段列表
	Case 2
		Call FieldFielddForm() '添加自定义字段表单
	Case 3
		Call FieldFielddDo() '执行添加自定义字段
	Case 4
		Call FieldEditForm() '修改自定义字段表单
	Case 5
		Call FieldEditDo() '执行修改自定义字段
	Case 6
		Call FieldDelDo() '执行删除自定义字段
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FieldList()
'作    用：自定义字段列表
'参    数：
'==========================================
Sub FieldList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Field.asp?Type=2');">添加新字段</a></li>
    </ul>
</div>
<div id="ListTop">
    自定义字段管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">备注</td>
            <td align="center" class="ListTdTop">用途</td>
            <td align="center" class="ListTdTop">类型</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Field_Id,Fk_Field_Name,Fk_Field_Type,Fk_Field_Tag,Fk_Field_Model,Fk_Field_Remark From [Fk_Field] Order By Fk_Field_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
			Select Case Rs("Fk_Field_Type")
				Case 0
					Fk_Field_Type="文本框"
				Case 1
					Fk_Field_Type="编辑器框"
				Case 2
					Fk_Field_Type="上传框"
				Case 3
					Fk_Field_Type="下拉框"
			End Select
			Select Case Rs("Fk_Field_Model")
				Case 0
					Fk_Field_Model="内容页面"
				Case 1
					Fk_Field_Model="站点设置"
			End Select
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td>&nbsp;<%=Rs("Fk_Field_Name")%></td>
            <td>&nbsp;<%=Rs("Fk_Field_Remark")%></td>
            <td align="center"><%=Fk_Field_Model%></td>
            <td align="center"><%=Fk_Field_Type%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Field.asp?Type=4&Id=<%=Rs("Fk_Field_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Field_Name")%>”？此操作不可逆！','Field.asp?Type=6&Id=<%=Rs("Fk_Field_Id")%>','MainRight','Field.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="6" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="7">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FieldFielddForm()
'作    用：添加自定义字段表单
'参    数：
'==========================================
Sub FieldFielddForm()
%>
<form id="FieldFieldd" name="FieldFieldd" method="post" action="Field.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:800px;">添加新自定义字段[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:800px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Field_Name" type="text" class="Input" id="Fk_Field_Name" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>自定义字段名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">备注：</td>
	        <td>&nbsp;<input name="Fk_Field_Remark" type="text" class="Input" id="Fk_Field_Remark" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>用于识别自定义字段的用途，此段文字仅显示在自定义字段列表上，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">标签：</td>
            <td>&nbsp;<input name="Fk_Field_Tag" type="text" class="Input" id="Fk_Field_Tag" size="20" />*输入字母或数字，添加后不可修改&nbsp;&nbsp;<span class="qbox" title="<p>自定义字段标签，请输入字母或数字，请输入1-50个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    <tr>
	        <td height="25" align="right">帮助语句：</td>
	        <td>&nbsp;<input name="Fk_Field_Help" type="text" class="Input" id="Fk_Field_Help" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>帮助语句，显示在输入框边上，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">输入类型：</td>
            <td>&nbsp;<select name="Fk_Field_Type" class="Input" id="Fk_Field_Type" onchange="ChangeFieldType(this.options[this.selectedIndex].value)">
                    <option value="0">文本框</option>
                    <option value="1">编辑器框</option>
                    <option value="2">上传框</option>
                    <option value="3">下拉框</option>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择输入类型。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">用途：</td>
            <td>&nbsp;<select name="Fk_Field_Model" class="Input" id="Fk_Field_Model" onchange="ChangeField(this.options[this.selectedIndex].value)">
                    <option value="0">内容页面</option>
                    <option value="1">站点设置</option>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>自定义字段用途。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="Fk_Field_Options" style="display:none;">
            <td height="30" align="right">选项：</td>
            <td>&nbsp;<textarea name="Fk_Field_Option" class="TextArea" cols="60" rows="5" id="Fk_Field_Option"></textarea>&nbsp;&nbsp;<span class="qbox" title="<p>选项每行一个：选项值||选项字符串。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="FieldUser">
            <td height="30" align="right">适用模块&nbsp;&nbsp;<span class="qbox" title="<p>可以指定模块类型，也可以指定适用的模块，如果指定了适用的模块，则新增模块不会自动加入。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td>
            	<ul class="triState">
                	<li><span class="title">适用模块类型</span>
                    	<ul>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Article" /><label href="#" class="label">所有文章模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Product" /><label href="#" class="label">所有产品模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Down" /><label href="#" class="label">所有下载模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Info" /><label href="#" class="label">所有信息模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Job" /><label href="#" class="label">所有招聘模块</label></li>
                        </ul>
                    </li>
                </ul>
                <ul class="triState">
                    <li><span class="title">适用模块</span>
                        <ul>
<%
	Dim MenuList
	Sqlstr="Select Fk_Menu_Id,Fk_Menu_Name From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof 
		If MenuList="" Then
			MenuList=Rs("Fk_Menu_Id")&"||"&Rs("Fk_Menu_Name")
		Else
			MenuList=MenuList&","&Rs("Fk_Menu_Id")&"||"&Rs("Fk_Menu_Name")
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	TempArr=Split(MenuList,",")
	For Each Temp In TempArr
		

%>
                            <li><span class="title"><%=Split(Temp,"||")(1)%></span>
<%
		Call FKAdmin.GetModuleList(3,Split(Temp,"||")(0),0,0,"")
%>
                            </li>
<%
	Next
%>            
                        </ul>
                    </li>
                </ul>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:780px;">
        <input type="submit" onclick="Sends('FieldFieldd','Field.asp?Type=3',0,'',0,1,'MainRight','Field.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FieldFielddDo
'作    用：执行添加自定义字段
'参    数：
'==============================
Sub FieldFielddDo()
	Fk_Field_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Name")))
	Fk_Field_Tag=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Tag")))
	Fk_Field_Help=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Help")))
	Fk_Field_Option=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Option")))
	Fk_Field_Remark=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Remark")))
	Fk_Field_Content=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Field_Content")),", ",","))&","
	Fk_Field_Type=Trim(Request.Form("Fk_Field_Type"))
	Fk_Field_Model=Trim(Request.Form("Fk_Field_Model"))
	Call FKFun.ShowString(Fk_Field_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Field_Help,1,255,0,"请输入帮助语句！","帮助语句不能大于255个字符！")
	Call FKFun.ShowString(Fk_Field_Remark,0,255,0,"请输入备注！","备注不能大于255个字符！")
	Call FKFun.ShowString(Fk_Field_Tag,1,50,0,"请输入标签！","标签不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Field_Type,"请选择自定义字段类型！")
	Call FKFun.ShowNum(Fk_Field_Model,"请选择自定义字段用途！")
	If Fk_Field_Type=3 Then
		Call FKFun.ShowString(Fk_Field_Option,1,10000,0,"请输入选项！","选项不能大于255个字符！")
	End If
	Sqlstr="Select Fk_Field_Id,Fk_Field_Name,Fk_Field_Tag,Fk_Field_Type,Fk_Field_Help,Fk_Field_Content,Fk_Field_Model,Fk_Field_Remark,Fk_Field_Option From [Fk_Field] Where Fk_Field_Name='"&Fk_Field_Name&"' Or Fk_Field_Tag='"&Fk_Field_Tag&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Field_Name")=Fk_Field_Name
		Rs("Fk_Field_Tag")=Fk_Field_Tag
		Rs("Fk_Field_Type")=Fk_Field_Type
		Rs("Fk_Field_Help")=Fk_Field_Help
		Rs("Fk_Field_Content")=Fk_Field_Content
		Rs("Fk_Field_Model")=Fk_Field_Model
		Rs("Fk_Field_Remark")=Fk_Field_Remark
		Rs("Fk_Field_Option")=Fk_Field_Option
		Rs.Update()
		Application.UnLock()
		Response.Write("新自定义字段添加成功！")
	Else
		Response.Write("该自定义字段已经被存在，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：FieldEditForm()
'作    用：修改自定义字段表单
'参    数：
'==========================================
Sub FieldEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Field_Name,Fk_Field_Tag,Fk_Field_Type,Fk_Field_Help,Fk_Field_Content,Fk_Field_Model,Fk_Field_Remark,Fk_Field_Option From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Field_Name=FKFun.HTMLDncode(Rs("Fk_Field_Name"))
		Fk_Field_Tag=FKFun.HTMLDncode(Rs("Fk_Field_Tag"))
		Fk_Field_Remark=FKFun.HTMLDncode(Rs("Fk_Field_Remark"))
		Fk_Field_Option=FKFun.HTMLDncode(Rs("Fk_Field_Option"))
		Fk_Field_Type=Rs("Fk_Field_Type")
		Fk_Field_Help=Rs("Fk_Field_Help")
		Fk_Field_Content=Rs("Fk_Field_Content")
		Fk_Field_Model=Rs("Fk_Field_Model")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此自定义字段，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="FieldEdit" name="FieldEdit" method="post" action="Field.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:800px;">修改自定义字段[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:800px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Field_Name" value="<%=Fk_Field_Name%>" type="text" class="Input" id="Fk_Field_Name" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>自定义字段名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">备注：</td>
	        <td>&nbsp;<input name="Fk_Field_Remark" value="<%=Fk_Field_Remark%>" type="text" class="Input" id="Fk_Field_Remark" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>用于识别自定义字段的用途，此段文字仅显示在自定义字段列表上，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">标签：</td>
            <td>&nbsp;<input name="Fk_Field_Tag" value="<%=Fk_Field_Tag%>" type="text" class="Input" id="Fk_Field_Tag" size="20" disabled="disabled" />&nbsp;&nbsp;<span class="qbox" title="<p>不可修改。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    <tr>
	        <td height="25" align="right">帮助语句：</td>
	        <td>&nbsp;<input name="Fk_Field_Help" value="<%=Fk_Field_Help%>" type="text" class="Input" id="Fk_Field_Help" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>帮助语句，显示在输入框边上，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">输入类型：</td>
            <td>&nbsp;<select name="Fk_Field_Type" class="Input" id="Fk_Field_Type" onchange="ChangeFieldType(this.options[this.selectedIndex].value)">
                    <option value="0"<%=FKFun.BeSelect(Fk_Field_Type,0)%>>文本框</option>
                    <option value="1"<%=FKFun.BeSelect(Fk_Field_Type,1)%>>编辑器框</option>
                    <option value="2"<%=FKFun.BeSelect(Fk_Field_Type,2)%>>上传框</option>
                    <option value="3"<%=FKFun.BeSelect(Fk_Field_Type,3)%>>下拉框</option>
                    </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择输入类型。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">用途：</td>
            <td>&nbsp;<select name="Fk_Field_Model" class="Input" id="Fk_Field_Model" onchange="ChangeField(this.options[this.selectedIndex].value)">
                    <option value="0"<%=FKFun.BeSelect(Fk_Field_Model,0)%>>内容页面</option>
                    <option value="1"<%=FKFun.BeSelect(Fk_Field_Model,1)%>>站点设置</option>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p>自定义字段用途。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="Fk_Field_Options"<%If Fk_Field_Type<>3 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right">选项：</td>
            <td>&nbsp;<textarea name="Fk_Field_Option" class="TextArea" cols="60" rows="5" id="Fk_Field_Option"><%=Fk_Field_Option%></textarea>&nbsp;&nbsp;<span class="qbox" title="<p>选项每行一个：选项值||选项字符串。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr id="FieldUser"<%If Fk_Field_Model=1 Then%> style="display:none;"<%End If%>>
            <td height="30" align="right">适用模块&nbsp;&nbsp;<span class="qbox" title="<p>可以指定模块类型，也可以指定适用的模块，如果指定了适用的模块，则新增模块不会自动加入。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td>
            	<ul class="triState">
                	<li><span class="title">适用模块类型</span>
                    	<ul>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Article"<%If Instr(Fk_Field_Content,",Article,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">所有文章模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Product"<%If Instr(Fk_Field_Content,",Product,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">所有产品模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Down"<%If Instr(Fk_Field_Content,",Down,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">所有下载模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Info"<%If Instr(Fk_Field_Content,",Info,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">所有信息模块</label></li>
                            <li><input type="checkbox" name="Fk_Field_Content" value="Job"<%If Instr(Fk_Field_Content,",Job,")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">所有招聘模块</label></li>
                        </ul>
                    </li>
                </ul>
                <ul class="triState">
                    <li><span class="title">适用模块</span>
                        <ul>
<%
	Dim MenuList
	Sqlstr="Select Fk_Menu_Id,Fk_Menu_Name From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof 
		If MenuList="" Then
			MenuList=Rs("Fk_Menu_Id")&"||"&Rs("Fk_Menu_Name")
		Else
			MenuList=MenuList&","&Rs("Fk_Menu_Id")&"||"&Rs("Fk_Menu_Name")
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	TempArr=Split(MenuList,",")
	For Each Temp In TempArr
		

%>
                            <li><span class="title"><%=Split(Temp,"||")(1)%></span>
<%
		Call FKAdmin.GetModuleList(3,Split(Temp,"||")(0),0,0,Fk_Field_Content)
%>
                            </li>
<%
	Next
%>            
                        </ul>
                    </li>
                </ul>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:780px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('FieldEdit','Field.asp?Type=5',0,'',0,1,'MainRight','Field.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FieldEditDo
'作    用：执行修改自定义字段
'参    数：
'==============================
Sub FieldEditDo()
	Fk_Field_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Name")))
	Fk_Field_Content=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Field_Content")),", ",","))&","
	Fk_Field_Help=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Help")))
	Fk_Field_Option=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Option")))
	Fk_Field_Remark=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Remark")))
	Fk_Field_Type=Trim(Request.Form("Fk_Field_Type"))
	Fk_Field_Model=Trim(Request.Form("Fk_Field_Model"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Field_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Field_Help,1,255,0,"请输入帮助语句！","帮助语句不能大于255个字符！")
	Call FKFun.ShowString(Fk_Field_Remark,0,255,0,"请输入备注！","备注不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Field_Type,"请选择自定义字段类型！")
	Call FKFun.ShowNum(Fk_Field_Model,"请选择自定义字段用途！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	If Fk_Field_Type=3 Then
		Call FKFun.ShowString(Fk_Field_Option,1,10000,0,"请输入选项！","选项不能大于255个字符！")
	End If
	Sqlstr="Select Fk_Field_Id,Fk_Field_Name,Fk_Field_Type,Fk_Field_Help,Fk_Field_Content,Fk_Field_Model,Fk_Field_Option,Fk_Field_Remark From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Field_Name")=Fk_Field_Name
		Rs("Fk_Field_Type")=Fk_Field_Type
		Rs("Fk_Field_Help")=Fk_Field_Help
		Rs("Fk_Field_Content")=Fk_Field_Content
		Rs("Fk_Field_Model")=Fk_Field_Model
		Rs("Fk_Field_Option")=Fk_Field_Option
		Rs("Fk_Field_Remark")=Fk_Field_Remark
		Rs.Update()
		Application.UnLock()
		Response.Write("自定义字段修改成功！")
	Else
		Response.Write("自定义字段不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：FieldDelDo
'作    用：执行删除自定义字段
'参    数：
'==============================
Sub FieldDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Field_Id From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("自定义字段删除成功！")
	Else
		Response.Write("自定义字段不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->