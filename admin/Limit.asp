<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Admin/Limit.asp
'文件用途：用户权限管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"0",Request.Cookies("FkAdminLimit1"))

'定义页面变量
Dim Fk_Limit_Name,Fk_Limit_Set1,Fk_Limit_Set2,Fk_Limit_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call LimitList() '权限列表
	Case 2
		Call LimitAddForm() '添加权限表单
	Case 3
		Call LimitAddDo() '执行添加权限
	Case 4
		Call LimitEditForm() '修改权限表单
	Case 5
		Call LimitEditDo() '执行修改权限
	Case 6
		Call LimitDelDo() '执行删除权限
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：LimitList()
'作    用：权限列表
'参    数：
'==========================================
Sub LimitList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Limit.asp?Type=2');">添加新权限</a></li>
    </ul>
</div>
<div id="ListTop">
    权限管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">序号</td>
            <td align="center" class="ListTdTop">权限名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Limit_Id,Fk_Limit_Name From [Fk_Limit] Order By Fk_Limit_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_Limit_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Limit.asp?Type=4&Id=<%=Rs("Fk_Limit_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Limit_Name")%>”，此操作不可逆！','Limit.asp?Type=6&Id=<%=Rs("Fk_Limit_Id")%>','MainRight','Limit.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="3" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="3">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：LimitAddForm()
'作    用：添加权限表单
'参    数：
'==========================================
Sub LimitAddForm()
%>
<form id="LimitAdd" name="LimitAdd" method="post" action="Limit.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新权限[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="18%" height="25" align="right">权限名称：</td>
	        <td width="82%">&nbsp;<input name="Fk_Limit_Name" type="text" class="Input" id="Fk_Limit_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>权限名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">系统权限&nbsp;&nbsp;<span class="qbox" title="<p>系统权限包括常规管理的权限和内容设置的权限两部分，勾选则说明允许管理选中功能。</p>"><img src="Images/help.jpg" /></span>：</td>
	        <td>
                <ul class="triState">
                    <li><span class="title">系统权限</span>
                        <ul>
                            <li><span class="title">常规管理权限</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System1" /><label href="#" class="label">系统设置</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System2" /><label href="#" class="label">模板管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System3" /><label href="#" class="label">生成管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System4" /><label href="#" class="label">过滤字符管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System5" /><label href="#" class="label">关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System6" /><label href="#" class="label">搜索引擎地图生成</label></li>
                                </ul>
                            </li>
                            <li><span class="title">内容设置权限</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System1" /><label href="#" class="label">菜单管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System2" /><label href="#" class="label">友情连接管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System3" /><label href="#" class="label">广告管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System4" /><label href="#" class="label">站内关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System5" /><label href="#" class="label">推荐管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System6" /><label href="#" class="label">专题管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System7" /><label href="#" class="label">在线投票管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System8" /><label href="#" class="label">独立信息管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System9" /><label href="#" class="label">自定义字段管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System10" /><label href="#" class="label">客服浮窗代码管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System11" /><label href="#" class="label">留言模型管理</label></li>
                                </ul>
                            </li>
                        </ul>
                    </li>
                </ul>
                <div class="Cal"></div>
            </td>
	        </tr>
	    <tr>
	        <td height="25" align="right">内容管理权限&nbsp;&nbsp;<span class="qbox" title="<p>内容管理权限用于设置用户管理内容的权限，会按菜单输出可管理项目，勾选则说明用户可管理此项目。</p>"><img src="Images/help.jpg" /></span>：</td>
	        <td>
                <ul class="triState">
                    <li><span class="title">内容管理权限</span>
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
		Call FKAdmin.GetModuleList(2,Split(Temp,"||")(0),0,0,"")
%>
                            </li>
<%
	Next
%>            
                        </ul>
                    </li>
                </ul>
                <div class="Cal"></div>
            </td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('LimitAdd','Limit.asp?Type=3',0,'',0,1,'MainRight','Limit.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：LimitAddDo
'作    用：执行添加权限
'参    数：
'==============================
Sub LimitAddDo()
	Dim Temp2,TempArr2
	Fk_Limit_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Limit_Name")))
	Fk_Limit_Set1=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Set1"))," ",""))&","
	Fk_Limit_Set2=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Set2"))," ",""))&","
	Temp=FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Content"))," ",""))
	Call FKFun.ShowString(Fk_Limit_Name,1,50,0,"请输入权限名称！","权限名称不能大于50个字符！")
	Fk_Limit_Content=","
	TempArr=Split(Temp,",")
	For Each Temp In TempArr
		Fk_Limit_Content=Fk_Limit_Content&"Module"&Temp&","
	Next
	For Each Temp In TempArr
		Sqlstr="Select Fk_Module_LevelList,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Temp
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Temp2=Rs("Fk_Module_LevelList")
			If Instr(Fk_Limit_Content,",Menu"&Rs("Fk_Module_Menu")&",")=0 Then
				Fk_Limit_Content=Fk_Limit_Content&"Menu"&Rs("Fk_Module_Menu")&","
			End if
		End If
		Rs.Close
		If Temp2<>"" Then
			TempArr2=Split(Temp2,",")
			For Each Temp2 In TempArr2
				If Temp2<>"" Then
					If Instr(Fk_Limit_Content,",Module"&Temp2&",")=0 And Instr(Fk_Limit_Content,",See"&Temp2&",")=0 Then
						Fk_Limit_Content=Fk_Limit_Content&"See"&Temp2&","
					End If
				End If
			Next
		End If
	Next
	Sqlstr="Select Fk_Limit_Id,Fk_Limit_Name,Fk_Limit_Set1,Fk_Limit_Set2,Fk_Limit_Content From [Fk_Limit] Where Fk_Limit_Name='"&Fk_Limit_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Limit_Name")=Fk_Limit_Name
		Rs("Fk_Limit_Set1")=Fk_Limit_Set1
		Rs("Fk_Limit_Set2")=Fk_Limit_Set2
		Rs("Fk_Limit_Content")=Fk_Limit_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新权限添加成功！")
	Else
		Response.Write("该权限名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：LimitEditForm()
'作    用：修改权限表单
'参    数：
'==========================================
Sub LimitEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Limit_Name,Fk_Limit_Set1,Fk_Limit_Set2,Fk_Limit_Content From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Limit_Name=Rs("Fk_Limit_Name")
		Fk_Limit_Set1=Rs("Fk_Limit_Set1")
		Fk_Limit_Set2=Rs("Fk_Limit_Set2")
		Fk_Limit_Content=Rs("Fk_Limit_Content")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此权限，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="LimitEdit" name="LimitEdit" method="post" action="Limit.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改权限[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="18%" height="25" align="right">权限名称：</td>
	        <td width="82%">&nbsp;<input name="Fk_Limit_Name" type="text" value="<%=Fk_Limit_Name%>" class="Input" id="Fk_Limit_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>权限名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">系统权限&nbsp;&nbsp;<span class="qbox" title="<p>系统权限包括常规管理的权限和内容设置的权限两部分，勾选则说明允许管理选中功能。</p>"><img src="Images/help.jpg" /></span>：</td>
	        <td>
                <ul class="triState">
                    <li><span class="title">系统权限</span>
                        <ul>
                            <li><span class="title">常规管理权限</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System1"<%If Instr(Fk_Limit_Set1,"System1")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">系统设置</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System2"<%If Instr(Fk_Limit_Set1,"System2")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">模板管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System3"<%If Instr(Fk_Limit_Set1,"System3")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">生成管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System4"<%If Instr(Fk_Limit_Set1,"System4")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">过滤字符管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System5"<%If Instr(Fk_Limit_Set1,"System5")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set1" value="System6"<%If Instr(Fk_Limit_Set1,"System6")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">搜索引擎地图生成</label></li>
                                </ul>
                            </li>
                            <li><span class="title">内容设置权限</span>
                                <ul>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System1"<%If Instr(Fk_Limit_Set2,"System1")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">菜单管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System2"<%If Instr(Fk_Limit_Set2,"System2")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">友情连接管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System3"<%If Instr(Fk_Limit_Set2,"System3")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">广告管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System4"<%If Instr(Fk_Limit_Set2,"System4")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">站内关键字管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System5"<%If Instr(Fk_Limit_Set2,"System5")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">推荐管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System6"<%If Instr(Fk_Limit_Set2,"System6")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">专题管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System7"<%If Instr(Fk_Limit_Set2,"System7")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">在线投票管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System8"<%If Instr(Fk_Limit_Set2,"System8")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">独立信息管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System9"<%If Instr(Fk_Limit_Set2,"System9")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">自定义字段管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System10"<%If Instr(Fk_Limit_Set2,"System10")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">客服浮窗代码管理</label></li>
                                    <li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Set2" value="System11"<%If Instr(Fk_Limit_Set2,"System11")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label">留言模型管理</label></li>
                                </ul>
                            </li>
                        </ul>
                    </li>
                </ul>
                <div class="Cal"></div>
            </td>
	        </tr>
	    <tr>
	        <td height="25" align="right">内容管理权限&nbsp;&nbsp;<span class="qbox" title="<p>内容管理权限用于设置用户管理内容的权限，会按菜单输出可管理项目，勾选则说明用户可管理此项目。</p>"><img src="Images/help.jpg" /></span>：</td>
	        <td>
                <ul class="triState">
                    <li><span class="title">内容管理权限</span>
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
		Call FKAdmin.GetModuleList(2,Split(Temp,"||")(0),0,0,Fk_Limit_Content)
%>
                            </li>
<%
	Next
%>            
                        </ul>
                    </li>
                </ul>
                <div class="Cal"></div>
            </td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('LimitEdit','Limit.asp?Type=5',0,'',0,1,'MainRight','Limit.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：LimitEditDo
'作    用：执行修改权限
'参    数：
'==============================
Sub LimitEditDo()
	Dim Temp2,TempArr2
	Fk_Limit_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Limit_Name")))
	Fk_Limit_Set1=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Set1"))," ",""))&","
	Fk_Limit_Set2=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Set2"))," ",""))&","
	Temp=FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Content"))," ",""))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Limit_Name,1,50,0,"请输入权限名称！","权限名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Fk_Limit_Content=","
	TempArr=Split(Temp,",")
	For Each Temp In TempArr
		Fk_Limit_Content=Fk_Limit_Content&"Module"&Temp&","
	Next
	For Each Temp In TempArr
		Sqlstr="Select Fk_Module_LevelList,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Id=" & Temp
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Temp2=Rs("Fk_Module_LevelList")
			If Instr(Fk_Limit_Content,",Menu"&Rs("Fk_Module_Menu")&",")=0 Then
				Fk_Limit_Content=Fk_Limit_Content&"Menu"&Rs("Fk_Module_Menu")&","
			End if
		End If
		Rs.Close
		If Temp2<>"" Then
			TempArr2=Split(Temp2,",")
			For Each Temp2 In TempArr2
				If Temp2<>"" Then
					If Instr(Fk_Limit_Content,",Module"&Temp2&",")=0 And Instr(Fk_Limit_Content,",See"&Temp2&",")=0 Then
						Fk_Limit_Content=Fk_Limit_Content&"See"&Temp2&","
					End If
				End If
			Next
		End If
	Next
	Sqlstr="Select Fk_Limit_Id,Fk_Limit_Name,Fk_Limit_Set1,Fk_Limit_Set2,Fk_Limit_Content From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Limit_Name")=Fk_Limit_Name
		Rs("Fk_Limit_Set1")=Fk_Limit_Set1
		Rs("Fk_Limit_Set2")=Fk_Limit_Set2
		Rs("Fk_Limit_Content")=Fk_Limit_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("权限修改成功！")
	Else
		Response.Write("权限不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：LimitDelDo
'作    用：执行删除权限
'参    数：
'==============================
Sub LimitDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Admin_Id From [Fk_Admin] Where Fk_Admin_Limit=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该权限尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Limit_Id From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("权限删除成功！")
	Else
		Response.Write("权限不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->