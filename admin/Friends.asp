<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Amdin/Friends.asp
'文件用途：友情链接管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System2",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Friends_Name,Fk_Friends_About,Fk_Friends_Url,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_Friends_FriendsType

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FriendsList() '友情链接列表
	Case 2
		Call FriendsAddForm() '添加友情链接表单
	Case 3
		Call FriendsAddDo() '执行添加友情链接
	Case 4
		Call FriendsEditForm() '修改友情链接表单
	Case 5
		Call FriendsEditDo() '执行修改友情链接
	Case 6
		Call FriendsDelDo() '执行删除友情链接
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FriendsList()
'作    用：友情链接列表
'参    数：
'==========================================
Sub FriendsList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Friends.asp?Type=2');">添加新友情链接</a></li>
    </ul>
</div>
<div id="ListTop">
    友情链接管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">序号</td>
            <td align="center" class="ListTdTop">站点名称</td>
            <td align="center" class="ListTdTop">站点LOGO</td>
            <td align="center" class="ListTdTop">显示模式</td>
            <td align="center" class="ListTdTop">链接类型</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Friends_Id,Fk_Friends_Name,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_FriendsType_Name From [Fk_FriendsList] Order By Fk_Friends_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
			Select Case Rs("Fk_Friends_ShowType")
				Case 1
					Fk_Friends_ShowType="LOGO"
				Case 2
					Fk_Friends_ShowType="文字"
			End Select
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_Friends_Name")%></td>
            <td align="center"><%If Rs("Fk_Friends_Logo")<>"" Then%><img src="<%=Rs("Fk_Friends_Logo")%>" width="88" height="31" /><%Else%>无LOGO<%End If%></td>
            <td align="center"><%=Fk_Friends_ShowType%></td>
            <td align="center"><%=Rs("Fk_FriendsType_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Friends.asp?Type=4&Id=<%=Rs("Fk_Friends_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Friends_Name")%>”？此操作不可逆！','Friends.asp?Type=6&Id=<%=Rs("Fk_Friends_Id")%>','MainRight','Friends.asp?Type=1');">删除</a></td>
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
            <td height="30" colspan="6">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FriendsAddForm()
'作    用：添加友情链接表单
'参    数：
'==========================================
Sub FriendsAddForm()
%>
<form id="FriendsAdd" name="FriendsAdd" method="post" action="Friends.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:500px;">添加新友情链接[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">站点名称：</td>
	        <td>&nbsp;<input name="Fk_Friends_Name" type="text" class="Input" id="Fk_Friends_Name" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点地址：</td>
	        <td>&nbsp;<input name="Fk_Friends_Url" type="text" class="Input" id="Fk_Friends_Url" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点链接，请用http://开头，可输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点介绍：</td>
	        <td>&nbsp;<input name="Fk_Friends_About" type="text" class="Input" id="Fk_Friends_About" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点的简介，可留空，请输入0-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点LOGO：</td>
	        <td>&nbsp;<input name="Fk_Friends_Logo" type="text" class="Input" id="Fk_Friends_Logo" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点的LOGO，可留空，请输入0-255个字符，可以输入链接，也可以上传到空间。</p>"><img src="Images/help.jpg" /></span><br />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Friends_Logos" name="Fk_Friends_Logos" src="PicUpLoad.asp?Type=2&Form=FriendsAdd&Input=Fk_Friends_Logo"></iframe>
            </td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接类型：</td>
	        <td>&nbsp;<select name="Fk_Friends_FriendsType" class="Input" id="Fk_Friends_FriendsType">
<%
	Sqlstr="Select * From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_FriendsType_Id")%>"><%=Rs("Fk_FriendsType_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择友情链接类型。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">显示模式：</td>
	        <td>&nbsp;<input name="Fk_Friends_ShowType" class="Input" type="radio" id="Fk_Friends_ShowType" value="1" />
	        LOGO
            <input name="Fk_Friends_ShowType" type="radio" class="Input" id="Fk_Friends_ShowType" value="2" checked="checked" />
          文字&nbsp;&nbsp;<span class="qbox" title="<p>设置显示模式为文字或者LOGO。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
        <input type="submit" onclick="Sends('FriendsAdd','Friends.asp?Type=3',0,'',0,1,'MainRight','Friends.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsAddDo
'作    用：执行添加友情链接
'参    数：
'==============================
Sub FriendsAddDo()
	Fk_Friends_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Name")))
	Fk_Friends_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_About")))
	Fk_Friends_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Url")))
	Fk_Friends_Logo=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Logo")))
	Fk_Friends_ShowType=Trim(Request.Form("Fk_Friends_ShowType"))
	Fk_Friends_FriendsType=Trim(Request.Form("Fk_Friends_FriendsType"))
	Call FKFun.ShowString(Fk_Friends_Name,1,255,0,"请输入友情链接名称！","友情链接名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_About,1,255,2,"请输入友情链接介绍！","友情链接介绍不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Url,1,255,0,"请输入友情链接地址！","友情链接地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Logo,1,255,2,"请输入友情链接LOGO！","友情链接LOGO不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Friends_ShowType,"请选择友情链接显示类型！")
	Call FKFun.ShowNum(Fk_Friends_FriendsType,"请选择友情链接类型！")
	Sqlstr="Select Fk_Friends_Id,Fk_Friends_Name,Fk_Friends_About,Fk_Friends_Url,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_Friends_FriendsType From [Fk_Friends] Where Fk_Friends_Name='"&Fk_Friends_Name&"' And Fk_Friends_Url='"&Fk_Friends_Url&"' And Fk_Friends_ShowType="&Fk_Friends_ShowType&""
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Friends_Name")=Fk_Friends_Name
		Rs("Fk_Friends_About")=Fk_Friends_About
		Rs("Fk_Friends_Url")=Fk_Friends_Url
		Rs("Fk_Friends_Logo")=Fk_Friends_Logo
		Rs("Fk_Friends_ShowType")=Fk_Friends_ShowType
		Rs("Fk_Friends_FriendsType")=Fk_Friends_FriendsType
		Rs.Update()
		Application.UnLock()
		Response.Write("新友情链接添加成功！")
	Else
		Response.Write("该名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：FriendsEditForm()
'作    用：修改友情链接表单
'参    数：
'==========================================
Sub FriendsEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Friends_Name,Fk_Friends_About,Fk_Friends_Url,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_Friends_FriendsType From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Friends_Name=Rs("Fk_Friends_Name")
		Fk_Friends_About=Rs("Fk_Friends_About")
		Fk_Friends_Url=Rs("Fk_Friends_Url")
		Fk_Friends_Logo=Rs("Fk_Friends_Logo")
		Fk_Friends_ShowType=Rs("Fk_Friends_ShowType")
		Fk_Friends_FriendsType=Rs("Fk_Friends_FriendsType")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此友情链接，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="FriendsEdit" name="FriendsEdit" method="post" action="Friends.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:500px;">修改友情链接[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">站点名称：</td>
	        <td>&nbsp;<input name="Fk_Friends_Name" value="<%=Fk_Friends_Name%>" type="text" class="Input" id="Fk_Friends_Name" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点地址：</td>
	        <td>&nbsp;<input name="Fk_Friends_Url" value="<%=Fk_Friends_Url%>" type="text" class="Input" id="Fk_Friends_Url" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点链接，请用http://开头，可输入1-255个字符。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点介绍：</td>
	        <td>&nbsp;<input name="Fk_Friends_About" value="<%=Fk_Friends_About%>" type="text" class="Input" id="Fk_Friends_About" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点的简介，可留空，请输入0-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点LOGO：</td>
	        <td>&nbsp;<input name="Fk_Friends_Logo" value="<%=Fk_Friends_Logo%>" type="text" class="Input" id="Fk_Friends_Logo" size="35" />&nbsp;&nbsp;<span class="qbox" title="<p>对方站点的LOGO，可留空，请输入0-255个字符，可以输入链接，也可以上传到空间。</p>"><img src="Images/help.jpg" /></span><br />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Fk_Friends_Logos" name="Fk_Friends_Logos" src="PicUpLoad.asp?Type=2&Form=FriendsEdit&Input=Fk_Friends_Logo"></iframe></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接类型：</td>
	        <td>&nbsp;<select name="Fk_Friends_FriendsType" class="Input" id="Fk_Friends_FriendsType">
<%
	Sqlstr="Select * From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_FriendsType_Id")%>"<%=FKFun.BeSelect(Rs("Fk_FriendsType_Id"),Fk_Friends_FriendsType)%>><%=Rs("Fk_FriendsType_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select>&nbsp;&nbsp;<span class="qbox" title="<p>选择友情链接类型。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">显示模式：</td>
	        <td>&nbsp;<input name="Fk_Friends_ShowType" class="Input" type="radio" id="Fk_Friends_ShowType" value="1"<%=FKFun.BeCheck(Fk_Friends_ShowType,1)%> />LOGO
            <input type="radio" name="Fk_Friends_ShowType" class="Input" id="Fk_Friends_ShowType" value="2"<%=FKFun.BeCheck(Fk_Friends_ShowType,2)%> />文字&nbsp;&nbsp;<span class="qbox" title="<p>设置显示模式为文字或者LOGO。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('FriendsEdit','Friends.asp?Type=5',0,'',0,1,'MainRight','Friends.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsEditDo
'作    用：执行修改友情链接
'参    数：
'==============================
Sub FriendsEditDo()
	Fk_Friends_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Name")))
	Fk_Friends_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_About")))
	Fk_Friends_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Url")))
	Fk_Friends_Logo=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Logo")))
	Fk_Friends_ShowType=Trim(Request.Form("Fk_Friends_ShowType"))
	Fk_Friends_FriendsType=Trim(Request.Form("Fk_Friends_FriendsType"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Friends_Name,1,255,0,"请输入友情链接名称！","友情链接名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_About,1,255,2,"请输入友情链接介绍！","友情链接介绍不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Url,1,255,0,"请输入友情链接地址！","友情链接地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Logo,1,255,2,"请输入友情链接LOGO！","友情链接LOGO不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Friends_ShowType,"请选择友情链接显示类型！")
	Call FKFun.ShowNum(Fk_Friends_FriendsType,"请选择友情链接类型！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Friends_Id,Fk_Friends_Name,Fk_Friends_About,Fk_Friends_Url,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_Friends_FriendsType From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Friends_Name")=Fk_Friends_Name
		Rs("Fk_Friends_About")=Fk_Friends_About
		Rs("Fk_Friends_Url")=Fk_Friends_Url
		Rs("Fk_Friends_Logo")=Fk_Friends_Logo
		Rs("Fk_Friends_ShowType")=Fk_Friends_ShowType
		Rs("Fk_Friends_FriendsType")=Fk_Friends_FriendsType
		Rs.Update()
		Application.UnLock()
		Response.Write("友情链接修改成功！")
	Else
		Response.Write("友情链接不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：FriendsDelDo
'作    用：执行删除友情链接
'参    数：
'==============================
Sub FriendsDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Friends_Id From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("友情链接删除成功！")
	Else
		Response.Write("友情链接不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->