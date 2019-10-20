<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/FriendsType.asp
'文件用途：友情链接类型管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System2",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_FriendsType_Name

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FriendsTypeList() '友情链接类型列表
	Case 2
		Call FriendsTypeAddForm() '添加友情链接类型表单
	Case 3
		Call FriendsTypeAddDo() '执行添加友情链接类型
	Case 4
		Call FriendsTypeEditForm() '修改友情链接类型表单
	Case 5
		Call FriendsTypeEditDo() '执行修改友情链接类型
	Case 6
		Call FriendsTypeDelDo() '执行删除友情链接类型
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FriendsTypeList()
'作    用：友情链接类型列表
'参    数：
'==========================================
Sub FriendsTypeList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('FriendsType.asp?Type=2');">添加新友情链接类型</a></li>
    </ul>
</div>
<div id="ListTop">
    友情链接类型管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">序号</td>
            <td align="center" class="ListTdTop">类型名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_FriendsType_Id,Fk_FriendsType_Name From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_FriendsType_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('FriendsType.asp?Type=4&Id=<%=Rs("Fk_FriendsType_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_FriendsType_Name")%>”？此操作不可逆！','FriendsType.asp?Type=6&Id=<%=Rs("Fk_FriendsType_Id")%>','MainRight','FriendsType.asp?Type=1');">删除</a></td>
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
'函 数 名：FriendsTypeAddForm()
'作    用：添加友情链接类型表单
'参    数：
'==========================================
Sub FriendsTypeAddForm()
%>
<form id="FriendsTypeAdd" name="FriendsTypeAdd" method="post" action="FriendsType.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">添加新友情链接类型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">类型名称：</td>
	        <td>&nbsp;<input name="Fk_FriendsType_Name" type="text" class="Input" id="Fk_FriendsType_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>友情链接类型名称，不能为空或重复，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
        <input type="submit" onclick="Sends('FriendsTypeAdd','FriendsType.asp?Type=3',0,'',0,1,'MainRight','FriendsType.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsTypeAddDo
'作    用：执行添加友情链接类型
'参    数：
'==============================
Sub FriendsTypeAddDo()
	Fk_FriendsType_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_FriendsType_Name")))
	Call FKFun.ShowString(Fk_FriendsType_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Sqlstr="Select Fk_FriendsType_Id,Fk_FriendsType_Name From [Fk_FriendsType] Where Fk_FriendsType_Name='"&Fk_FriendsType_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_FriendsType_Name")=Fk_FriendsType_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("新友情链接类型添加成功！")
	Else
		Response.Write("该类型名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：FriendsTypeEditForm()
'作    用：修改友情链接类型表单
'参    数：
'==========================================
Sub FriendsTypeEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_FriendsType_Name From [Fk_FriendsType] Where Fk_FriendsType_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_FriendsType_Name=Rs("Fk_FriendsType_Name")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此友情链接类型，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="FriendsTypeEdit" name="FriendsTypeEdit" method="post" action="FriendsType.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">修改友情链接类型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">类型名称：</td>
	        <td>&nbsp;<input name="Fk_FriendsType_Name" value="<%=Fk_FriendsType_Name%>" type="text" class="Input" id="Fk_FriendsType_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>友情链接类型名称，不能为空或重复，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('FriendsTypeEdit','FriendsType.asp?Type=5',0,'',0,1,'MainRight','FriendsType.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsTypeEditDo
'作    用：执行修改友情链接类型
'参    数：
'==============================
Sub FriendsTypeEditDo()
	Fk_FriendsType_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_FriendsType_Name")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_FriendsType_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_FriendsType_Id,Fk_FriendsType_Name From [Fk_FriendsType] Where Fk_FriendsType_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_FriendsType_Name")=Fk_FriendsType_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("友情链接类型修改成功！")
	Else
		Response.Write("友情链接类型不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：FriendsTypeDelDo
'作    用：执行删除友情链接类型
'参    数：
'==============================
Sub FriendsTypeDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Friends_Id From [Fk_Friends] Where Fk_Friends_FriendsType=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该友情链接类型尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_FriendsType_Id From [Fk_FriendsType] Where Fk_FriendsType_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("友情链接类型删除成功！")
	Else
		Response.Write("友情链接类型不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->