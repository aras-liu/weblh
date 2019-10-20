<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Infos.asp
'文件用途：独立信息管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System8",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Info_Name,Fk_Info_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call InfoList() '独立信息列表
	Case 2
		Call InfoAddForm() '添加独立信息表单
	Case 3
		Call InfoAddDo() '执行添加独立信息
	Case 4
		Call InfoEditForm() '修改独立信息表单
	Case 5
		Call InfoEditDo() '执行修改独立信息
	Case 6
		Call InfoDelDo() '执行删除独立信息
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：InfoList()
'作    用：独立信息列表
'参    数：
'==========================================
Sub InfoList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Infos.asp?Type=2');">添加新独立信息</a></li>
    </ul>
</div>
<div id="ListTop">
    独立信息管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">标签</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Info_Id,Fk_Info_Name From [Fk_Info] Order By Fk_Info_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td>&nbsp;<%=Rs("Fk_Info_Name")%></td>
            <td align="center">{$Info(<%=Rs("Fk_Info_Id")%>)$}</td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Infos.asp?Type=4&Id=<%=Rs("Fk_Info_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Info_Name")%>”？此操作不可逆！','Infos.asp?Type=6&Id=<%=Rs("Fk_Info_Id")%>','MainRight','Infos.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="5" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="5">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：InfoAddForm()
'作    用：添加独立信息表单
'参    数：
'==========================================
Sub InfoAddForm()
%>
<form id="InfoAdd" name="InfoAdd" method="post" action="Infos.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">添加新独立信息[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="8%" height="25" align="right">标题：</td>
	        <td width="92%">&nbsp;<input name="Fk_Info_Name" type="text" class="Input" id="Fk_Info_Name" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>独立信息标题，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">内容&nbsp;&nbsp;<span class="qbox" title="<p>独立信息内容，请输入1-50000个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td><textarea name="Fk_Info_Content" style="width:100%;" rows="20" class="<%=EditorClass%>" id="Fk_Info_Content"></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:880px;">
        <input type="submit" onclick="Sends('InfoAdd','Infos.asp?Type=3',0,'',0,1,'MainRight','Infos.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：InfoAddDo
'作    用：执行添加独立信息
'参    数：
'==============================
Sub InfoAddDo()
	Fk_Info_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Info_Name")))
	Fk_Info_Content=Request.Form("Fk_Info_Content")
	Call FKFun.ShowString(Fk_Info_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Info_Content,1,50000,0,"请输入内容！","内容不能大于50000个字符！")
	Sqlstr="Select Fk_Info_Id,Fk_Info_Name,Fk_Info_Content From [Fk_Info] Where Fk_Info_Name='"&Fk_Info_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Info_Name")=Fk_Info_Name
		Rs("Fk_Info_Content")=Fk_Info_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新独立信息添加成功！")
	Else
		Response.Write("该独立信息已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：InfoEditForm()
'作    用：修改独立信息表单
'参    数：
'==========================================
Sub InfoEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Info_Name,Fk_Info_Content From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Info_Name=FKFun.HTMLDncode(Rs("Fk_Info_Name"))
		Fk_Info_Content=Rs("Fk_Info_Content")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此独立信息，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="InfoEdit" name="InfoEdit" method="post" action="Infos.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:900px;">修改独立信息[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:900px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="8%" height="25" align="right">标题：</td>
	        <td width="92%">&nbsp;<input name="Fk_Info_Name" type="text" class="Input" id="Fk_Info_Name" value="<%=Fk_Info_Name%>" size="40" />&nbsp;&nbsp;<span class="qbox" title="<p>独立信息标题，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">内容&nbsp;&nbsp;<span class="qbox" title="<p>独立信息内容，请输入1-50000个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span>：</td>
            <td><textarea name="Fk_Info_Content" style="width:100%;" rows="20" class="<%=EditorClass%>" id="Fk_Info_Content"><%=Fk_Info_Content%></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:880px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('InfoEdit','Infos.asp?Type=5',0,'',0,1,'MainRight','Infos.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：InfoEditDo
'作    用：执行修改独立信息
'参    数：
'==============================
Sub InfoEditDo()
	Fk_Info_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Info_Name")))
	Fk_Info_Content=Request.Form("Fk_Info_Content")
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Info_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Info_Content,1,50000,0,"请输入内容！","内容不能大于50000个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Info_Id,Fk_Info_Name,Fk_Info_Content From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Info_Name")=Fk_Info_Name
		Rs("Fk_Info_Content")=Fk_Info_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("独立信息修改成功！")
	Else
		Response.Write("独立信息不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：InfoDelDo
'作    用：执行删除独立信息
'参    数：
'==============================
Sub InfoDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Info_Id From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("独立信息删除成功！")
	Else
		Response.Write("独立信息不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->