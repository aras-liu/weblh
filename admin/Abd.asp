<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Abd.asp
'文件用途：广告管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System3",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Ad_Name,Fk_Ad_Content,Fk_Ad_File

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call AdList() '广告列表
	Case 2
		Call AdAddForm() '添加广告表单
	Case 3
		Call AdAddDo() '执行添加广告
	Case 4
		Call AdEditForm() '修改广告表单
	Case 5
		Call AdEditDo() '执行修改广告
	Case 6
		Call AdDelDo() '执行删除广告
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：AdList()
'作    用：广告列表
'参    数：
'==========================================
Sub AdList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Abd.asp?Type=2');">添加新广告</a></li>
    </ul>
</div>
<div id="ListTop">
    广告管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">文件名</td>
            <td align="center" class="ListTdTop">JS文件</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Ad_Id,Fk_Ad_Name,Fk_Ad_File From [Fk_Ad] Order By Fk_Ad_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td>&nbsp;<%=Rs("Fk_Ad_Name")%>&nbsp;&nbsp;<a href="javascript:void(0);" onclick="copyToClipboard('<script type=\'text/javascript\' src=\'<%=SiteDir%>Js/My/<%=Rs("Fk_Ad_File")%>.js\'></script>');">[复制代码]</a></td>
            <td align="center"><%=Rs("Fk_Ad_File")%></td>
            <td align="center"><%If FKFso.IsFile("../Js/My/"&Rs("Fk_Ad_File")&".js") Then%>存在<%Else%>不存在<%End If%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Abd.asp?Type=4&Id=<%=Rs("Fk_Ad_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Ad_Name")%>”广告？此操作不可逆！','Abd.asp?Type=6&Id=<%=Rs("Fk_Ad_Id")%>','MainRight','Abd.asp?Type=1');">删除</a></td>
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
'函 数 名：AdAddForm()
'作    用：添加广告表单
'参    数：
'==========================================
Sub AdAddForm()
%>
<form id="AdAdd" name="AdAdd" method="post" action="Abd.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新广告[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Ad_Name" type="text" class="Input" id="Fk_Ad_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>广告名称，请输入广告的相关描述以便于管理。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">文件名：</td>
            <td>&nbsp;<input name="Fk_Ad_File" type="text" class="Input" id="Fk_Ad_File" />*必须是字母或数字&nbsp;&nbsp;<span class="qbox" title="<p>每个广告都会独立生成JS文件，这里设置JS的文件名，JS文件名由字母或数字组成。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">广告内容：</td>
            <td>&nbsp;<textarea name="Fk_Ad_Content" cols="60" rows="10" class="TextArea" id="Fk_Ad_Content"></textarea>&nbsp;&nbsp;<span class="qbox" title="<p>广告内容请直接输入JS代码，可输入1-5000个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('AdAdd','Abd.asp?Type=3',0,'',0,1,'MainRight','Abd.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：AdAddDo
'作    用：执行添加广告
'参    数：
'==============================
Sub AdAddDo()
	Fk_Ad_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Ad_Name")))
	Fk_Ad_Content=Trim(Request.Form("Fk_Ad_Content"))
	Fk_Ad_File=FKFun.HTMLEncode(Trim(Request.Form("Fk_Ad_File")))
	Call FKFun.ShowString(Fk_Ad_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Ad_File,1,50,0,"请输入文件名！","文件名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Ad_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Sqlstr="Select Fk_Ad_Id,Fk_Ad_Name,Fk_Ad_Content,Fk_Ad_File From [Fk_Ad] Where Fk_Ad_Name='"&Fk_Ad_Name&"' Or Fk_Ad_Content='"&Fk_Ad_File&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Ad_Name")=Fk_Ad_Name
		Rs("Fk_Ad_Content")=Fk_Ad_Content
		Rs("Fk_Ad_File")=Fk_Ad_File
		Rs.Update()
		Application.UnLock()
		Call FKFso.CreateFile(FileDir&"Js/My/"&Fk_Ad_File&".js",FKFun.HtmlToJs(Fk_Ad_Content))
		Response.Write("新广告添加成功！")
	Else
		Response.Write("广告名称或者文件名已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：AdEditForm()
'作    用：修改广告表单
'参    数：
'==========================================
Sub AdEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Ad_Name,Fk_Ad_Content,Fk_Ad_File From [Fk_Ad] Where Fk_Ad_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Ad_Name=FKFun.HTMLDncode(Rs("Fk_Ad_Name"))
		Fk_Ad_Content=Server.HTMLEncode(Rs("Fk_Ad_Content"))
		Fk_Ad_File=FKFun.HTMLDncode(Rs("Fk_Ad_File"))
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此广告，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="AdEdit" name="AdEdit" method="post" action="Abd.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改广告[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Ad_Name" value="<%=Fk_Ad_Name%>" type="text" class="Input" id="Fk_Ad_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>广告名称，请输入广告的相关描述以便于管理。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">文件名：</td>
            <td>&nbsp;<input name="Fk_Ad_File" value="<%=Fk_Ad_File%>" readonly="readonly" type="text" class="Input" id="Fk_Ad_File" />*必须是字母或数字&nbsp;&nbsp;<span class="qbox" title="<p>每个广告都会独立生成JS文件，这里设置JS的文件名，JS文件名由字母或数字组成。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td height="30" align="right">广告内容：</td>
            <td>&nbsp;<textarea name="Fk_Ad_Content" cols="60" rows="10" class="TextArea" id="Fk_Ad_Content"><%=Fk_Ad_Content%></textarea>&nbsp;&nbsp;<span class="qbox" title="<p>广告内容请直接输入JS代码，可输入1-5000个字符。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('AdEdit','Abd.asp?Type=5',0,'',0,1,'MainRight','Abd.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：AdEditDo
'作    用：执行修改广告
'参    数：
'==============================
Sub AdEditDo()
	Fk_Ad_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Ad_Name")))
	Fk_Ad_Content=Trim(Request.Form("Fk_Ad_Content"))
	Fk_Ad_File=FKFun.HTMLEncode(Trim(Request.Form("Fk_Ad_File")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Ad_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Ad_File,1,50,0,"请输入文件名！","文件名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Ad_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Ad_Id,Fk_Ad_Name,Fk_Ad_Content,Fk_Ad_File From [Fk_Ad] Where Fk_Ad_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Ad_Name")=Fk_Ad_Name
		Rs("Fk_Ad_Content")=Fk_Ad_Content
		Rs("Fk_Ad_File")=Fk_Ad_File
		Rs.Update()
		Application.UnLock()
		Call FKFso.CreateFile(FileDir&"Js/My/"&Fk_Ad_File&".js",FKFun.HtmlToJs(Fk_Ad_Content))
		Response.Write("广告修改成功！")
	Else
		Response.Write("广告不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：AdDelDo
'作    用：执行删除广告
'参    数：
'==============================
Sub AdDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Ad] Where Fk_Ad_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("广告删除成功！")
	Else
		Response.Write("广告不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->