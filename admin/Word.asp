<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/Word.asp
'文件用途：站内关键字管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System4",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Word_Name,Fk_Word_Url,Fk_Word_RNum

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WordList() '站内关键字列表
	Case 2
		Call WordAddForm() '添加站内关键字表单
	Case 3
		Call WordAddDo() '执行添加站内关键字
	Case 4
		Call WordEditForm() '修改站内关键字表单
	Case 5
		Call WordEditDo() '执行修改站内关键字
	Case 6
		Call WordDelDo() '执行删除站内关键字
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WordList()
'作    用：站内关键字列表
'参    数：
'==========================================
Sub WordList()
	Session("NowPage")=FkFun.GetNowUrl()
	PageNow=FKFun.GetNumeric("Page",1)
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Word.asp?Type=2');">添加新关键字</a></li>
    </ul>
</div>
<div id="ListTop">
    站内关键字管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">关键字</td>
            <td align="center" class="ListTdTop">链接</td>
            <td align="center" class="ListTdTop">替换次数</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_Word] Order By Fk_Word_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
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
			If Rs("Fk_Word_RNum")=0 Then
				Fk_Word_RNum="替换所有"
			Else
				Fk_Word_RNum="替换"&Rs("Fk_Word_RNum")&"次"
			End If
%>
        <tr>
            <td height="20" align="center"><%=j%></td>
            <td align="left">&nbsp;&nbsp;<%=Rs("Fk_Word_Name")%></td>
            <td align="left">&nbsp;&nbsp;<%=Rs("Fk_Word_Url")%></td>
            <td align="center"><%=Fk_Word_RNum%>次</td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Word.asp?Type=4&Id=<%=Rs("Fk_Word_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Word_Name")%>”，此操作不可逆！','Word.asp?Type=6&Id=<%=Rs("Fk_Word_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
			j=j+1
		Wend
%>
        <tr>
            <td height="30" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("Word.asp?Type=1&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="5" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：WordAddForm()
'作    用：添加站内关键字表单
'参    数：
'==========================================
Sub WordAddForm()
%>
<form id="WordAdd" name="WordAdd" method="post" action="Word.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">添加新站内关键字[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">关键字：</td>
	        <td>&nbsp;<input name="Fk_Word_Name" type="text" class="Input" id="Fk_Word_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接：</td>
	        <td>&nbsp;<input name="Fk_Word_Url" type="text" class="Input" id="Fk_Word_Url" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入链接地址，支持站内链接，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">替换次数：</td>
	        <td>&nbsp;<input name="Fk_Word_RNum" type="text" class="Input" id="Fk_Word_RNum" value="1" /> *0为不限制&nbsp;&nbsp;<span class="qbox" title="<p>必须输入数字，如输入0则替换全部，输入0以上数字则替换相应个数。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
        <input type="submit" onclick="Sends('WordAdd','Word.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WordAddDo
'作    用：执行添加站内关键字
'参    数：
'==============================
Sub WordAddDo()
	Fk_Word_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Name")))
	Fk_Word_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Url")))
	Fk_Word_RNum=Trim(Request.Form("Fk_Word_RNum"))
	Call FKFun.ShowString(Fk_Word_Name,1,50,0,"请输入关键字！","关键字不能大于50个字符！")
	Call FKFun.ShowString(Fk_Word_Url,1,255,0,"请输入链接！","链接不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Word_RNum,"请输入替换次数！")
	Sqlstr="Select Fk_Word_Id,Fk_Word_Name,Fk_Word_Url,Fk_Word_RNum From [Fk_Word] Where Fk_Word_Name='"&Fk_Word_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Word_Name")=Fk_Word_Name
		Rs("Fk_Word_Url")=Fk_Word_Url
		Rs("Fk_Word_RNum")=Fk_Word_RNum
		Rs.Update()
		Application.UnLock()
		Response.Write("新站内关键字添加成功！")
	Else
		Response.Write("该关键字已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：WordEditForm()
'作    用：修改站内关键字表单
'参    数：
'==========================================
Sub WordEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Word_Name,Fk_Word_Url,Fk_Word_RNum From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Word_Name=FKFun.HTMLDncode(Rs("Fk_Word_Name"))
		Fk_Word_Url=FKFun.HTMLDncode(Rs("Fk_Word_Url"))
		Fk_Word_RNum=Rs("Fk_Word_RNum")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此关键字，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="WordEdit" name="WordEdit" method="post" action="Word.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">修改站内关键字[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">关键字：</td>
	        <td>&nbsp;<input name="Fk_Word_Name" value="<%=Fk_Word_Name%>" type="text" class="Input" id="Fk_Word_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接：</td>
	        <td>&nbsp;<input name="Fk_Word_Url" value="<%=Fk_Word_Url%>" type="text" class="Input" id="Fk_Word_Url" />&nbsp;&nbsp;<span class="qbox" title="<p>请输入链接地址，支持站内链接，请输入1-255个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">替换次数：</td>
	        <td>&nbsp;<input name="Fk_Word_RNum" value="<%=Fk_Word_RNum%>" type="text" class="Input" id="Fk_Word_RNum" /> *0为不限制&nbsp;&nbsp;<span class="qbox" title="<p>必须输入数字，如输入0则替换全部，输入0以上数字则替换相应个数。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('WordEdit','Word.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WordEditDo
'作    用：执行修改站内关键字
'参    数：
'==============================
Sub WordEditDo()
	Fk_Word_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Name")))
	Fk_Word_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Url")))
	Fk_Word_RNum=Trim(Request.Form("Fk_Word_RNum"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Word_Name,1,50,0,"请输入关键字！","关键字不能大于50个字符！")
	Call FKFun.ShowString(Fk_Word_Url,1,255,0,"请输入链接！","链接不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Word_RNum,"请输入替换次数！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Word_Id,Fk_Word_Name,Fk_Word_Url,Fk_Word_RNum From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Word_Name")=Fk_Word_Name
		Rs("Fk_Word_Url")=Fk_Word_Url
		Rs("Fk_Word_RNum")=Fk_Word_RNum
		Rs.Update()
		Application.UnLock()
		Response.Write("站内关键字修改成功！")
	Else
		Response.Write("站内关键字不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：WordDelDo
'作    用：执行删除站内关键字
'参    数：
'==============================
Sub WordDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Word_Id From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("站内关键字删除成功！")
	Else
		Response.Write("站内关键字不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->