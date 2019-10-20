<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Vote.asp
'文件用途：在线投票管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System7",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket,Fk_Vote_Count

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call VoteList() '在线投票列表
	Case 2
		Call VoteAddForm() '添加在线投票表单
	Case 3
		Call VoteAddDo() '执行添加在线投票
	Case 4
		Call VoteEditForm() '修改在线投票表单
	Case 5
		Call VoteEditDo() '执行修改在线投票
	Case 6
		Call VoteDelDo() '执行删除在线投票
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：VoteList()
'作    用：在线投票列表
'参    数：
'==========================================
Sub VoteList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Vote.asp?Type=2');">添加新在线投票</a></li>
    </ul>
</div>
<div id="ListTop">
    在线投票管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Vote_Id,Fk_Vote_Name From [Fk_Vote] Order By Fk_Vote_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td>&nbsp;<%=Rs("Fk_Vote_Name")%>&nbsp;&nbsp;<a href="javascript:void(0);" onclick="window.clipboardData.setData('Text','<script type=\'text/javascript\' src=\'<%=SiteDir%>Plugin/Vote/Vote.asp?Id=<%=Rs("Fk_Vote_Id")%>\'></script>');alert('在线投票代码复制成功');">[复制代码]</a></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Vote.asp?Type=4&Id=<%=Rs("Fk_Vote_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Vote_Name")%>”，此操作不可逆！','Vote.asp?Type=6&Id=<%=Rs("Fk_Vote_Id")%>','MainRight','Vote.asp?Type=1');">删除</a></td>
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
'函 数 名：VoteAddForm()
'作    用：添加在线投票表单
'参    数：
'==========================================
Sub VoteAddForm()
%>
<form id="VoteAdd" name="VoteAdd" method="post" action="Vote.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新在线投票[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_Vote_Name" type="text" class="Input" id="Fk_Vote_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>在线投票名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">条目：</td>
            <td><p>&nbsp;</p>
            	<div id="t">&nbsp;&nbsp;添加投票条目</div>
                &nbsp;&nbsp;选项：<input name="Fk_Vote_Contents" type="text" class="Input" id="Fk_Vote_Contents" />
                票数：<input name="Fk_Vote_Tickets" type="text" class="Input" value="0" id="Fk_Vote_Tickets" />
                <input type="button" class="Button" onclick="VoteAdds($('#Fk_Vote_Contents').val(),$('#Fk_Vote_Tickets').val());" name="Adds" id="Adds" value="添 加" />&nbsp;&nbsp;<span class="qbox" title="<p>在线投票条目，理论可添加无数个，票数必须是数字。</p>"><img src="Images/help.jpg" /></span>
                <p>&nbsp;</p>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('VoteAdd','Vote.asp?Type=3',0,'',0,1,'MainRight','Vote.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：VoteAddDo
'作    用：执行添加在线投票
'参    数：
'==============================
Sub VoteAddDo()
	Fk_Vote_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Name")))
	Fk_Vote_Content=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Content"))),", ","<br />")
	Fk_Vote_Ticket=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Ticket"))),", ","|")
	Call FKFun.ShowString(Fk_Vote_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Vote_Content,1,5000,0,"请输入选项！","选项不能大于5000个字符！")
	TempArr=Split(Fk_Vote_Content,"<br />")
	For Each Temp In TempArr
		If Trim(Temp)="" Then
			Call FKFun.ShowErr("每个选项必须要有选项值！",2)
		End If
	Next
	TempArr=Split(Fk_Vote_Ticket,"|")
	For Each Temp In TempArr
		If Trim(Temp)="" Then
			Call FKFun.ShowErr("票数不能为空！",2)
		Else
			Call FKFun.ShowNum(Temp,"票数必须是数字！")
		End If
	Next
	Sqlstr="Select Fk_Vote_Id,Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket From [Fk_Vote] Where Fk_Vote_Name='"&Fk_Vote_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Vote_Name")=Fk_Vote_Name
		Rs("Fk_Vote_Content")=Fk_Vote_Content
		Rs("Fk_Vote_Ticket")=Fk_Vote_Ticket
		Rs.Update()
		Application.UnLock()
		Response.Write("新在线投票添加成功！")
	Else
		Response.Write("该在线投票已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：VoteEditForm()
'作    用：修改在线投票表单
'参    数：
'==========================================
Sub VoteEditForm()
	Dim TempArr2
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Vote_Name=FKFun.HTMLDncode(Rs("Fk_Vote_Name"))
		Fk_Vote_Content=Replace(Rs("Fk_Vote_Content"),"<br /> ","<br />")
		Fk_Vote_Ticket=Rs("Fk_Vote_Ticket")
		TempArr=Split(Fk_Vote_Content,"<br />")
		TempArr2=Split(Fk_Vote_Ticket,"|")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此投票，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="VoteEdit" name="VoteEdit" method="post" action="Vote.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改在线投票[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Vote_Name" value="<%=Fk_Vote_Name%>" type="text" class="Input" id="Fk_Vote_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>在线投票名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">条目：</td>
            <td><p>&nbsp;</p>
<%
	For i=0 To UBound(TempArr)
%>
            <div id="s<%=i%>">&nbsp;&nbsp;选项：<input name="Fk_Vote_Content" type="text" class="Input" value="<%=TempArr(i)%>" id="Fk_Vote_Content" />&nbsp;票数：<input name="Fk_Vote_Ticket" type="text" class="Input" value="<%=TempArr2(i)%>" id="Fk_Vote_Ticket" />&nbsp;<a href="javascript:void(0);" onclick="$('#s<%=i%>').remove();" title="删除">删除</a></div>
<%
	Next
%>
            	<div id="t">&nbsp;&nbsp;添加投票条目</div>
                &nbsp;&nbsp;选项：<input name="Fk_Vote_Contents" type="text" class="Input" id="Fk_Vote_Contents" />
                票数：<input name="Fk_Vote_Tickets" type="text" class="Input" value="0" id="Fk_Vote_Tickets" />
                <input type="button" class="Button" onclick="VoteAdds($('#Fk_Vote_Contents').val(),$('#Fk_Vote_Tickets').val());" name="Adds" id="Adds" value="添 加" />&nbsp;&nbsp;<span class="qbox" title="<p>在线投票条目，理论可添加无数个，票数必须是数字。</p>"><img src="Images/help.jpg" /></span>
                <p>&nbsp;</p>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('VoteEdit','Vote.asp?Type=5',0,'',0,1,'MainRight','Vote.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：VoteEditDo
'作    用：执行修改在线投票
'参    数：
'==============================
Sub VoteEditDo()
	Fk_Vote_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Name")))
	Fk_Vote_Content=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Content"))),", ","<br />")
	Fk_Vote_Ticket=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Ticket"))),", ","|")
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Vote_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Vote_Content,1,5000,0,"请输入选项！","选项不能大于5000个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	TempArr=Split(Fk_Vote_Content,"<br />")
	For Each Temp In TempArr
		If Trim(Temp)="" Then
			Call FKFun.ShowErr("每个选项必须要有选项值！",2)
		End If
	Next
	TempArr=Split(Fk_Vote_Ticket,"|")
	For Each Temp In TempArr
		If Trim(Temp)="" Then
			Call FKFun.ShowErr("票数不能为空！",2)
		Else
			Call FKFun.ShowNum(Temp,"票数必须是数字！")
		End If
	Next
	Sqlstr="Select Fk_Vote_Id,Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Vote_Name")=Fk_Vote_Name
		Rs("Fk_Vote_Content")=Fk_Vote_Content
		Rs("Fk_Vote_Ticket")=Fk_Vote_Ticket
		Rs.Update()
		Application.UnLock()
		Response.Write("在线投票修改成功！")
	Else
		Response.Write("在线投票不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：VoteDelDo
'作    用：执行删除在线投票
'参    数：
'==============================
Sub VoteDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Vote_Id From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("在线投票删除成功！")
	Else
		Response.Write("在线投票不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->