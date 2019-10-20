<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：Admin/GBook.asp
'文件用途：留言管理拉取页面
'版权所有：方卡在线
'==========================================

'定义页面变量
Dim Fk_GBook_Content,Fk_GBook_Ip,Fk_GBook_Time,Fk_GBook_ReContent,Fk_GBook_ReAdmin,Fk_GBook_ReIp,Fk_GBook_ReTime
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_GModel,Fk_GModel_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GBookList() '留言列表
	Case 2
		Call GBookReForm() '回复留言表单
	Case 3
		Call GBookReDo() '执行回复留言
	Case 4
		Call GBookDelDo() '执行删除留言
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：GBookList()
'作    用：留言列表
'参    数：
'==========================================
Sub GBookList()
	Dim Temp2,TempArr2,TempArr3,ShowModel,ii
	Session("NowPage")=FkFun.GetNowUrl()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限！",2)
	End If
	PageNow=FKFun.GetNumeric("Page",1)
	Sqlstr="Select Fk_Module_Name,Fk_Module_Menu,Fk_Module_GModel From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_GModel=Rs("Fk_Module_GModel")
	Else
		Rs.Close
		Call FKFun.ShowErr("模块不存在！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_GModel_Content From [Fk_GModel] Where Fk_GModel_Id=" & Fk_Module_GModel
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_GModel_Content=Rs("Fk_GModel_Content")
		TempArr=Split(Fk_GModel_Content,"|-_-|Fangka|-_-|")
	Else
		Rs.Close
		Call FKFun.ShowErr("留言模型不存在！",2)
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','GBook.asp?Type=1&ModuleId=<%=Fk_Module_Id%>');">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    “<%=Fk_Module_Name%>”模块&nbsp;&nbsp;
    &nbsp;&nbsp;快速通道：<select name="D1" id="D1" onChange="eval(this.options[this.selectedIndex].value);" class="Input">
      <option value="alert('请选择模块');">请选择模块</option>
<%
Call FKAdmin.GetModuleList(0,Fk_Module_Menu,0,Fk_Module_Id,"")
%>
</select>

</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop"><%=Split(TempArr(0),"|-^-|Fangka|-^-|")(0)%></td>
<%
For i=1 To UBound(TempArr)
	If Instr(TempArr(i),"|-^-|Fangka|-^-|") Then
		TempArr3=Split(TempArr(i),"|-^-|Fangka|-^-|")
		If UBound(TempArr3)=4 Then
%>
            <td align="center" class="ListTdTop"><%=TempArr3(0)%></td>
<%
		End If
	End If
Next
%>
            <td align="center" class="ListTdTop">IP</td>
            <td align="center" class="ListTdTop">时间</td>
            <td align="center" class="ListTdTop">处理时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_GBook_Id,Fk_GBook_Content,Fk_GBook_Ip,Fk_GBook_Time,Fk_GBook_ReIp,Fk_GBook_ReTime From [Fk_GBook] Where Fk_GBook_Module="&Fk_Module_Id&" Order By Fk_GBook_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
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
			Fk_GBook_Content=Split(Rs("Fk_GBook_Content"),"|-_-|Fangka|-_-|")
%>
        <tr>
            <td height="20" align="center"><%=j%></td>
            <td align="left">&nbsp;&nbsp;<%=Split(Fk_GBook_Content(0),"|-^-|Fangka|-^-|")(1)%></td>
<%
For ii=1 To UBound(TempArr)
	If Instr(TempArr(ii),"|-^-|Fangka|-^-|") Then
		TempArr3=Split(TempArr(ii),"|-^-|Fangka|-^-|")
		If UBound(TempArr3)=4 Then
			Temp2=""
			For Each Temp In Fk_GBook_Content
				If Split(Temp,"|-^-|Fangka|-^-|")(2)=TempArr3(3) Then
					Temp2=Split(Fk_GBook_Content(ii),"|-^-|Fangka|-^-|")(1)
					Exit For
				End If
			Next
%>
            <td align="left">&nbsp;&nbsp;<%=Temp2%></td>
<%
		End If
	End If
Next
%>
            <td align="center"><%=Rs("Fk_GBook_Ip")%></td>
            <td align="center"><%=Rs("Fk_GBook_Time")%></td>
            <td align="center">&nbsp;<%=Rs("Fk_GBook_ReTime")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('GBook.asp?Type=2&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_GBook_Id")%>');">处理</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除本条记录？此操作不可逆！','GBook.asp?Type=4&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_GBook_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
			j=j+1
		Wend
%>
        <tr>
            <td height="30" colspan="111">&nbsp;<%Call FKFun.ShowPageCode("GBook.asp?Type=1&ModuleId="&Fk_Module_Id&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="111" align="center">暂无记录</td>
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
'函 数 名：GBookReForm()
'作    用：回复留言表单
'参    数：
'==========================================
Sub GBookReForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_GBook_Content,Fk_GBook_Ip,Fk_GBook_Time,Fk_GBook_ReContent,Fk_GBook_Module From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_GBook_Content=Split(Rs("Fk_GBook_Content"),"|-_-|Fangka|-_-|")
		Fk_GBook_Ip=Rs("Fk_GBook_Ip")
		Fk_GBook_Time=Rs("Fk_GBook_Time")
		Fk_GBook_ReContent=Rs("Fk_GBook_ReContent")
		Fk_Module_Id=Rs("Fk_GBook_Module")
	Else
		Call FKFun.ShowErr("未找到留言，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
	If Not FkAdmin.AdminCheck(4,"Module"&Fk_Module_Id,Request.Cookies("FkAdminLimit3")) Then
		Call FKFun.ShowErr("无权限，请按键盘上的ESC键退出操作！",1)
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Id=Rs("Fk_Module_Id")
	Else
		Call FKFun.ShowErr("未找到模块，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="GBookRe" name="GBookRe" method="post" action="GBook.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">“<%=Fk_Module_Name%>”记录处理[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
<%
	For Each Temp In Fk_GBook_Content
%>
    <tr>
        <td height="28" align="right"><%=Split(Temp,"|-^-|Fangka|-^-|")(0)%>：</td>
        <td>&nbsp;<%=Split(Temp,"|-^-|Fangka|-^-|")(1)%></td>
    </tr>
<%
	Next
%>
    <tr>
        <td height="28" align="right">IP：</td>
        <td>&nbsp;<%=Fk_GBook_Ip%></td>
    </tr>
    <tr>
        <td height="28" align="right">时间：</td>
        <td>&nbsp;<%=Fk_GBook_Time%></td>
    </tr>
    <tr>
        <td height="28" align="right">回复/处理情况：</td>
        <td>&nbsp;<textarea name="Fk_GBook_ReContent" cols="40" rows="5" class="TextArea" id="Fk_GBook_ReContent"><%=Fk_GBook_ReContent%></textarea></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:580px;">
        <input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('GBookRe','GBook.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="Enter" id="Enter" value="处 理" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：GBookReDo
'作    用：执行回复留言
'参    数：
'==============================
Sub GBookReDo()
	Fk_GBook_ReContent=FKFun.HTMLEncode(Request.Form("Fk_GBook_ReContent"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_GBook_ReContent,1,1,1,"请输入回复/处理情况内容，不少于1个字符！","")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_GBook_Module,Fk_GBook_ReContent,Fk_GBook_ReAdmin,Fk_GBook_ReIp,Fk_GBook_ReTime From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_GBook_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		Application.Lock()
		Rs("Fk_GBook_ReContent")=Fk_GBook_ReContent
		Rs("Fk_GBook_ReAdmin")=Request.Cookies("FkAdminId")
		Rs("Fk_GBook_ReIp")=Request.ServerVariables("REMOTE_ADDR")
		Rs("Fk_GBook_ReTime")=Now()
		Rs.Update()
		Application.UnLock()
		Response.Write("处理成功！")
	Else
		Response.Write("记录不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：GBookDelDo
'作    用：执行删除留言
'参    数：
'==============================
Sub GBookDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_GBook_Id,Fk_GBook_Module From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If Not FkAdmin.AdminCheck(4,"Module"&Rs("Fk_GBook_Module"),Request.Cookies("FkAdminLimit3")) Then
			Rs.Close
			Call FKFun.ShowErr("您无此权限！",2)
		End If
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("删除成功！")
	Else
		Response.Write("记录不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->