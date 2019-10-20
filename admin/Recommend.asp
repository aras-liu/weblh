<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Admin/Recommend.asp
'文件用途：推荐类型管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System5",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_Recommend_Name

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call RecommendList() '推荐类型列表
	Case 2
		Call RecommendAddForm() '添加推荐类型表单
	Case 3
		Call RecommendAddDo() '执行添加推荐类型
	Case 4
		Call RecommendEditForm() '修改推荐类型表单
	Case 5
		Call RecommendEditDo() '执行修改推荐类型
	Case 6
		Call RecommendDelDo() '执行删除推荐类型
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：RecommendList()
'作    用：推荐类型列表
'参    数：
'==========================================
Sub RecommendList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Recommend.asp?Type=2');">添加新推荐类型</a></li>
    </ul>
</div>
<div id="ListTop">
    推荐类型管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">序号</td>
            <td align="center" class="ListTdTop">类型名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend] Order By Fk_Recommend_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td align="center"><%=Rs("Fk_Recommend_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Recommend.asp?Type=4&Id=<%=Rs("Fk_Recommend_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Recommend_Name")%>”，此操作不可逆！','Recommend.asp?Type=6&Id=<%=Rs("Fk_Recommend_Id")%>','MainRight','Recommend.asp?Type=1');">删除</a></td>
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
'函 数 名：RecommendAddForm()
'作    用：添加推荐类型表单
'参    数：
'==========================================
Sub RecommendAddForm()
%>
<form id="RecommendAdd" name="RecommendAdd" method="post" action="Recommend.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">添加新推荐类型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">类型名称：</td>
	        <td>&nbsp;<input name="Fk_Recommend_Name" type="text" class="Input" id="Fk_Recommend_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>推荐类型的名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
        <input type="submit" onclick="Sends('RecommendAdd','Recommend.asp?Type=3',0,'',0,1,'MainRight','Recommend.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：RecommendAddDo
'作    用：执行添加推荐类型
'参    数：
'==============================
Sub RecommendAddDo()
	Fk_Recommend_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Recommend_Name")))
	Call FKFun.ShowString(Fk_Recommend_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend] Where Fk_Recommend_Name='"&Fk_Recommend_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Recommend_Name")=Fk_Recommend_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("新推荐类型添加成功！")
	Else
		Response.Write("该推荐类型已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：RecommendEditForm()
'作    用：修改推荐类型表单
'参    数：
'==========================================
Sub RecommendEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Recommend_Name From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Recommend_Name=Rs("Fk_Recommend_Name")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此推荐类型，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="RecommendEdit" name="RecommendEdit" method="post" action="Recommend.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">修改推荐类型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">类型名称：</td>
	        <td>&nbsp;<input name="Fk_Recommend_Name" value="<%=Fk_Recommend_Name%>" type="text" class="Input" id="Fk_Recommend_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>推荐类型的名称，请输入1-50个字符（两个字符为一个汉字）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('RecommendEdit','Recommend.asp?Type=5',0,'',0,1,'MainRight','Recommend.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：RecommendEditDo
'作    用：执行修改推荐类型
'参    数：
'==============================
Sub RecommendEditDo()
	Fk_Recommend_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Recommend_Name")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Recommend_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Recommend_Id,Fk_Recommend_Name From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Recommend_Name")=Fk_Recommend_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("推荐类型修改成功！")
	Else
		Response.Write("推荐类型不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：RecommendDelDo
'作    用：执行删除推荐类型
'参    数：
'==============================
Sub RecommendDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Article_Id From [Fk_Article] Where Fk_Article_Recommend Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该推荐类型尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Product_Id From [Fk_Product] Where Fk_Product_Recommend Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该推荐类型尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Down_Id From [Fk_Down] Where Fk_Down_Recommend Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该推荐类型尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_Recommend_Id From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("推荐类型删除成功！")
	Else
		Response.Write("推荐类型不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->