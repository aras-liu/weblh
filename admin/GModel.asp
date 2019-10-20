<!--#Include File="AdminCheck.asp"--><%
'==========================================
'文 件 名：GModel.asp
'文件用途：留言模型管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"System11",Request.Cookies("FkAdminLimit2"))

'定义页面变量
Dim Fk_GModel_Name,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MaxStr,Fk_GModel_MinStr,Fk_GModel_Content,Fk_GModel_GoUrl

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GModelList() '留言模型列表
	Case 2
		Call GModelAddForm() '添加留言模型表单
	Case 3
		Call GModelAddDo() '执行添加留言模型
	Case 4
		Call GModelEditForm() '修改留言模型表单
	Case 5
		Call GModelEditDo() '执行修改留言模型
	Case 6
		Call GModelDelDo() '执行删除留言模型
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：GModelList()
'作    用：留言模型列表
'参    数：
'==========================================
Sub GModelList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('GModel.asp?Type=2');">添加新留言模型</a></li>
    </ul>
</div>
<div id="ListTop">
    留言模型管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select Fk_GModel_Id,Fk_GModel_Name From [Fk_GModel] Order By Fk_GModel_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		i=1
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=i%></td>
            <td>&nbsp;<%=Rs("Fk_GModel_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('GModel.asp?Type=4&Id=<%=Rs("Fk_GModel_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_GModel_Name")%>”，此操作不可逆！','GModel.asp?Type=6&Id=<%=Rs("Fk_GModel_Id")%>','MainRight','GModel.asp?Type=1');">删除</a></td>
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
'函 数 名：GModelAddForm()
'作    用：添加留言模型表单
'参    数：
'==========================================
Sub GModelAddForm()
%>
<form id="GModelAdd" name="GModelAdd" method="post" action="GModel.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新留言模型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Name" type="text" class="Input" id="Fk_GModel_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>留言模型名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">成功留言转向：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_GoUrl" value="" type="text" class="Input" id="Fk_GModel_GoUrl" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功后转向的链接，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">留言成功提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Succeed" value="留言成功，我们会尽快回复！" type="text" class="Input" id="Fk_GModel_Succeed" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">重复留言提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Repeat" value="请勿重复留言！" type="text" class="Input" id="Fk_GModel_Repeat" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>重复留言提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">反垃圾留言：</td>
            <td>&nbsp;&nbsp;&nbsp;<input type="radio" name="Fk_GModel_NoTrash" class="Input" id="Fk_GModel_NoTrash" value="0" />关闭
            <input name="Fk_GModel_NoTrash" class="Input" type="radio" id="Fk_GModel_NoTrash" value="1" checked="checked" />开启
            <input name="Fk_GModel_NoTrash" class="Input" type="radio" id="Fk_GModel_NoTrash" value="2" />部分开启&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，对留言进行中文字数占有量和链接个数进行判断，如果中文字数过低或者链接过多（超过2个），则认为是垃圾留言，如果英文站，请选择关闭或者部分开启，部分开启是不判断中文字数，只判断链接个数。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    <tr>
	        <td height="25" align="right">垃圾留言提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_TrashPint" value="请勿发垃圾信息！" type="text" class="Input" id="Fk_GModel_TrashPint" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>垃圾留言提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">字符串过短提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_MinStr" value="请输入{-留言条目-}！" type="text" class="Input" id="Fk_GModel_MinStr" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>字符串过短提示，{-留言条目-}:留言条目，{-留言长度-}：限制长度，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">字符串过长提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_MaxStr" value="{-留言条目-}不能大于{-留言长度-}个字符！" type="text" class="Input" id="Fk_GModel_MaxStr" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>字符串过长提示，{-留言条目-}:留言条目，{-留言长度-}：限制长度，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">留言条目：</td>
            <td><p>&nbsp;&nbsp;注：留言标题不可删除，但可以改为其他名字用于前台提示，另外表单标识不可重复</p>
            	<div class="gModel">&nbsp;&nbsp;名称：<input name="Fk_GModel_Content" type="text" class="Input" id="Fk_GModel_Content" value="留言标题" /><br />&nbsp;&nbsp;最小字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="1" />&nbsp;最大字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="50" /><br />&nbsp;&nbsp;表单标识：<input name="Fk_GModel_Content" type="text" class="Input" value="GBook_Title" id="Fk_GModel_Content" /><input type="hidden" name="Fk_GModel_Content" id="Fk_GModel_Content" value="|-_-|Fangka|-_-|" /></div>
            	<div id="t">&nbsp;&nbsp;添加留言条目</div>
                &nbsp;&nbsp;名称：<input name="model1" type="text" class="Input" id="model1" /><br />
                &nbsp;&nbsp;最小字符数：<input name="model2" type="text" class="Input" value="0" size="5" id="model2" />
                最大字符数：<input name="model3" type="text" class="Input" value="50" size="5" id="model3" /><br />
                &nbsp;&nbsp;表单标识：<input name="model4" type="text" class="Input" value="GBook_" id="model4" />&nbsp;<input type="checkbox" name="model5" class="Input" value="Show" id="model5" />后台列表显示
                <input type="button" class="Button" onclick="GModelAdds($('#model1').val(),$('#model2').val(),$('#model3').val(),$('#model4').val(),$('#model5:checked').val());" name="Adds" id="Adds" value="添 加" />&nbsp;&nbsp;<span class="qbox" title="<p>留言模型条目，理论可添加无数个，票数必须是数字。</p>"><img src="Images/help.jpg" /></span>
                <p>&nbsp;</p>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('GModelAdd','GModel.asp?Type=3',0,'',0,1,'MainRight','GModel.asp?Type=1');" class="Button" name="Enter" id="Enter" value="添 加" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：GModelAddDo
'作    用：执行添加留言模型
'参    数：
'==============================
Sub GModelAddDo()
	Dim TempArr2
	Fk_GModel_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Name")))
	Fk_GModel_GoUrl=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_GoUrl")))
	Fk_GModel_Succeed=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Succeed")))
	Fk_GModel_Repeat=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Repeat")))
	Fk_GModel_TrashPint=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_TrashPint")))
	Fk_GModel_MaxStr=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_MaxStr")))
	Fk_GModel_MinStr=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_MinStr")))
	Fk_GModel_Content=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Content"))),", ","|-^-|Fangka|-^-|")
	Fk_GModel_Content=Replace(Fk_GModel_Content,"|-^-|Fangka|-^-||-_-|Fangka|-_-|","|-_-|Fangka|-_-|")
	Fk_GModel_Content=Replace(Fk_GModel_Content,"|-_-|Fangka|-_-||-^-|Fangka|-^-|","|-_-|Fangka|-_-|")
	Fk_GModel_NoTrash=Trim(Request.Form("Fk_GModel_NoTrash"))
	Call FKFun.ShowString(Fk_GModel_Name,1,50,0,"请输入模型名称！","模型名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_GModel_GoUrl,0,255,0,"请输入留言成功转向链接！","留言成功转向链接不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Succeed,1,255,0,"请输入留言成功提示！","留言成功提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Repeat,1,255,0,"请输入重复留言提示！","重复留言不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_TrashPint,1,255,0,"请输入垃圾留言提示！","垃圾留言提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_MaxStr,1,255,0,"请输入字符过长提示！","字符过长提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_MinStr,1,255,0,"请输入字符过短提示！","字符过短提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Content,20,50000,0,"请录入留言条目！","留言条目不能大于50000个字符！")
	Call FKFun.ShowNum(Fk_GModel_NoTrash,"请选择反垃圾留言是否开启选项！")
	TempArr=Split(Fk_GModel_Content,"|-_-|Fangka|-_-|")
	For Each Temp In TempArr
		If Trim(Temp)<>"" Then
			TempArr2=Split(Temp,"|-^-|Fangka|-^-|")
			Call FKFun.ShowString(TempArr2(0),1,50,0,"留言条目名称不能留空！","留言条目名称不能大于50个字符！")
			Call FKFun.ShowString(TempArr2(3),1,50,0,"留言条目标识不能留空！","留言条目标识不能大于50个字符！")
			Call FKFun.ShowNum(TempArr2(1),"最小字符数必须是数字！")
			Call FKFun.ShowNum(TempArr2(2),"最大字符数必须是数字！")
		End If
	Next
	Sqlstr="Select Fk_GModel_Id,Fk_GModel_Name,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MaxStr,Fk_GModel_MinStr,Fk_GModel_Content,Fk_GModel_GoUrl From [Fk_GModel] Where Fk_GModel_Name='"&Fk_GModel_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_GModel_Name")=Fk_GModel_Name
		Rs("Fk_GModel_GoUrl")=Fk_GModel_GoUrl
		Rs("Fk_GModel_Succeed")=Fk_GModel_Succeed
		Rs("Fk_GModel_Repeat")=Fk_GModel_Repeat
		Rs("Fk_GModel_NoTrash")=Fk_GModel_NoTrash
		Rs("Fk_GModel_TrashPint")=Fk_GModel_TrashPint
		Rs("Fk_GModel_MaxStr")=Fk_GModel_MaxStr
		Rs("Fk_GModel_MinStr")=Fk_GModel_MinStr
		Rs("Fk_GModel_Content")=Fk_GModel_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新留言模型添加成功！")
	Else
		Response.Write("该留言模型已经存在，请重新输入！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：GModelEditForm()
'作    用：修改留言模型表单
'参    数：
'==========================================
Sub GModelEditForm()
	Dim TempArr2
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_GModel_Name,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MaxStr,Fk_GModel_MinStr,Fk_GModel_Content,Fk_GModel_GoUrl From [Fk_GModel] Where Fk_GModel_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_GModel_Name=FKFun.HTMLDncode(Rs("Fk_GModel_Name"))
		Fk_GModel_Succeed=FKFun.HTMLDncode(Rs("Fk_GModel_Succeed"))
		Fk_GModel_Repeat=FKFun.HTMLDncode(Rs("Fk_GModel_Repeat"))
		Fk_GModel_NoTrash=Rs("Fk_GModel_NoTrash")
		Fk_GModel_TrashPint=FKFun.HTMLDncode(Rs("Fk_GModel_TrashPint"))
		Fk_GModel_MaxStr=FKFun.HTMLDncode(Rs("Fk_GModel_MaxStr"))
		Fk_GModel_MinStr=FKFun.HTMLDncode(Rs("Fk_GModel_MinStr"))
		Fk_GModel_Content=FKFun.HTMLDncode(Rs("Fk_GModel_Content"))
		Fk_GModel_GoUrl=Rs("Fk_GModel_GoUrl")
		TempArr=Split(Fk_GModel_Content,"|-_-|Fangka|-_-|")
		TempArr2=Split(TempArr(0),"|-^-|Fangka|-^-|")
	Else
		Rs.Close
		Call FKFun.ShowErr("未找到此留言模型，请按键盘上的ESC键退出操作！",1)
	End If
	Rs.Close
%>
<form id="GModelEdit" name="GModelEdit" method="post" action="GModel.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改留言模型[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">名称：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Name" value="<%=Fk_GModel_Name%>" type="text" class="Input" id="Fk_GModel_Name" />&nbsp;&nbsp;<span class="qbox" title="<p>留言模型名称，请输入1-50个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">成功留言转向：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_GoUrl" value="<%=Fk_GModel_GoUrl%>" type="text" class="Input" id="Fk_GModel_GoUrl" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功后转向的链接，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">留言成功提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Succeed" value="<%=Fk_GModel_Succeed%>" type="text" class="Input" id="Fk_GModel_Succeed" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>留言成功提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">重复留言提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_Repeat" value="<%=Fk_GModel_Repeat%>" type="text" class="Input" id="Fk_GModel_Repeat" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>重复留言提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right" class="MainTableTop">反垃圾留言：</td>
            <td>&nbsp;&nbsp;&nbsp;<input type="radio" name="Fk_GModel_NoTrash" class="Input" id="Fk_GModel_NoTrash" value="0"<%=FKFun.BeCheck(Fk_GModel_NoTrash,0)%> />关闭
            <input name="Fk_GModel_NoTrash" class="Input" type="radio" id="Fk_GModel_NoTrash" value="1"<%=FKFun.BeCheck(Fk_GModel_NoTrash,1)%> />开启
            <input name="Fk_GModel_NoTrash" class="Input" type="radio" id="Fk_GModel_NoTrash" value="2"<%=FKFun.BeCheck(Fk_GModel_NoTrash,2)%> />部分开启&nbsp;&nbsp;<span class="qbox" title="<p>此功能开启时，对留言进行中文字数占有量和链接个数进行判断，如果中文字数过低或者链接过多（超过2个），则认为是垃圾留言，如果英文站，请选择关闭或者部分开启，部分开启是不判断中文字数，只判断链接个数。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
	    <tr>
	        <td height="25" align="right">垃圾留言提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_TrashPint" value="<%=Fk_GModel_TrashPint%>" type="text" class="Input" id="Fk_GModel_TrashPint" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>垃圾留言提示，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">字符串过短提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_MinStr" value="<%=Fk_GModel_MinStr%>" type="text" class="Input" id="Fk_GModel_MinStr" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>字符串过短提示，{-留言条目-}:留言条目，{-留言长度-}：限制长度，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">字符串过长提示：</td>
	        <td>&nbsp;&nbsp;&nbsp;<input name="Fk_GModel_MaxStr" value="<%=Fk_GModel_MaxStr%>" type="text" class="Input" id="Fk_GModel_MaxStr" size="50" />&nbsp;&nbsp;<span class="qbox" title="<p>字符串过长提示，{-留言条目-}:留言条目，{-留言长度-}：限制长度，请输入1-255个字符（1个汉字为2个字符）。</p>"><img src="Images/help.jpg" /></span></td>
	        </tr>
        <tr>
            <td height="30" align="right">留言条目：</td>
            <td><p>&nbsp;&nbsp;注：留言标题不可删除，但可以改为其他名字用于前台提示，另外表单标识不可重复</p>
            	<div class="gModel">&nbsp;&nbsp;名称：<input name="Fk_GModel_Content" type="text" class="Input" id="Fk_GModel_Content" value="<%=TempArr2(0)%>" /><br />&nbsp;&nbsp;最小字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="<%=TempArr2(1)%>" />&nbsp;最大字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="<%=TempArr2(2)%>" /><br />&nbsp;&nbsp;表单标识：<input name="Fk_GModel_Content" type="text" class="Input" value="<%=TempArr2(3)%>" id="Fk_GModel_Content" /><input type="hidden" name="Fk_GModel_Content" id="Fk_GModel_Content" value="|-_-|Fangka|-_-|" /></div>
<%
For i=1 To UBound(TempArr)
	If Instr(TempArr(i),"|-^-|Fangka|-^-|") Then
		TempArr2=Split(TempArr(i),"|-^-|Fangka|-^-|")
%>
            	<div id="s<%=i%>" class="gModel">&nbsp;&nbsp;名称：<input name="Fk_GModel_Content" type="text" class="Input" id="Fk_GModel_Content" value="<%=TempArr2(0)%>" /><br />&nbsp;&nbsp;最小字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="<%=TempArr2(1)%>" />&nbsp;最大字符数：<input name="Fk_GModel_Content" type="text" class="Input" size="5" id="Fk_GModel_Content" value="<%=TempArr2(2)%>" /><br />&nbsp;&nbsp;表单标识：<input name="Fk_GModel_Content" type="text" class="Input" value="<%=TempArr2(3)%>" id="Fk_GModel_Content" />&nbsp;<input type="checkbox" name="Fk_GModel_Content" class="Input" value="Show" id="Fk_GModel_Content"<%If UBound(TempArr2)=4 Then%> checked="checked"<%End If%> />后台列表显示<input type="hidden" name="Fk_GModel_Content" id="Fk_GModel_Content" value="|-_-|Fangka|-_-|" />&nbsp;<a href="javascript:void(0);" onclick="$('#s<%=i%>').remove();" title="删除">删除</a></div>
<%
	End If
Next
%>
            	<div id="t">&nbsp;&nbsp;添加留言条目</div>
                &nbsp;&nbsp;名称：<input name="model1" type="text" class="Input" id="model1" /><br />
                &nbsp;&nbsp;最小字符数：<input name="model2" type="text" class="Input" value="0" size="5" id="model2" />
                最大字符数：<input name="model3" type="text" class="Input" value="50" size="5" id="model3" /><br />
                &nbsp;&nbsp;表单标识：<input name="model4" type="text" class="Input" value="GBook_" id="model4" />&nbsp;<input type="checkbox" name="model5" class="Input" value="Show" id="model5" />后台列表显示
                <input type="button" class="Button" onclick="GModelAdds($('#model1').val(),$('#model2').val(),$('#model3').val(),$('#model4').val(),$('#model5:checked').val());" name="Adds" id="Adds" value="添 加" />&nbsp;&nbsp;<span class="qbox" title="<p>留言模型条目，理论可添加无数个，票数必须是数字。</p>"><img src="Images/help.jpg" /></span>
                <p>&nbsp;</p>
            </td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('GModelEdit','GModel.asp?Type=5',0,'',0,1,'MainRight','GModel.asp?Type=1');" class="Button" name="Enter" id="Enter" value="修 改" />
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：GModelEditDo
'作    用：执行修改留言模型
'参    数：
'==============================
Sub GModelEditDo()
	Dim TempArr2
	Fk_GModel_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Name")))
	Fk_GModel_GoUrl=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_GoUrl")))
	Fk_GModel_Succeed=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Succeed")))
	Fk_GModel_Repeat=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Repeat")))
	Fk_GModel_TrashPint=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_TrashPint")))
	Fk_GModel_MaxStr=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_MaxStr")))
	Fk_GModel_MinStr=FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_MinStr")))
	Fk_GModel_Content=Replace(FKFun.HTMLEncode(Trim(Request.Form("Fk_GModel_Content"))),", ","|-^-|Fangka|-^-|")
	Fk_GModel_Content=Replace(Fk_GModel_Content,"|-^-|Fangka|-^-||-_-|Fangka|-_-|","|-_-|Fangka|-_-|")
	Fk_GModel_Content=Replace(Fk_GModel_Content,"|-_-|Fangka|-_-||-^-|Fangka|-^-|","|-_-|Fangka|-_-|")
	Fk_GModel_NoTrash=Trim(Request.Form("Fk_GModel_NoTrash"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_GModel_Name,1,50,0,"请输入模型名称！","模型名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_GModel_GoUrl,0,255,0,"请输入留言成功转向链接！","留言成功转向链接不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Succeed,1,255,0,"请输入留言成功提示！","留言成功提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Repeat,1,255,0,"请输入重复留言提示！","重复留言不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_TrashPint,1,255,0,"请输入垃圾留言提示！","垃圾留言提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_MaxStr,1,255,0,"请输入字符过长提示！","字符过长提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_MinStr,1,255,0,"请输入字符过短提示！","字符过短提示不能大于255个字符！")
	Call FKFun.ShowString(Fk_GModel_Content,20,50000,0,"请录入留言条目！","留言条目不能大于50000个字符！")
	Call FKFun.ShowNum(Fk_GModel_NoTrash,"请选择反垃圾留言是否开启选项！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	TempArr=Split(Fk_GModel_Content,"|-_-|Fangka|-_-|")
	For Each Temp In TempArr
		If Trim(Temp)<>"" Then
			TempArr2=Split(Temp,"|-^-|Fangka|-^-|")
			Call FKFun.ShowString(TempArr2(0),1,50,0,"留言条目名称不能留空！","留言条目名称不能大于50个字符！")
			Call FKFun.ShowString(TempArr2(3),1,50,0,"留言条目标识不能留空！","留言条目标识不能大于50个字符！")
			Call FKFun.ShowNum(TempArr2(1),"最小字符数必须是数字！")
			Call FKFun.ShowNum(TempArr2(2),"最大字符数必须是数字！")
		End If
	Next
	Sqlstr="Select Fk_GModel_Id,Fk_GModel_Name,Fk_GModel_Succeed,Fk_GModel_Repeat,Fk_GModel_NoTrash,Fk_GModel_TrashPint,Fk_GModel_MaxStr,Fk_GModel_MinStr,Fk_GModel_Content,Fk_GModel_GoUrl From [Fk_GModel] Where Fk_GModel_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_GModel_Name")=Fk_GModel_Name
		Rs("Fk_GModel_GoUrl")=Fk_GModel_GoUrl
		Rs("Fk_GModel_Succeed")=Fk_GModel_Succeed
		Rs("Fk_GModel_Repeat")=Fk_GModel_Repeat
		Rs("Fk_GModel_NoTrash")=Fk_GModel_NoTrash
		Rs("Fk_GModel_TrashPint")=Fk_GModel_TrashPint
		Rs("Fk_GModel_MaxStr")=Fk_GModel_MaxStr
		Rs("Fk_GModel_MinStr")=Fk_GModel_MinStr
		Rs("Fk_GModel_Content")=Fk_GModel_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("留言模型修改成功！")
	Else
		Response.Write("留言模型不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：GModelDelDo
'作    用：执行删除留言模型
'参    数：
'==============================
Sub GModelDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select Fk_Module_Id From [Fk_Module] Where Fk_Module_GModel=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Rs.Close
		Call FKFun.ShowErr("该留言模型尚在使用中，无法删除！",2)
	End If
	Rs.Close
	Sqlstr="Select Fk_GModel_Id From [Fk_GModel] Where Fk_GModel_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("留言模型删除成功！")
	Else
		Response.Write("留言模型不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->