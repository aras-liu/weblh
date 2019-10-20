<%
'==========================================
'文 件 名：Class/Cls_Admin.asp
'文件用途：后台函数类
'版权所有：
'==========================================

Class Cls_Admin
	'==============================
	'函 数 名：GetNavGo
	'作    用：输出模块管理页面
	'参    数：
	'GetModuleType  模块类型
	'GetModuleId    模块编号
	'==============================
	Public Function GetNavGo(GetModuleType,GetModuleId)
		Select Case GetModuleType
			Case 0
				GetNavGo="void(0);"
			Case 1
				GetNavGo="SetRContent('MainRight','Article.asp?Type=1&ModuleId="&GetModuleId&"')"
			Case 2
				GetNavGo="SetRContent('MainRight','Product.asp?Type=1&ModuleId="&GetModuleId&"')"
			Case 3
				GetNavGo="ShowBox('Info.asp?Type=1&ModuleId="&GetModuleId&"');"
			Case 4
				GetNavGo="SetRContent('MainRight','GBook.asp?Type=1&ModuleId="&GetModuleId&"')"
			Case 5
				GetNavGo="void(0);"
			Case 6
				GetNavGo="SetRContent('MainRight','Job.asp?Type=1&ModuleId="&GetModuleId&"')"
			Case 7
				GetNavGo="SetRContent('MainRight','Down.asp?Type=1&ModuleId="&GetModuleId&"')"
			Case 8
				GetNavGo="void(0);"
		End Select
	End Function

	'==============================
	'函 数 名：ShowField
	'作    用：输出自定义字段
	'参    数：
	'GetType           自定义字段类型
	'GetModel          字段用途
	'GetSql            查询SQL区段
	'GetContent        自定义字段内容数组
	'GetEditorClass    编辑器类型
	'==============================
	Public Function ShowField(GetType,GetModel,GetSql,GetContent,GetEditorClass)
		Dim FieldValue,FieldTemp,SelectTempArr,SelectTemp
		Sqlstr="Select Fk_Field_Name,Fk_Field_Tag,Fk_Field_Help,Fk_Field_Option From [Fk_Field] Where Fk_Field_Type="&GetType&" And Fk_Field_Model="&GetModel&""&GetSql&" Order By Fk_Field_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
			FieldValue=""
			If Not IsNull(GetContent) Then
				For Each FieldTemp In GetContent
					If Split(FieldTemp,"|-Fangka_Field-|")(0)=Rs("Fk_Field_Tag") Then
						FieldValue=FKFun.HTMLDncode(Split(FieldTemp,"|-Fangka_Field-|")(1))
						Exit For
					End If
				Next
			End If
			Select Case GetType
				Case 0
%>
		<tr>
			<td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
			<td colspan="3">&nbsp;<input name="Field_<%=Rs("Fk_Field_Tag")%>" type="text" class="Input" id="Field_<%=Rs("Fk_Field_Tag")%>" size="50" value="<%=FieldValue%>" />&nbsp;&nbsp;<span class="qbox" title="<p><%=Rs("Fk_Field_Help")%></p>"><img src="Images/help.jpg" /></span></td>
		</tr>
<%
				Case 1
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：<br /><span class="qbox" title="<p><%=Rs("Fk_Field_Help")%></p>"><img src="Images/help.jpg" /></span>&nbsp;&nbsp;</td>
        <td colspan="3"><textarea name="Field_<%=Rs("Fk_Field_Tag")%>" style="width:100%;" class="<%=GetEditorClass%>" rows="15" id="Field_<%=Rs("Fk_Field_Tag")%>"><%=FieldValue%></textarea></td>
    </tr>
<%
				Case 2
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
	    <td colspan="3">&nbsp;<input name="Field_<%=Rs("Fk_Field_Tag")%>" type="text" class="Input" id="Field_<%=Rs("Fk_Field_Tag")%>" size="35" value="<%=FieldValue%>" />&nbsp;&nbsp;<span class="qbox" title="<p><%=Rs("Fk_Field_Help")%></p>"><img src="Images/help.jpg" /></span><br />
        &nbsp;<iframe frameborder="0" width="330" height="25" scrolling="No" id="Field_<%=Rs("Fk_Field_Tag")%>s" name="Field_<%=Rs("Fk_Field_Tag")%>s" src="PicUpLoad.asp?Type=5&Form=SystemSet&Input=Field_<%=Rs("Fk_Field_Tag")%>"></iframe></td>
    </tr>
<%
				Case 3
					SelectTempArr=Split(Rs("Fk_Field_Option"),"<br />")
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="3">&nbsp;<select name="Field_<%=Rs("Fk_Field_Tag")%>" class="Input" id="Field_<%=Rs("Fk_Field_Tag")%>">
<%
		For Each SelectTemp In SelectTempArr
			If Instr(SelectTemp,"||")>0 Then
%>
            <option value="<%=Split(SelectTemp,"||")(0)%>"<%=FKFun.BeSelect(Split(SelectTemp,"||")(0),FieldValue)%>><%=Split(SelectTemp,"||")(1)%></option>
<%
			End If
		Next
%>
            </select>&nbsp;&nbsp;<span class="qbox" title="<p><%=Rs("Fk_Field_Help")%></p>"><img src="Images/help.jpg" /></span></td>
    </tr>
<%
			End Select
			Rs.MoveNext
		Wend
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：ShowField
	'作    用：获取自定义字段值
	'参    数：
	'GetModel          字段用途
	'GetSql            查询SQL区段
	'==============================
	Public Function GetFieldData(GetModel,GetSql)
		Dim FieldValue
		GetFieldData=""
		If GetSql<>"" Then
			GetSql=" And "&GetSql
		End If
		Sqlstr="Select Fk_Field_Name,Fk_Field_Tag,Fk_Field_Type From [Fk_Field] Where Fk_Field_Model="&GetModel&""&GetSql&" Order By Fk_Field_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
			FieldValue=Trim(Request.Form("Field_"&Rs("Fk_Field_Tag")))
			If Rs("Fk_Field_Type")=0 Then
				FieldValue=FKFun.HTMLEncode(FieldValue)
			End If
			If GetFieldData="" Then
				GetFieldData=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FieldValue
			Else
				GetFieldData=GetFieldData&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FieldValue
			End If
			Rs.MoveNext
		Wend
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：CheckDandF
	'作    用：检测是否是禁止使用的目录名或文件名
	'参    数：
	'CheckType  检测类型，dir为目录,file为文件
	'CheckName  检测名称
	'DirLevel   目录级别
	'==============================
	Public Function CheckDandF(CheckType,CheckName,DirLevel)
		If CheckName<>"" Then
			If CheckType="dir" Then
				If DirLevel=0 Then
					If Instr(",admin,class,inc,js,plugin,skin,subject,up,search,",","&LCase(CheckName)&",")>0 Then
						Call FKFun.ShowErr("模块文件名/目录与系统已有文件夹冲突！",2)
					End If
				End If
			ElseIf CheckType="file" Then
				If IsNumeric(CheckName) Then
					Call FKFun.ShowErr("文件名不可用纯数字！",2)
				End If
			End If
			If Instr(CheckName,".")>0 Or Left(LCase(CheckName),5)="index" Then
				Call FKFun.ShowErr("文件名不能有“.”或者前5个字符为index！",2)
			End If
		End If
	End Function

	'==============================
	'函 数 名：AdminCheck
	'作    用：管理员校验
	'参    数：
	'CheckType  检测类型，1：登录检测；2：登录二次检测；3：子功能权限判别；4：权限判断
	'LimitStr   检测的权限
	'Limits     用户权限判别字符串
	'==============================
	Public Function AdminCheck(CheckType,LimitStr,Limits)
		Dim TempName,TempPass
		Dim SetHiddenId,SetHiddenStr
		If CheckType=1 Then
			If Request.Cookies("FkAdminName")<>"" And Request.Cookies("FkAdminPass")<>"" Then
				TempName=FKFun.HTMLEncode(Request.Cookies("FkAdminName"))
				TempPass=FKFun.HTMLEncode(Request.Cookies("FkAdminPass"))
				Sqlstr="Select Fk_Admin_Id,Fk_Admin_Limit From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&TempName&"' And Fk_Admin_LoginPass='"&TempPass&"'"
				Rs.Open Sqlstr,Conn,1,1
				If Not Rs.Eof Then
					Response.Cookies("FkAdminId")=Rs("Fk_Admin_Id")
					Response.Cookies("FkAdminLimitId")=Rs("Fk_Admin_Limit")
					If Fk_Site_Dir<>"" Then
						Response.Cookies("FkAdminId").Path="/"
						Response.Cookies("FkAdminLimitId").Path="/"
					End If
					If Rs("Fk_Admin_Limit")>0 Then
						Rs.Close
						Sqlstr="Select Fk_Limit_Set1,Fk_Limit_Set2,Fk_Limit_Content From [Fk_Limit] Where Fk_Limit_Id=" & Request.Cookies("FkAdminLimitId")
						Rs.Open Sqlstr,Conn,1,1
						If Not Rs.Eof Then
							Response.Cookies("FkAdminLimit1")=Rs("Fk_Limit_Set1")
							Response.Cookies("FkAdminLimit2")=Rs("Fk_Limit_Set2")
							Response.Cookies("FkAdminLimit3")=Rs("Fk_Limit_Content")
						Else
							Response.Cookies("FkAdminLimit1")="No"
							Response.Cookies("FkAdminLimit2")="No"
							Response.Cookies("FkAdminLimit3")="No"
						End If
						If Fk_Site_Dir<>"" Then
							Response.Cookies("FkAdminLimit1").Path="/"
							Response.Cookies("FkAdminLimit2").Path="/"
							Response.Cookies("FkAdminLimit3").Path="/"
						End If
					End If
					AdminCheck=True
				Else
					AdminCheck=False
				End If
				Rs.Close
			Else
				AdminCheck=False
			End If
		ElseIf CheckType=2 Then
			If Request.Cookies("FkAdminId")="" Or Request.Cookies("FkAdminLimitId")="" Then
				Call FKFun.ShowErr("您没有登录或者已经退出，请返回首页重新登录！<a href=""Index.asp"">返回首页</a>",1)
			End If
		ElseIf CheckType=3 Then
			If Request.Cookies("FkAdminLimitId")>0 Then
				If LimitStr<>"0" Then
					If Not AdminCheck(4,LimitStr,Limits) Then
						Call FKFun.ShowErr("无权限！",1)
					End If
				Else
					Call FKFun.ShowErr("无权限！",1)
				End If
			End If
		ElseIf CheckType=4 Then
			If Instr(Limits,"|sethidden|") Then
				TempArr=Split(Limits,"|sethidden|")
				Limits=TempArr(0)
				SetHiddenId=TempArr(1)
				SetHiddenStr=TempArr(2)
			Else
				SetHiddenId=""
				SetHiddenStr=",,,,"
			End If
			If Request.Cookies("FkAdminLimitId")>0 Then
				AdminCheck=False
				If Instr(Limits,","&LimitStr&",")>0 Then
					AdminCheck=True
				End If
			Else
				AdminCheck=True
			End If
			If Request.Cookies("CloseSysHidden")<>"1" Then
				If SetHiddenStr<>",,,," And SetHiddenId<>"" Then
					If Instr(SetHiddenStr,","&SetHiddenId&",")>0 Then
						AdminCheck=False
					End If
				End If
			End If
		ElseIf CheckType=5 Then
			If Request.Cookies("CloseSysHidden")<>"1" Then
				If Instr(Limits,","&LimitStr&",")=0 Then
					AdminCheck=True
				Else
					AdminCheck=False
				End If
			Else
				AdminCheck=True
			End If
		End If
	End Function

	'==============================
	'函 数 名：GetAdminDir
	'作    用：获取管理目录
	'参    数：
	'==============================
	Public Function GetAdminDir()
		If Request.ServerVariables("SERVER_PORT")<>"80" Then
			GetAdminDir="http://"&Request.ServerVariables("SERVER_NAME")&":"&Request.ServerVariables("SERVER_PORT")&Request.ServerVariables("URL")
		Else
			GetAdminDir="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
		End If
		GetAdminDir=Left(GetAdminDir,InstrRev(GetAdminDir,"/")-1)
		GetAdminDir=LCase(Mid(GetAdminDir,InstrRev(GetAdminDir,"/")+1))
	End Function

	'==============================
	'函 数 名：ReLoadTemplate
	'作    用：重载模板缓存
	'参    数：
	'NewF   新模板目录
	'==============================
	Public Function ReLoadTemplate(NewF)
		Dim ObjFiles,ObjFile,ObjFiles2,ObjFile2,F2
		Dim ObjFloders,ObjFloder
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		Set F=Fso.GetFolder(Server.MapPath(FileDir&"Skin/"&NewF))
		Set ObjFiles=F.Files
		Application.Lock()
		For Each ObjFile In ObjFiles
			If LCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,"."))))="html" Then
				Sqlstr="Select Fk_Template_Id,Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&Replace(LCase(ObjFile.Name),".html","")&"'"
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					Rs("Fk_Template_Content")=FKFso.FsoFileRead(FileDir&"Skin/"&NewF&"/"&ObjFile.Name)
					Rs.Update()
				Else
					Rs.AddNew()
					Rs("Fk_Template_Name")=LCase(Replace(LCase(ObjFile.Name),".html",""))
					Rs("Fk_Template_Content")=FKFso.FsoFileRead(FileDir&"Skin/"&NewF&"/"&ObjFile.Name)
					Rs.Update()
				End If
				Rs.Close
			End If
		Next
		Set ObjFiles=Nothing
		Set F=Nothing
		Set F=Fso.GetFolder(Server.MapPath(FileDir&"Skin/"&NewF&"/"))
		Set ObjFloders=F.Subfolders
		For Each ObjFloder In ObjFloders
			If Left(ObjFloder.Name,2)="t_" Then
				Set F2=Fso.GetFolder(Server.MapPath(FileDir&"Skin/"&NewF&"/"&ObjFloder.Name&"/"))
				Set ObjFiles2=F2.Files
				For Each ObjFile2 In ObjFiles2
					If LCase(Split(ObjFile2.Name,".")(UBound(Split(ObjFile2.Name,"."))))="html" Then
						Sqlstr="Select Fk_Template_Id,Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&Replace(LCase(ObjFloder.Name&"/"&ObjFile2.Name),".html","")&"'"
						Rs.Open Sqlstr,Conn,1,3
						If Not Rs.Eof Then
							Rs("Fk_Template_Content")=FKFso.FsoFileRead(FileDir&"Skin/"&NewF&"/"&ObjFloder.Name&"/"&ObjFile2.Name)
							Rs.Update()
						Else
							Rs.AddNew()
							Rs("Fk_Template_Name")=Replace(LCase(ObjFloder.Name&"/"&ObjFile2.Name),".html","")
							Rs("Fk_Template_Content")=FKFso.FsoFileRead(FileDir&"Skin/"&NewF&"/"&ObjFloder.Name&"/"&ObjFile2.Name)
							Rs.Update()
						End If
						Rs.Close
					End If
				Next
				Set ObjFiles2=Nothing
				Set F2=Nothing
			End If
		Next
		Application.UnLock()
		Set ObjFiles=Nothing
		Set F=Nothing
		Set Fso=Nothing
	End Function
	
	'==============================
	'函 数 名：ModuleMUrl
	'作    用：生成模块地址
	'参    数：
	'ModuleId          模块ID
	'ModuleType        模块类型
	'ModuleLevelList   模块级联详情
	'ModuleUrlType     转向链接类型
	'ModuleUrl         转向链接
	'ModuleDir         模块目录
	'MenuDir           菜单目录
	'UrlType           链接类型
	'HtmlSuffix        HTML后缀
	'==============================
	Public Function ModuleMUrl(ModuleId,ModuleType,ModuleLevelList,ModuleUrlType,ModuleUrl,ModuleDir,MenuDir,UrlType,HtmlSuffix)
		Dim Rs2,ModuleSuffix,Temp2,TempArr2,TempUrl
		If ModuleUrlType=0 And ModuleType=5 Then
			ModuleMUrl=ModuleUrl
			Exit Function
		End If
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Select Case HtmlSuffix
			Case 0
				ModuleSuffix=".html"
			Case 1
				ModuleSuffix=".htm"
			Case 2
				ModuleSuffix=".shtml"
			Case 3
				ModuleSuffix=".xml"
		End Select
		If UrlType=0 Then
			If ModuleDir="" Or IsNull(ModuleDir) Then
				Select Case ModuleType
					Case 0
						ModuleMUrl="Page"&ModuleId&ModuleSuffix
					Case 1
						ModuleMUrl="Article"&ModuleId&"/"
					Case 2
						ModuleMUrl="Product"&ModuleId&"/"
					Case 3
						ModuleMUrl="Info"&ModuleId&ModuleSuffix
					Case 4
						ModuleMUrl="GBook"&ModuleId&ModuleSuffix
					Case 5
						ModuleMUrl="Url"&ModuleId&ModuleSuffix
					Case 6
						ModuleMUrl="Job"&ModuleId&ModuleSuffix
					Case 7
						ModuleMUrl="Down"&ModuleId&"/"
					Case 8
						ModuleMUrl="None"&ModuleId&"/"
				End Select
			Else
				Select Case ModuleType
					Case 0
						ModuleMUrl=ModuleDir&ModuleSuffix
					Case 1
						ModuleMUrl=ModuleDir&"/"
					Case 2
						ModuleMUrl=ModuleDir&"/"
					Case 3
						ModuleMUrl=ModuleDir&ModuleSuffix
					Case 4
						ModuleMUrl=ModuleDir&ModuleSuffix
					Case 5
						ModuleMUrl=ModuleDir&ModuleSuffix
					Case 6
						ModuleMUrl=ModuleDir&ModuleSuffix
					Case 7
						ModuleMUrl=ModuleDir&"/"
					Case 8
						ModuleMUrl=ModuleDir&"/"
				End Select
			End If
		ElseIf UrlType=1 Then
			ModuleMUrl=GetModuleDirUrl(ModuleId,ModuleType,ModuleDir)
			TempUrl=""
			If ModuleLevelList<>"" Then
				TempArr2=Split(ModuleLevelList,",")
				For Each Temp2 In TempArr2
					If Temp2<>"" Then
						Sqlstr="Select Fk_Module_Type,Fk_Module_Dir From [Fk_Module] Where Fk_Module_Id="&Temp2&""
						Rs2.Open Sqlstr,Conn,1,3
						If Not Rs2.Eof Then
							TempUrl=TempUrl&GetModuleDirUrl(Temp2,Rs2("Fk_Module_Type"),Rs2("Fk_Module_Dir"))
						End If
						Rs2.Close
					End If
				Next
			End If
			ModuleMUrl=TempUrl&ModuleMUrl
			If MenuDir<>"" Then
				ModuleMUrl=MenuDir&"/"&ModuleMUrl
			End If
		Else
			ModuleMUrl="#"
		End If
		Set Rs2=Nothing
	End Function
	
	'==============================
	'函 数 名：GetModuleDirUrl
	'作    用：获取目录链接
	'参    数：
	'ModuleId    模块ID
	'ModuleType  模块类型
	'ModuleDir   模块目录名
	'==============================
	Public Function GetModuleDirUrl(ModuleId,ModuleType,ModuleDir)
		If ModuleDir="" Or IsNull(ModuleDir) Then
			Select Case ModuleType
				Case 0
					GetModuleDirUrl="Page"&ModuleId&"/"
				Case 1
					GetModuleDirUrl="Article"&ModuleId&"/"
				Case 2
					GetModuleDirUrl="Product"&ModuleId&"/"
				Case 3
					GetModuleDirUrl="Info"&ModuleId&"/"
				Case 4
					GetModuleDirUrl="GBook"&ModuleId&"/"
				Case 5
					GetModuleDirUrl="Url"&ModuleId&"/"
				Case 6
					GetModuleDirUrl="Job"&ModuleId&"/"
				Case 7
					GetModuleDirUrl="Down"&ModuleId&"/"
				Case 8
					GetModuleDirUrl="None"&ModuleId&"/"
			End Select
		Else
			GetModuleDirUrl=ModuleDir&"/"
		End If
	End Function
	
	'==============================
	'函 数 名：ReLoadModuleUrl
	'作    用：重新设置模块链接
	'参    数：
	'MenuIds      菜单ID
	'UrlType      链接类型
	'HtmlSuffix   HTML后缀
	'==============================
	Public Function ReLoadModuleUrl(MenuIds,UrlType,HtmlSuffix)
		Dim Rs3
		Set Rs3=Server.Createobject("Adodb.RecordSet")
		Sqlstr="Select Fk_Menu_Dir From [Fk_Menu] Where Fk_Menu_Id=" & MenuIds
		Rs3.Open Sqlstr,Conn,1,1
		If Not Rs3.Eof Then
			Temp=Rs3("Fk_Menu_Dir")
		End If
		Rs3.Close
		Set Rs3=Nothing
		Call ReLoadModuleUrlM(MenuIds,0,Temp,UrlType,HtmlSuffix)
	End Function
	Public Function ReLoadModuleUrlM(MenuIds,LevelId,MenuDir,UrlType,HtmlSuffix)
		Dim Rs2
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sqlstr="Select Fk_Module_Id,Fk_Module_MUrl,Fk_Module_IsIndex,Fk_Module_Type,Fk_Module_LevelList,Fk_Module_Dir,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		Rs2.Open Sqlstr,Conn,1,3
		Application.Lock()
		While Not Rs2.Eof
			If MenuDir<>"" And Rs2("Fk_Module_IsIndex")=1 And UrlType=1 Then
				Rs2("Fk_Module_MUrl")=MenuDir&"/"
			Else
				Rs2("Fk_Module_MUrl")=ModuleMUrl(Rs2("Fk_Module_Id"),Rs2("Fk_Module_Type"),Rs2("Fk_Module_LevelList"),Rs2("Fk_Module_UrlType"),Rs2("Fk_Module_Url"),Rs2("Fk_Module_Dir"),MenuDir,UrlType,HtmlSuffix)
			End If
			Rs2.Update
			Call ReLoadModuleUrlM(MenuIds,Rs2("Fk_Module_Id"),MenuDir,UrlType,HtmlSuffix)
			Rs2.MoveNext
		Wend
		Application.Lock()
		Rs2.Close
		Set Rs2=Nothing
	End Function

	'==============================
	'函 数 名：GetModuleList
	'作    用：输出模块列表
	'参    数：
	'GetType           获取类型
	'MenuIds           菜单ID
	'LevelId           获取级别ID
	'AutoId            选中级别ID
	'ShowModuleCheck   输出内容规则
	'==============================
	Public Function GetModuleList(GetType,MenuIds,LevelId,AutoId,ShowModuleCheck)
		Call GetModuleListM(GetType,MenuIds,LevelId,AutoId,ShowModuleCheck,"")
	End Function
	Public Function GetModuleListM(GetType,MenuIds,LevelId,AutoId,ShowModuleCheck,TitleBack)
		Dim Rs2,TitleBacks,ModuleSee,ii,Temp2
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		If LevelId=0 Then
			TitleBack=""
		End If
		If GetType=0 Or GetType=1 Then  '0：输出ModuleSelectURL列表；1：输出ModuleSelectId列表
			Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			While Not Rs2.Eof
				ModuleSee=0
				If AdminCheck(4,"Module"&Rs2("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
					ModuleSee=1
				ElseIf AdminCheck(4,"See"&Rs2("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
					ModuleSee=2
				End If
				Response.Write(GetType)
				If ModuleSee>0 Then
					If GetType=0 Then
						If ModuleSee=1 Then
							Response.Write("<option value="""&GetNavGo(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"))&""""&FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))&">"&TitleBack&Rs2("Fk_Module_Name")&"</option>")
						Else
							Response.Write("<option value=""void(0);"""&FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))&">"&TitleBack&Rs2("Fk_Module_Name")&"</option>")
						End If
					Else
						If ModuleSee=1 Then
							Response.Write("<option value="""&Rs2("Fk_Module_Id")&""""&FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))&">"&TitleBack&Rs2("Fk_Module_Name")&"</option>")
						Else
							Response.Write("<option value=""0"""&FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))&">"&TitleBack&Rs2("Fk_Module_Name")&"</option>")
						End If
					End If
					If LevelId=0 Then
						TitleBacks="&nbsp;&nbsp;&nbsp;├"
					Else
						TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
					End If
					Call GetModuleListM(GetType,MenuIds,Rs2("Fk_Module_Id"),AutoId,ShowModuleCheck,TitleBacks)
				End If
				Rs2.MoveNext
			Wend
			Rs2.Close
		ElseIf GetType=2 Or GetType=3 Then  '2：输出ModuleLimit列表；3：输出ModuleField列表；
			Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			If Not Rs2.Eof Then
				Response.Write("<ul>")
				While Not Rs2.Eof
					If GetType=2 Then
%>
<li><span class="fleft">├</span><input type="checkbox" name="Fk_Limit_Content" value="<%=Rs2("Fk_Module_Id")%>"<%If Instr(ShowModuleCheck,",Module"&Rs2("Fk_Module_Id")&",")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label"><%=Rs2("Fk_Module_Name")%></label>
<%
						Call GetModuleListM(2,MenuIds,Rs2("Fk_Module_Id"),0,ShowModuleCheck,"")
					Else
						Select Case Rs2("Fk_Module_Type")
							Case 1
								Temp="SeeArticle,"
							Case 2
								Temp="SeeProduct,"
							Case 3
								Temp="SeeInfo,"
							Case 6
								Temp="SeeJob,"
							Case 7
								Temp="SeeDown,"
							Case Else
								Temp=""
						End Select
%>
<li><span class="fleft">├</span><%If Instr(",1,2,3,6,7,",","&Rs2("Fk_Module_Type")&",")>0 Then%><input type="checkbox" name="Fk_Field_Content" value="<%=Temp%>,Module<%=Rs2("Fk_Module_Id")%>"<%If Instr(ShowModuleCheck,",Module"&Rs2("Fk_Module_Id")&",")>0 Then%> checked="checked"<%End If%> /><label href="#" class="label"><%=Rs2("Fk_Module_Name")%> [<%=FKFun.CheckModule(Rs2("Fk_Module_Type"))%>]</label><%Else%><span class="title"><%=Rs2("Fk_Module_Name")%> [<%=FKFun.CheckModule(Rs2("Fk_Module_Type"))%>]</span><%End If%>
<%
						Call GetModuleListM(3,MenuIds,Rs2("Fk_Module_Id"),0,ShowModuleCheck,"")
					End If
					Response.Write("</li>")
					Rs2.MoveNext
				Wend
				Response.Write("</ul>")
			End If
			Rs2.Close
		ElseIf GetType=4 Then  '4：输出Module列表；
			Dim Rs3
			Set Rs3=Server.Createobject("Adodb.RecordSet")
			Sqlstr="Select Fk_Module_Id,Fk_Module_Template,Fk_Module_Name,Fk_Module_Type,Fk_Module_MUrl,Fk_Module_Url,Fk_Module_UrlType,Fk_Module_Dir From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			While Not Rs2.Eof
				If Rs2("Fk_Module_Template")>0 Then
					Sqlstr="Select Fk_Template_Name From [Fk_Template] Where Fk_Template_Id=" & Rs2("Fk_Module_Template")
					Rs3.Open Sqlstr,Conn,1,1
					If Not Rs3.Eof Then
						Temp=Rs3("Fk_Template_Name")
					Else
						Temp="未知模板"
					End If
					Rs3.Close
				Else
					Temp="默认模板"
				End If
%>
<tr>
    <td height="25" align="center"><%=Rs2("Fk_Module_Id")%></td>
    <td align="left">&nbsp;&nbsp;<%=TitleBack%><%=Rs2("Fk_Module_Name")%></td>
    <td align="left">&nbsp;&nbsp;<%=GetIsModuleUrl(Rs2("Fk_Module_Id"),Rs2("Fk_Module_Type"),Rs2("Fk_Module_UrlType"),Rs2("Fk_Module_MUrl"))%></td>
    <td align="center"><%=FKFun.CheckModule(Rs2("Fk_Module_Type"))%><%If Rs2("Fk_Module_Type")=5 Then%>[<%=Rs2("Fk_Module_Url")%>]<%End If%></td>
    <td align="center"><%=Rs2("Fk_Module_Dir")%></td>
    <td align="center"><%=Temp%></td>
    <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Module.asp?Type=4&MenuId=<%=MenuIds%>&Id=<%=Rs2("Fk_Module_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs2("Fk_Module_Name")%>”模块？此操作不可恢复！','Module.asp?Type=6&Id=<%=Rs2("Fk_Module_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
</tr>
<%
				If LevelId=0 Then
					TitleBacks="&nbsp;&nbsp;&nbsp;├"
				Else
					TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
				End If
				Call GetModuleListM(4,MenuIds,Rs2("Fk_Module_Id"),AutoId,ShowModuleCheck,TitleBacks)
				Rs2.MoveNext
			Wend
			Rs2.Close
			Set Rs3=Nothing
		ElseIf GetType=5 Then  '5：输出Module排序操作列表；
			Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type,Fk_Module_Order,Fk_Module_Url From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			While Not Rs2.Eof
%>
<tr>
    <td height="25" align="center"><%=Rs2("Fk_Module_Id")%></td>
    <td align="left">&nbsp;&nbsp;<%=TitleBack%><%=Rs2("Fk_Module_Name")%></td>
    <td align="center"><%=FKFun.CheckModule(Rs2("Fk_Module_Type"))%><%If Rs2("Fk_Module_Type")=5 Then%>[<%=Rs2("Fk_Module_Url")%>]<%End If%></td>
    <td align="center"><input name="Fk_Module_Order<%=Rs2("Fk_Module_Id")%>" type="text" class="Input" id="Fk_Module_Order<%=Rs2("Fk_Module_Id")%>" value="<%=Rs2("Fk_Module_Order")%>" /></td>
</tr>
<%
				If LevelId=0 Then
					TitleBacks="&nbsp;&nbsp;&nbsp;├"
				Else
					TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
				End If
				Call GetModuleListM(5,MenuIds,Rs2("Fk_Module_Id"),AutoId,ShowModuleCheck,TitleBacks)
				Rs2.MoveNext
			Wend
			Rs2.Close
		ElseIf GetType=6 Then  '6：输出ModuleSelect列表；
			Sqlstr="Select Fk_Module_Id,Fk_Module_Name From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			While Not Rs2.Eof
%>
<option value="<%=Rs2("Fk_Module_Id")%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
<%
				If LevelId=0 Then
					TitleBacks="&nbsp;&nbsp;&nbsp;├"
				Else
					TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
				End If
				Call GetModuleListM(6,MenuIds,Rs2("Fk_Module_Id"),AutoId,ShowModuleCheck,TitleBacks)
				Rs2.MoveNext
			Wend
			Rs2.Close
		ElseIf GetType=7 Then
		
		End If
		Set Rs2=Nothing
	End Function
	
	'==============================
	'函 数 名：GetModuleLevelList
	'作    用：输出分类级数参数
	'参    数：
	'ModuleLevelId  要输出的模块
	'==============================
	Public Function GetModuleLevelList(ModuleLevelId)
		GetModuleLevelList=","&GetModuleLevelListM(ModuleLevelId)&ModuleLevelId&","
	End Function
	Public Function GetModuleLevelListM(ModuleLevelId)
		Dim Rs2
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sqlstr="Select Fk_Module_Level From [Fk_Module] Where Fk_Module_Id=" & ModuleLevelId
		Rs2.Open Sqlstr,Conn,1,1
		If Not Rs2.Eof Then
			If Rs2("Fk_Module_Level")>0 Then
				GetModuleLevelListM=GetModuleLevelListM(Rs2("Fk_Module_Level"))&Rs2("Fk_Module_Level")&","
			End If
		End If
		Rs2.Close
		Set Rs2=Nothing
	End Function
	
	'==============================
	'函 数 名：GetIsModuleUrl
	'作    用：获取模块访问地址
	'参    数：
	'ModuleId        模块ID
	'ModuleType      模块类型
	'ModuleUrlType   转向链接类型
	'ModuleUrl       模块链接
	'==============================
	Public Function GetIsModuleUrl(ModuleId,ModuleType,ModuleUrlType,ModuleUrl)
		If (ModuleType=5 And ModuleUrlType=0) Or ModuleType=8 Then
			GetIsModuleUrl="-"
			Exit Function
		Else
			GetIsModuleUrl=ModuleUrl
		End If
		If Fk_Site_Html=1 Then
			GetIsModuleUrl=SiteDir&"Index.asp?Type="&ModuleType&"&Module="&ModuleId
		Else
			If Fk_Site_Sign<>"" And Fk_Site_Html=0 Then
				GetIsModuleUrl=Replace(GetIsModuleUrl,"/",Fk_Site_Sign)
			End If
			If Right(GetIsModuleUrl,1)=Fk_Site_Sign Then
				GetIsModuleUrl=Left(GetIsModuleUrl,Len(GetIsModuleUrl)-1)
			End If
			If Fk_Site_Html=0 Then
				GetIsModuleUrl=SiteDir&"?"&GetIsModuleUrl
			Else
				GetIsModuleUrl=SiteDir&GetIsModuleUrl
			End If
		End If
	End Function
End Class
%>
