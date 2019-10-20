<%
'==========================================
'文 件 名：Class/Cls_PageCode.asp
'文件用途：图片处理函数类
'版权所有：
'==========================================

Class Cls_PageCode
	Private pCode
	
	'==============================
	'函 数 名：cIndex
	'作    用：首页代码
	'参    数：
	'==============================
	Public Function cIndex()
		pCode=FKTemplate.GetTemplate("sy",0,0,"")
		pCode=FKTemplate.FileChange(pCode)
		pCode=FKTemplate.SiteChange(pCode)
		pCode=Replace(pCode,"{$ModuleId$}","0")
		pCode=FKTemplate.TemplateDo(pCode)
		pCode=FKTemplate.ReChangeField(pCode)
		cIndex=pCode
	End Function
	
	'==============================
	'函 数 名：cModule
	'作    用：模块代码
	'参    数：
	'pModuleId  模块ID
	'pType      模块类型
	'==============================
	Public Function cModule(pModuleId,pType)
		Dim TempModuleTemplate,TempModuleUrl,TempModuleMenu,TempModuleIsIndex,TempMenuTepmlate
		Sqlstr="Select Fk_Module_Template,Fk_Module_Url,Fk_Module_Menu,Fk_Module_IsIndex From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Id="&pModuleId&""
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempModuleTemplate=Rs("Fk_Module_Template")
			TempModuleUrl=Rs("Fk_Module_Url")
			TempModuleMenu=Rs("Fk_Module_Menu")
			TempModuleIsIndex=Rs("Fk_Module_IsIndex")
		Else
			Call FKFun.ShowErr("内容页面没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
		End If
		Rs.Close
		Sqlstr="Select Fk_Menu_Template From [Fk_Menu] Where Fk_Menu_Id="&TempModuleMenu&""
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempMenuTepmlate=Rs("Fk_Menu_Template")
		End If
		Rs.Close
		Select Case pType
			Case 0
				pCode=FKTemplate.GetTemplate("yspage",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 1
				pCode=FKTemplate.GetTemplate("ysxwlist",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 2
				pCode=FKTemplate.GetTemplate("yscplist",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 3
				pCode=FKTemplate.GetTemplate("ysxxy",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 4
				pCode=FKTemplate.GetTemplate("ysly",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 5
				pCode="正在转向"&TempModuleUrl&"，请稍等！<meta http-equiv=""refresh"" content=""1;URL="&TempModuleUrl&""">"
			Case 6
				pCode=FKTemplate.GetTemplate("yszp",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
			Case 7
				pCode=FKTemplate.GetTemplate("ysxzlist",TempModuleTemplate,TempModuleIsIndex,TempMenuTepmlate)
		End Select
		If pType<>5 Then
			pCode=FKTemplate.FileChange(pCode)
			pCode=FKTemplate.SiteChange(pCode)
			pCode=FKTemplate.ModuleChange(pCode,pModuleId,pType)
			pCode=FKTemplate.TemplateDo(pCode)
			If pType=4 Then
				pCode=FKTemplate.GBookPageChange(pCode,pModuleId)
			End If
		End If
		pCode=FKTemplate.ReChangeField(pCode)
		cModule=pCode
	End Function

	'==============================
	'函 数 名：cPage
	'作    用：内容页代码
	'参    数：
	'pId         内容ID
	'pModuleId   模块ID
	'pType       模块类型
	'==============================
	Public Function cPage(pId,pModuleId,pType)
		Dim TempModuleLowTemplate,TempModuleMenu,TempMenuTepmlate
		Sqlstr="Select Fk_Module_LowTemplate,Fk_Module_Menu From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Id="&pModuleId&""
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempModuleLowTemplate=Rs("Fk_Module_LowTemplate")
			TempModuleMenu=Rs("Fk_Module_Menu")
		Else
			Call FKFun.ShowErr("内容页面没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
		End If
		Rs.Close
		Sqlstr="Select Fk_Menu_Template From [Fk_Menu] Where Fk_Menu_Id="&TempModuleMenu&""
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempMenuTepmlate=Rs("Fk_Menu_Template")
		End If
		Rs.Close
		If pType=1 Then
			If IsNumeric(pId) Then
				Sqlstr="Select Fk_Article_Id,Fk_Article_Template From [Fk_Article] Where Fk_Article_Show=1 And Fk_Article_Id="&pId&""
			Else
				Sqlstr="Select Fk_Article_Id,Fk_Article_Template From [Fk_Article] Where Fk_Article_Show=1 And Fk_Article_Module="&pModuleId&" And Fk_Article_FileName='"&pId&"'"
			End If
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				pId=Rs("Fk_Article_Id")
				If Rs("Fk_Article_Template")>0 Then
					TempModuleLowTemplate=Rs("Fk_Article_Template")
				End If
			Else
				Call FKFun.ShowErr("文章没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
			End If
			Rs.Close
		ElseIf pType=2 Then
			If IsNumeric(pId) Then
				Sqlstr="Select Fk_Product_Id,Fk_Product_Template From [Fk_Product] Where Fk_Product_Show=1 And Fk_Product_Id="&pId&""
			Else
				Sqlstr="Select Fk_Product_Id,Fk_Product_Template From [Fk_Product] Where Fk_Product_Show=1 And Fk_Product_Module="&pModuleId&" And Fk_Product_FileName='"&pId&"'"
			End If
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				pId=Rs("Fk_Product_Id")
				If Rs("Fk_Product_Template")>0 Then
					TempModuleLowTemplate=Rs("Fk_Product_Template")
				End If
			Else
				Call FKFun.ShowErr("产品没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
			End If
			Rs.Close
		ElseIf pType=7 Then
			If IsNumeric(pId) Then
				Sqlstr="Select Fk_Down_Id,Fk_Down_Template From [Fk_Down] Where Fk_Down_Show=1 And Fk_Down_Id="&pId&""
			Else
				Sqlstr="Select Fk_Down_Id,Fk_Down_Template From [Fk_Down] Where Fk_Down_Show=1 And Fk_Down_Module="&pModuleId&" And Fk_Down_FileName='"&pId&"'"
			End If
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				pId=Rs("Fk_Down_Id")
				If Rs("Fk_Down_Template")>0 Then
					TempModuleLowTemplate=Rs("Fk_Down_Template")
				End If
			Else
				Call FKFun.ShowErr("下载没找到，3秒后返回首页！<meta http-equiv=""refresh"" content=""3;URL="&SiteDir&""">",0)
			End If
			Rs.Close
		End If
		If pType=1 Then
			pCode=FKTemplate.GetTemplate("ysxw",TempModuleLowTemplate,0,TempMenuTepmlate)
		ElseIf pType=2 Then
			pCode=FKTemplate.GetTemplate("yscp",TempModuleLowTemplate,0,TempMenuTepmlate)
		ElseIf pType=7 Then
			pCode=FKTemplate.GetTemplate("ysxz",TempModuleLowTemplate,0,TempMenuTepmlate)
		End If
		pCode=FKTemplate.FileChange(pCode)
		pCode=FKTemplate.SiteChange(pCode)
		If pType=1 Then
			pCode=FKTemplate.ArticleChange(pCode,pId)
		ElseIf pType=2 Then
			pCode=FKTemplate.ProductChange(pCode,pId)
		ElseIf pType=7 Then
			pCode=FKTemplate.DownChange(pCode,pId)
		End If
		pCode=FKTemplate.TemplateDo(pCode)
		pCode=FKTemplate.ReChangeField(pCode)
		cPage=pCode
	End Function
	
	'==============================
	'函 数 名：cSubject
	'作    用：专题代码
	'参    数：
	'==============================
	Public Function cSubject(pId,pTemplate)
		pCode=FKTemplate.GetTemplate("subject",pTemplate,0,"")
		pCode=FKTemplate.FileChange(pCode)
		pCode=FKTemplate.SiteChange(pCode)
		pCode=FKTemplate.SubjectChange(pCode,pId)
		pCode=FKTemplate.TemplateDo(pCode)
		pCode=FKTemplate.ReChangeField(pCode)
		cSubject=pCode
	End Function
	
	'==============================
	'函 数 名：cSearch
	'作    用：搜索代码
	'参    数：
	'==============================
	Public Function cSearch(pTemplate)
		If pTemplate<>"" Then
			pCode=FKTemplate.GetTemplate(pTemplate,0,0,"")
		Else
			pCode=FKTemplate.GetTemplate("search",0,0,"")
		End If
		pCode=FKTemplate.FileChange(pCode)
		pCode=FKTemplate.SiteChange(pCode)
		pCode=FKTemplate.SearchChange(pCode)
		pCode=FKTemplate.TemplateDo(pCode)
		pCode=FKTemplate.SearchPageChange(pCode)
		pCode=FKTemplate.ReChangeField(pCode)
		cSearch=pCode
	End Function

	'==============================
	'函 数 名：PageReset
	'作    用：地址跳转
	'参    数：
	'==============================
	Public Function PageReset(PageU)
		If PageU="" Then
			Response.Redirect("Index"&FKTemplate.GetHtmlSuffix())
			Call FKDB.DB_Close()
			Session.CodePage=936
			Response.End()
		Else
			PageU=Replace(Replace(Replace(PageU,"index.asp",""),"Index.asp",""),"?","")
			Response.Redirect(PageU)
			Call FKDB.DB_Close()
			Session.CodePage=936
			Response.End()
		End If
	End Function
End Class
%>
