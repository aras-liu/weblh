<%
'==========================================
'文 件 名：Class/Cls_Html.asp
'文件用途：静态生成函数类
'版权所有：
'==========================================

Class Cls_Html
	Private PageCode
	
	'==============================
	'函 数 名：CreatIndex
	'作    用：生成首页
	'参    数：
	'==============================
	Public Function CreatIndex()
		PageCode=FKPageCode.cIndex()
		Call FKFso.CreateFile(FileDir&"Index.html",PageCode)
		Response.Write("<p><a href=""../Index.html"" target=""_blank"">首页生成成功</a></p>")
		Response.Flush()
		Response.Clear()
	End Function
	
	'==============================
	'函 数 名：CreatModule
	'作    用：生成模块页
	'参    数：
	'==============================
	Public Function CreatModule(pModuleId,pType,pUrl,pName,pCreatUrl,pCreatModuleId,pCreatCheck)
		Dim ShowCreatFileName,CreatFileName,GBookPageCheck,PageCreat,PageCreatEnd
		If Right(pUrl,1)="/" Then
			CreatFileName=pUrl&"Index"&FKTemplate.GetHtmlSuffix()
			If Clng(PageNow)>1 Then
				CreatFileName=Replace(CreatFileName,"Index.","Index_"&PageNow&".")
			Else
				GBookPageCheck=pUrl&"Index_2"&FKTemplate.GetHtmlSuffix()
			End If
		ElseIf pType=4 Then
			CreatFileName=pUrl
			If Clng(PageNow)>1 Then
				CreatFileName=Replace(CreatFileName,FKTemplate.GetHtmlSuffix(),"__"&PageNow&FKTemplate.GetHtmlSuffix())
			Else
				GBookPageCheck=Replace(CreatFileName,FKTemplate.GetHtmlSuffix(),"__2"&FKTemplate.GetHtmlSuffix())
			End If
		Else
			CreatFileName=pUrl
		End If
		ShowCreatFileName=CreatFileName
		If SiteDir<>FileDir Then
			CreatFileName=Replace(CreatFileName,SiteDir,FileDir,1,1)
		End If
		PageCode=FKPageCode.cModule(pModuleId,pType)
		Call FKFso.CreateFile(CreatFileName,PageCode)
		If Instr(",1,2,7,",","&pType&",")>0 And pCreatCheck=1 Then
			If PageNow=1 Then
				Response.Write("<p><a href="""&ShowCreatFileName&""" target=""_blank"">“"&pName&"”第"&PageNow&"页生成成功</a></p>")
			End If
			PageCreat=PageNow+1
			If (PageCounts-PageNow)>20 Then
				PageCreatEnd=PageNow+20
			Else
				PageCreatEnd=PageCounts
			End If
			For PageCreat=PageCreat To PageCreatEnd
				PageNow=PageCreat
				Call CreatModule(pModuleId,pType,pUrl,pName,pCreatUrl,pCreatModuleId,0)
			Next
			If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=<%=pCreatUrl&pCreatModuleId%>&Page=<%=PageNow%>">
<%
			Else
%>
<meta http-equiv="refresh" content="1;URL=<%=pCreatUrl&pModuleId%>">
<%
			End If
			Call FKFun.ShowErr("正在生成，请稍候！",2)
		ElseIf pType=4 And Instr(PageCode,GBookPageCheck)>0 And pCreatCheck=1 Then
			If PageNow=1 Then
				Response.Write("<p><a href="""&ShowCreatFileName&""" target=""_blank"">“"&pName&"”第"&PageNow&"页生成成功</a></p>")
			End If
			PageCreat=PageNow+1
			If (PageCounts-PageNow)>1 Then
				PageCreatEnd=PageNow+1
			Else
				PageCreatEnd=PageCounts
			End If
			For PageCreat=PageCreat To PageCreatEnd
				PageNow=PageCreat
				Call CreatModule(pModuleId,pType,pUrl,pName,pCreatUrl,pCreatModuleId,0)
			Next
			If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=<%=pCreatUrl&pCreatModuleId%>&Page=<%=PageNow%>">
<%
			Else
%>
<meta http-equiv="refresh" content="1;URL=<%=pCreatUrl&pModuleId%>">
<%
			End If
			Call FKFun.ShowErr("正在生成，请稍候！",2)
		Else
			If PageNow>1 Then
				Response.Write("<p><a href="""&ShowCreatFileName&""" target=""_blank"">“"&pName&"”第"&PageNow&"页生成成功</a></p>")
			Else
				Response.Write("<p><a href="""&ShowCreatFileName&""" target=""_blank"">“"&pName&"”生成成功</a></p>")
			End If
		End If
		Response.Flush()
		Response.Clear()
	End Function
	
	'==============================
	'函 数 名：CreatPage
	'作    用：生成内容页
	'参    数：
	'==============================
	Public Function CreatPage(pId,pModuleId,pType,pUrl,pName)
		Dim ShowCreatFileName
		ShowCreatFileName=pUrl
		PageCode=FKPageCode.cPage(pId,pModuleId,pType)
		If SiteDir<>FileDir Then
			pUrl=Replace(pUrl,SiteDir,FileDir,1,1)
		End If
		Call FKFso.CreateFile(pUrl,PageCode)
		Response.Write("<p><a href="""&ShowCreatFileName&""" target=""_blank"">“"&pName&"”生成成功</a></p>")
		Response.Flush()
		Response.Clear()
	End Function
		
	'==============================
	'函 数 名：CreatSubject
	'作    用：生成专题页
	'参    数：
	'==============================
	Public Function CreatSubject(pSubjectId,pTemplate,pUrl,pName)
		pUrl=FileDir&"Subject/"&pUrl&FKTemplate.GetHtmlSuffix()
		PageCode=FKPageCode.cSubject(pSubjectId,pTemplate)
		Call FKFso.CreateFile(pUrl,PageCode)
		Response.Write("<p><a href="""&pUrl&""" target=""_blank"">“"&pName&"”生成成功</a></p>")
		Response.Flush()
		Response.Clear()
	End Function
End Class
%>
