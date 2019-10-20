<%
'==========================================
'文 件 名：Inc/PageCode.asp
'文件用途：页码处理文件
'版权所有：
'==========================================

'==============================
'函 数 名：CheckPageCode
'作    用：选择页面显示函数
'参    数：
'==============================
Function CheckPageCode(pId,PageUrl)
	pId=Clng(pId)
	Select Case pId
		Case 0
			CheckPageCode=ShowPageCode(PageUrl)
		Case 1
			CheckPageCode=ShowPageCode2(PageUrl)
	End Select
End Function
		
'==============================
'函 数 名：ShowPageCode1
'作    用：显示页码
'参    数：链接PageUrl
'==============================
Function ShowPageCode(PageUrl)
	If PageNow>1 Then
		ShowPageCode="<a href="""&Replace(PageUrl,"{Pages}",1)&""" title=""第一页"">第一页</a>"
		ShowPageCode=ShowPageCode&"&nbsp;"
		ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",PageNow-1)&""" title=""上一页"">上一页</a>"
		If Fk_Site_Html>0 And SearchStr="" Then
			ShowPageCode=Replace(ShowPageCode,"__1.",".")
			ShowPageCode=Replace(ShowPageCode,"_1.",".")
		End If
	Else
		ShowPageCode=ShowPageCode&"第一页"
		ShowPageCode=ShowPageCode&"&nbsp;"
		ShowPageCode=ShowPageCode&"上一页"
	End If
	ShowPageCode=ShowPageCode&"&nbsp;"
	If PageCounts>PageNow Then
		ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",PageNow+1)&""" title=""下一页"">下一页</a>"
		ShowPageCode=ShowPageCode&"&nbsp;"
		ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",PageCounts)&""" title=""尾页"">尾页</a>"
	Else
		ShowPageCode=ShowPageCode&"下一页"
		ShowPageCode=ShowPageCode&"&nbsp;"
		ShowPageCode=ShowPageCode&"尾页"
	End If
	ShowPageCode=ShowPageCode&"&nbsp;"&TempPageSize&"条/页&nbsp;共"&PageCounts&"页/"&PageAll&"条&nbsp;当前第"&PageNow&"页&nbsp;"
	ShowPageCode=ShowPageCode&"<select name=""Change_Page"" id=""Change_Page"" onChange=""window.location.href=this.options[this.selectedIndex].value"">"
	For i=1 To PageCounts
		If i=1 Then
			If i=PageNow Then
				If Instr(PageUrl,"_{Pages}")>0 Then
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"_{Pages}","")&""" selected=""selected"">第"&i&"页</option>"
				Else
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""" selected=""selected"">第"&i&"页</option>"
				End If
			Else
				If Instr(PageUrl,"_{Pages}")>0 Then
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"_{Pages}","")&""">第"&i&"页</option>"
				Else
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""">第"&i&"页</option>"
				End If
			End If
		Else
			If i=PageNow Then
				ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""" selected=""selected"">第"&i&"页</option>"
			Else
				ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""">第"&i&"页</option>"
			End If
		End If
	Next
	ShowPageCode=ShowPageCode&"</select>"
	If Fk_Site_Sign<>"" And Fk_Site_Html=0 Then
		ShowPageCode=Replace(ShowPageCode,Fk_Site_PageSign&"1","")
		If Right(ShowPageCode,1)=Fk_Site_PageSign Then
			ShowPageCode=Left(ShowPageCode,Len(ShowPageCode)-1)
		End If
	Else
		ShowPageCode=Replace(ShowPageCode,"Index_1"&FKTemplate.GetHtmlSuffix(),"")
		ShowPageCode=Replace(ShowPageCode,"Index"&FKTemplate.GetHtmlSuffix(),"")
	End If
End Function
		
'==============================
'函 数 名：ShowPageCode2
'作    用：显示页码方案二
'参    数：链接PageUrl
'==============================
Function ShowPageCode2(PageUrl)
	Dim pSId,pEId,outId
	pSId=0
	pEId=0
	If PageNow=1 Then
		ShowPageCode2="Previous Page"
	Else
		ShowPageCode2="<a href="""&Replace(PageUrl,"{Pages}",(PageNow-1))&""" title=""Previous Page"">Previous Page</a>"
	End If
	If PageCounts<=10 Then
		pSId=1
		pEId=PageCounts
	ElseIf PageNow<=5 Then
		pSId=1
		pEId=10
	ElseIf (PageCounts-PageNow)<=5 Then
		pSId=PageCounts-10
		pEId=10
	Else
		pSId=PageNow-4
		pEId=PageNow+5
	End If
	For outId=pSId To pEId
		If outId=PageNow Then
			ShowPageCode2=ShowPageCode2&" ["&outId&"] "
		Else
			ShowPageCode2=ShowPageCode2&" <a href="""&Replace(PageUrl,"{Pages}",outId)&""" title="""&outId&""">"&outId&"</a> "
		End If
	Next
	If PageNow<PageCounts Then
		ShowPageCode2=ShowPageCode2&" <a href="""&Replace(PageUrl,"{Pages}",(PageNow+1))&""" title=""Next Page"">Next Page</a> "
		ShowPageCode2=ShowPageCode2&" <a href="""&Replace(PageUrl,"{Pages}",PageCounts)&""">&gt;&gt;|</a> "
	Else
		ShowPageCode2=ShowPageCode2&" Next Page "
	End If
	If Fk_Site_Sign<>"" And Fk_Site_Html=0 Then
		ShowPageCode2=Replace(ShowPageCode2,Fk_Site_PageSign&"1","")
		If Right(ShowPageCode2,1)=Fk_Site_PageSign Then
			ShowPageCode2=Left(ShowPageCode2,Len(ShowPageCode2)-1)
		End If
	Else
		ShowPageCode2=Replace(ShowPageCode2,"Index_1"&FKTemplate.GetHtmlSuffix(),"")
		ShowPageCode2=Replace(ShowPageCode2,"Index"&FKTemplate.GetHtmlSuffix(),"")
	End If
End Function
%>
