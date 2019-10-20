<%
'==========================================
'文 件 名：Class/Cls_Template.asp
'文件用途：模板引擎函数类
'版权所有：
'==========================================

Class Cls_Template
	Private TemplateTag,TemplatePar,TemplateBCode
	Private If1,If2
	Private VauleTemp,VauleArr,ListNo,ListPageNo,ReadCount
	
'========================常规标签处理===========================	
	'==============================
	'函 数 名：SiteChange
	'作    用：替换站点参数
	'参    数：
	'==============================
	Public Function SiteChange(TemplateCode)
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteName$}",Fk_Site_Name)
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteUrl$}",Fk_Site_Url)
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteKeyword$}",Fk_Site_Keyword)
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteDescription$}",Fk_Site_Description)
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteSkin$}",SiteDir&"Skin/"&Fk_Site_Template&"/")
		TemplateCode=ReplaceTag(TemplateCode,"{$SiteDir$}",SiteDir)
		TemplateCode=ReplaceTag(TemplateCode,"{$ProductAllUrl$}",Product_All_Url)
		TemplateCode=ReplaceTag(TemplateCode,"{$NewsAllUrl$}",News_All_Url)
		TemplateCode=ReplaceTag(TemplateCode,"{$ArticleAllUrl$}",Article_All_Url)
		
		'处理自定义标签数据
		For Each Temp In Fk_Site_Field
			TemplateCode=ReplaceTag(TemplateCode,"{$Site_Field_"&Split(Temp,"|-Fangka_Field-|")(0)&"$}",Split(Temp,"|-Fangka_Field-|")(1))
		Next
		If Instr(TemplateCode,"{$SiteArticleCount$}")>0 Then
			Sqlstr="Select Count(Fk_Article_Id) From [Fk_Article] Where Fk_Article_Show=1"
			Rs.Open Sqlstr,Conn,1,1
			If Rs(0)<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteArticleCount$}",Rs(0))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteArticleCount$}",0)
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$SiteProductCount$}")>0 Then
			Sqlstr="Select Count(Fk_Product_Id) From [Fk_Product] Where Fk_Product_Show=1"
			Rs.Open Sqlstr,Conn,1,1
			If Rs(0)<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteProductCount$}",Rs(0))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteProductCount$}",0)
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$SiteDownCount$}")>0 Then
			Sqlstr="Select Count(Fk_Down_Id) From [Fk_Down] Where Fk_Down_Show=1"
			Rs.Open Sqlstr,Conn,1,1
			If Rs(0)<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteDownCount$}",Rs(0))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$SiteDownCount$}",0)
			End If
			Rs.Close
		End If
		SiteChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：ModuleChange
	'作    用：替换模块参数
	'参    数：
	'==============================
	Public Function ModuleChange(TemplateCode,pModuleId,pType)
		Dim TempName,TempLevelList,TempContent,TempMUrl,TempLevel,TempField,t_TempArr,t_Temp
		Sqlstr="Select Fk_Module_Name,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Menu,Fk_Module_Level,Fk_Module_LevelList,Fk_Module_MUrl,Fk_Module_PageCount,Fk_Module_Content,Fk_Module_Pic,Fk_Module_Subhead,Fk_Module_Field From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Id=" & pModuleId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempName=Rs("Fk_Module_Name")
			TempLevelList=Rs("Fk_Module_LevelList")
			TempPageSize=Rs("Fk_Module_PageCount")
			TempContent=Rs("Fk_Module_Content")
			TempField=Rs("Fk_Module_Field")
			TempLevel=Rs("Fk_Module_Level")
			TempMUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),pType,pModuleId)
			TemplateCode=ReplaceTag(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFId$}",TempLevel)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleId$}",pModuleId)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleUrl$}",TempMUrl)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleDescription$}",Rs("Fk_Module_Description"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModulePic$}",Rs("Fk_Module_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleSubhead$}",Rs("Fk_Module_Subhead"))
		Else
			Call FKFun.ShowErr("模块未找到！",0)
		End If
		Rs.Close
		If Clng(TempLevel)>0 Then
			Sqlstr="Select Fk_Module_Name From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Id=" & TempLevel
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFName$}",Rs("Fk_Module_Name"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFName$}","")
			End If
			Rs.Close
		Else
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFName$}","")
		End If
		If Instr(",0,1,2,3,4,6,7,",","&pType&",")>0 Then
			TemplateCode=ReplaceTag(TemplateCode,"{$PageCrumbs$}",GetPageCrumbs(TempLevelList,TempMUrl,TempName))
		End If
		If Instr(",3,",","&pType&",")>0 Then
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleContent$}",GetWordUrl(TempContent))
			If TempField<>"" Then
				t_TempArr=Split(TempField,"[-Fangka_Field-]")
				For Each t_Temp In t_TempArr
					TemplateCode=ReplaceTag(TemplateCode,"{$aaaaaaaaaaaaaaaaaabbbbbbField_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
				Next
			End If
		End If
		If pType=4 Then
			If TempPageSize=0 Then
				TempPageSize=Fk_Site_PageSize
			End If
		End If
		If Instr(",1,2,7,",","&pType&",")>0 Then
			If TempPageSize=0 Then
				TempPageSize=Fk_Site_PageSize
			End If
			If pType=1 Then
				Sqlstr="Select Fk_Article_Id From [Fk_ArticleList] Where Fk_Article_Show=1 And (Fk_Article_Module="&pModuleId&" Or Fk_Module_LevelList Like '%%,"&pModuleId&",%%')"
			ElseIf pType=2 Then
				Sqlstr="Select Fk_Product_Id From [Fk_ProductList] Where Fk_Product_Show=1 And (Fk_Product_Module="&pModuleId&" Or Fk_Module_LevelList Like '%%,"&pModuleId&",%%')"
			ElseIf pType=7 Then
				Sqlstr="Select Fk_Down_Id From [Fk_DownList] Where Fk_Down_Show=1 And (Fk_Down_Module="&pModuleId&" Or Fk_Module_LevelList Like '%%,"&pModuleId&",%%')"
			End If
			Rs.Open Sqlstr,Conn,1,1
			Rs.PageSize=TempPageSize
			PageCounts=Rs.PageCount
			PageAll=Rs.RecordCount
			Rs.Close
			PageFirst=""
			PagePrev=""
			PageNext=""
			PageLast=""
			If PageCounts>1 Then
				If PageNow>1 Then
					PageFirst=TempMUrl
					If PageNow=2 Then
						PagePrev=TempMUrl
					Else
						PagePrev=TempMUrl&"Index_"&(PageNow-1)&GetHtmlSuffix()
					End If
				End If
				If PageNow<PageLast Then
					PageNext=TempMUrl&"Index_"&(PageNow+1)&GetHtmlSuffix()
					PageLast=TempMUrl&"Index_"&PageLast&GetHtmlSuffix()
				End If
			End If
			TemplateCode=PageCodeChange(TemplateCode)
			TempArr=Split(FKFun.RegExpTest("\{\$ModulePageCode\(.*?\)\$\}",TemplateCode),"|-_-|")
			For Each Temp In TempArr
				If Temp<>"" Then
					If Fk_Site_Sign<>"" And Fk_Site_Html=0 Then
						TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&Fk_Site_PageSign&"{Pages}"))
					ElseIf Fk_Site_Html=1 Then
						TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),SiteDir&"Index.asp?Type="&pType&"&Module="&pModuleId&"&Page={Pages}"))
					Else
						TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&"Index_{Pages}"&GetHtmlSuffix()))
					End If
				End If
			Next
		End If
		ModuleChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：GBookPageChange
	'作    用：替换留言模块页码参数
	'参    数：
	'==============================
	Public Function GBookPageChange(TemplateCode,pModuleId)
		Dim TempMUrl
		Sqlstr="Select Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Id=" & pModuleId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempMUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
		Else
			Call FKFun.ShowErr("模块未找到！",0)
		End If
		Rs.Close
		PageFirst=""
		PagePrev=""
		PageNext=""
		PageLast=""
		If Fk_Site_HtmlType=0 Then
			TempMUrl=Replace(TempMUrl,GetHtmlSuffix(),"")
			If PageCounts>1 Then
				If PageNow>1 Then
					PageFirst=TempMUrl&GetHtmlSuffix()
					If PageNow=2 Then
						PagePrev=TempMUrl&GetHtmlSuffix()
					Else
						PagePrev=TempMUrl&"__"&(PageNow-1)&GetHtmlSuffix()
					End If
				End If
				If PageNow<PageLast Then
					PageNext=TempMUrl&"__"&(PageNow+1)&GetHtmlSuffix()
					PageLast=TempMUrl&"__"&PageLast&GetHtmlSuffix()
				End If
			End If
			TempArr=Split(FKFun.RegExpTest("\{\$ModulePageCode\(.*?\)\$\}",TemplateCode),"|-_-|")
			For Each Temp In TempArr
				If Temp<>"" Then
					TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&"__{Pages}"&GetHtmlSuffix()))
				End If
			Next
		Else
			If PageCounts>1 Then
				If PageNow>1 Then
					PageFirst=TempMUrl
					If PageNow=2 Then
						PagePrev=TempMUrl
					Else
						PagePrev=TempMUrl&"Index_"&(PageNow-1)&GetHtmlSuffix()
					End If
				End If
				If PageNow<PageLast Then
					PageNext=TempMUrl&"Index_"&(PageNow+1)&GetHtmlSuffix()
					PageLast=TempMUrl&"Index_"&PageLast&GetHtmlSuffix()
				End If
			End If
			TempArr=Split(FKFun.RegExpTest("\{\$ModulePageCode\(.*?\)\$\}",TemplateCode),"|-_-|")
			For Each Temp In TempArr
				If Temp<>"" Then
					TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&"Index_{Pages}"&GetHtmlSuffix()))
				End If
			Next
		End If
		TemplateCode=PageCodeChange(TemplateCode)
		GBookPageChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：SearchPageChange
	'作    用：替换搜索模块页码参数
	'参    数：
	'==============================
	Public Function SearchPageChange(TemplateCode)
		Dim TempMUrl
		PageFirst=""
		PagePrev=""
		PageNext=""
		PageLast=""
		TempMUrl=SiteDir&"Search/Index.asp?SearchStr="&Server.URLEncode(SearchStr)&"&SearchType="&SearchType&"&SearchTemplate="&Server.URLEncode(SearchTemplate)&"&SearchField="&Server.URLEncode(SearchField)&"&SearchFieldList="&Server.URLEncode(SearchFieldList)&"&Page="
		If PageCounts>1 Then
			If PageNow>1 Then
				PageFirst=TempMUrl&"1"
				If PageNow=2 Then
					PagePrev=TempMUrl&"1"
				Else
					PagePrev=TempMUrl&(PageNow-1)
				End If
			End If
			If PageNow<PageLast Then
				PageNext=TempMUrl&(PageNow+1)
				PageLast=TempMUrl&PageLast
			End If
		End If
		TempArr=Split(FKFun.RegExpTest("\{\$SearchPageCode\(.*?\)\$\}",TemplateCode),"|-_-|")
		For Each Temp In TempArr
			If Temp<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&"{Pages}"))
			End If
		Next
		TemplateCode=PageCodeChange(TemplateCode)
		SearchPageChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：ArticleChange
	'作    用：替换文章页参数
	'参    数：
	'==============================
	Public Function ArticleChange(TemplateCode,pId)
		Dim TempName,TempMUrl,TempContent,TempModule,TempLevelList,TempTime,t_TempArr,t_Temp,TempAdmin
		Sqlstr="Select Fk_Article_Title,Fk_Article_Keyword,Fk_Article_Description,Fk_Article_Content,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_PicList,Fk_Article_From,Fk_Article_Field,Fk_Article_Time,Fk_Article_Admin,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Pic,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_LevelList,Fk_Module_Type,Fk_Module_Level From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Article_Id=" & pId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempName=Rs("Fk_Module_Name")
			TempMUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
			TempContent=Rs("Fk_Article_Content")
			TempModule=Rs("Fk_Module_Id")
			
			TempLevelList=Rs("Fk_Module_LevelList")
			TempTime=Rs("Fk_Article_Time")
			TempAdmin=Rs("Fk_Article_Admin")
			TemplateCode=ReplaceTag(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModulePic$}",Rs("Fk_Module_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleName$}",TempName)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleUrl$}",TempMUrl)
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleId$}",pId)
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleTitle$}",CityName&Rs("Fk_Article_Title"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleKeyword$}",Rs("Fk_Article_Keyword"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleDescription$}",Rs("Fk_Article_Description"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePic$}",Rs("Fk_Article_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePicBig$}",Rs("Fk_Article_PicBig"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePicList$}",Rs("Fk_Article_PicList"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleFrom$}",Rs("Fk_Article_From"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ArticleClick$}","<span id=""Click""></span>")
			If Rs("Fk_Article_Field")<>"" Then
				t_TempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
				For Each t_Temp In t_TempArr
					TemplateCode=ReplaceTag(TemplateCode,"{$aaaaaaaaaaaaaaaaaabbbbbbField_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
				Next
			End If
		Else
			Call FKFun.ShowErr("文章未找到！",0)
		End If
		Rs.Close
		TemplateCode=ReplaceTag(TemplateCode,"{$ArticleTime$}",TempTime)
		TempArr=Split(FKFun.RegExpTest("\{\$ArticleTime\(.*?\)\$\}",TemplateCode),"|-_-|")
		For Each Temp In TempArr
			If Temp<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,Temp,ChangeTime(Split(Split(Temp,"(")(1),")")(0),TempTime))
			End If
		Next
		TemplateCode=ReplaceTag(TemplateCode,"{$PageCrumbs$}",GetPageCrumbs(TempLevelList,TempMUrl,TempName))
		TemplateCode=ReplaceTag(TemplateCode,"{$ArticleContent$}",GetWordUrl(TempContent))
		Dim ori
		If Instr(TemplateCode,"{$ArticlePrevTitle$}")>0 Or Instr(TemplateCode,"{$ArticlePrevUrl$}")>0 Or Instr(TemplateCode,"{$ArticlePrevPic$}")>0 Then
			Sqlstr="Select Fk_Article_Id,Fk_Article_Title,Fk_Module_Id,Fk_Module_Pic,Fk_Module_MUrl,Fk_Module_Type,Fk_Article_FileName,Fk_Article_Pic From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Article_Id<"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Article_Order Desc,Fk_Article_Id Desc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
			    ori = GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Article_Id"),Rs("Fk_Article_FileName"))
				ori= "/?"& City & "/" & Split(ori,"?")(1)
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevTitle$}",CityName&Rs("Fk_Article_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevUrl$}",ori)
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevPic$}",Rs("Fk_Article_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevTitle$}","无上一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticlePrevPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$ArticleNextTitle$}")>0 Or Instr(TemplateCode,"{$ArticleNextUrl$}")>0 Or Instr(TemplateCode,"{$ArticleNextPic$}")>0 Then
			Sqlstr="Select Fk_Article_Id,Fk_Article_Title,Fk_Module_Id,Fk_Module_Pic,Fk_Module_MUrl,Fk_Module_Type,Fk_Article_FileName,Fk_Article_Pic From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Article_Id>"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Article_Order Asc,Fk_Article_Id Asc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				ori = GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Article_Id"),Rs("Fk_Article_FileName"))
				ori= "/?"& City & "/" & Split(ori,"?")(1)
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextTitle$}",CityName&Rs("Fk_Article_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextUrl$}",ori)
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextPic$}",Rs("Fk_Article_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextTitle$}","无下一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleNextPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$ArticleAdmin$}")>0 Then
			Sqlstr="Select Fk_Admin_Name From [Fk_Admin] Where Fk_Admin_Id=" & TempAdmin
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleAdmin$}",Rs("Fk_Admin_Name"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ArticleAdmin$}","")
			End If
			Rs.Close
		End If
		ArticleChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：ProductChange
	'作    用：替换产品页参数
	'参    数：
	'==============================
	Public Function ProductChange(TemplateCode,pId)
		Dim ori
		Dim TempName,TempMUrl,TempContent,TempModule,TempLevelList,TempTime,t_TempArr,t_Temp,TempAdmin
		Sqlstr="Select Fk_Product_Title,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Content,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_PicList,Fk_Product_Field,Fk_Product_Time,Fk_Product_Admin,Fk_Module_Id,Fk_Module_Pic,Fk_Module_Menu,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_LevelList,Fk_Module_Type,Fk_Module_Level From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Product_Id=" & pId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempName=Rs("Fk_Module_Name")
			TempMUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
			TempTime=Rs("Fk_Product_Time")
			TempContent=Rs("Fk_Product_Content")
			TempModule=Rs("Fk_Module_Id")
			TempLevelList=Rs("Fk_Module_LevelList")
			TempAdmin=Rs("Fk_Product_Admin")
			TemplateCode=ReplaceTag(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModulePic$}",Rs("Fk_Module_Pic"))
			
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleName$}",TempName)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleUrl$}",TempMUrl)
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductId$}",pId)
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductTitle$}",CityName&Rs("Fk_Product_Title"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductKeyword$}",Rs("Fk_Product_Keyword"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductDescription$}",Rs("Fk_Product_Description"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductPic$}",Rs("Fk_Product_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductPicBig$}",Rs("Fk_Product_PicBig"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductPicList$}",Rs("Fk_Product_PicList"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ProductClick$}","<span id=""Click""></span>")
			If Rs("Fk_Product_Field")<>"" Then
				t_TempArr=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
				For Each t_Temp In t_TempArr
					TemplateCode=ReplaceTag(TemplateCode,"{$aaaaaaaaaaaaaaaaaabbbbbbField_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
				Next
			End If
		Else
			Call FKFun.ShowErr("产品未找到！",0)
		End If
		Rs.Close
		TemplateCode=ReplaceTag(TemplateCode,"{$ProductTime$}",TempTime)
		TempArr=Split(FKFun.RegExpTest("\{\$ProductTime\(.*?\)\$\}",TemplateCode),"|-_-|")
		For Each Temp In TempArr
			If Temp<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,Temp,ChangeTime(Split(Split(Temp,"(")(1),")")(0),TempTime))
			End If
		Next
		TemplateCode=ReplaceTag(TemplateCode,"{$PageCrumbs$}",GetPageCrumbs(TempLevelList,TempMUrl,TempName))
		TemplateCode=ReplaceTag(TemplateCode,"{$ProductContent$}",GetWordUrl(TempContent))
		If Instr(TemplateCode,"{$ProductPrevTitle$}")>0 Or Instr(TemplateCode,"{$ProductPrevUrl$}")>0 Or Instr(TemplateCode,"{$ProductPrevPic$}")>0 Then
			Sqlstr="Select Fk_Product_Id,Fk_Module_Pic,Fk_Product_Title,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type,Fk_Product_FileName,Fk_Product_Pic From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Product_Id<"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Product_Order Desc,Fk_Product_Id Desc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				ori=GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Product_Id"),Rs("Fk_Product_FileName"))
				ori= "/?"& City & "/" & Split(ori,"?")(1)
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevTitle$}",CityName&Rs("Fk_Product_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevUrl$}",ori)
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevPic$}",Rs("Fk_Product_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevTitle$}","无上一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductPrevPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$ProductNextTitle$}")>0 Or Instr(TemplateCode,"{$ProductNextUrl$}")>0 Or Instr(TemplateCode,"{$ProductNextPic$}")>0 Then
			Sqlstr="Select Fk_Product_Id,Fk_Module_Pic,Fk_Product_Title,Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type,Fk_Product_FileName,Fk_Product_Pic From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Product_Id>"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Product_Order Asc,Fk_Product_Id Asc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				ori=GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Product_Id"),Rs("Fk_Product_FileName"))
				ori= "/?"& City & "/" & Split(ori,"?")(1)
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextTitle$}",CityName&Rs("Fk_Product_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextUrl$}",ori)
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextPic$}",Rs("Fk_Product_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextTitle$}","无下一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductNextPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$ProductAdmin$}")>0 Then
			Sqlstr="Select Fk_Admin_Name From [Fk_Admin] Where Fk_Admin_Id=" & TempAdmin
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductAdmin$}",Rs("Fk_Admin_Name"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$ProductAdmin$}","")
			End If
			Rs.Close
		End If
		ProductChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：DownChange
	'作    用：替换下载页参数
	'参    数：
	'==============================
	Public Function DownChange(TemplateCode,pId)
		Dim TempName,TempMUrl,TempContent,TempModule,TempLevelList,TempTime,t_TempArr,t_Temp,TempAdmin
		Sqlstr="Select Fk_Down_Title,Fk_Down_Keyword,Fk_Down_Description,Fk_Down_Content,Fk_Down_Pic,Fk_Down_PicBig,Fk_Down_PicList,Fk_Down_Language,Fk_Down_System,Fk_Down_Field,Fk_Down_Time,Fk_Down_Admin,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Name,Fk_Module_Pic,Fk_Module_MUrl,Fk_Module_LevelList,Fk_Module_Type,Fk_Module_Level From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Down_Id=" & pId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempName=Rs("Fk_Module_Name")
			TempMUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
			TempTime=Rs("Fk_Down_Time")
			TempContent=Rs("Fk_Down_Content")
			TempModule=Rs("Fk_Module_Id")
			TempLevelList=Rs("Fk_Module_LevelList")
			TempAdmin=Rs("Fk_Down_Admin")
			TemplateCode=ReplaceTag(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModulePic$}",Rs("Fk_Module_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleName$}",TempName)
			TemplateCode=ReplaceTag(TemplateCode,"{$ModuleUrl$}",TempMUrl)
			TemplateCode=ReplaceTag(TemplateCode,"{$DownId$}",pId)
			TemplateCode=ReplaceTag(TemplateCode,"{$DownTitle$}",Rs("Fk_Down_Title"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownKeyword$}",Rs("Fk_Down_Keyword"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownDescription$}",Rs("Fk_Down_Description"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownPic$}",Rs("Fk_Down_Pic"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownPicBig$}",Rs("Fk_Down_PicBig"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownPicList$}",Rs("Fk_Down_PicList"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownLanguage$}",Rs("Fk_Down_Language"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownSystem$}",Rs("Fk_Down_System"))
			TemplateCode=ReplaceTag(TemplateCode,"{$DownFile$}",SiteDir&"File.asp?Id="&pId)
			TemplateCode=ReplaceTag(TemplateCode,"{$DownClick$}","<span id=""Click""></span>")
			TemplateCode=ReplaceTag(TemplateCode,"{$DownCount$}","<span id=""Count""></span>")
			If Rs("Fk_Down_Field")<>"" Then
				t_TempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
				For Each t_Temp In t_TempArr
					TemplateCode=ReplaceTag(TemplateCode,"{$aaaaaaaaaaaaaaaaaabbbbbbField_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
				Next
			End If
		Else
			Call FKFun.ShowErr("下载未找到！",0)
		End If
		Rs.Close
		TemplateCode=ReplaceTag(TemplateCode,"{$DownTime$}",TempTime)
		TempArr=Split(FKFun.RegExpTest("\{\$DownTime\(.*?\)\$\}",TemplateCode),"|-_-|")
		For Each Temp In TempArr
			If Temp<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,Temp,ChangeTime(Split(Split(Temp,"(")(1),")")(0),TempTime))
			End If
		Next
		TemplateCode=ReplaceTag(TemplateCode,"{$PageCrumbs$}",GetPageCrumbs(TempLevelList,TempMUrl,TempName))
		TemplateCode=ReplaceTag(TemplateCode,"{$DownContent$}",GetWordUrl(TempContent))
		If Instr(TemplateCode,"{$DownPrevTitle$}")>0 Or Instr(TemplateCode,"{$DownPrevUrl$}")>0 Or Instr(TemplateCode,"{$DownPrevPic$}")>0 Then
			Sqlstr="Select Fk_Down_Id,Fk_Down_Title,Fk_Module_Id,Fk_Module_Pic,Fk_Module_MUrl,Fk_Module_Type,Fk_Down_FileName,Fk_Down_Pic From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Down_Id<"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Down_Order Desc,Fk_Down_Id Desc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevTitle$}",Rs("Fk_Down_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevUrl$}",GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Down_Id"),Rs("Fk_Down_FileName")))
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevPic$}",Rs("Fk_Down_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevTitle$}","无上一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$DownPrevPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$DownNextTitle$}")>0 Or Instr(TemplateCode,"{$DownNextUrl$}")>0 Or Instr(TemplateCode,"{$DownNextPic$}")>0 Then
			Sqlstr="Select Fk_Down_Id,Fk_Down_Title,Fk_Module_Id,Fk_Module_Pic,Fk_Module_MUrl,Fk_Module_Type,Fk_Down_FileName,Fk_Down_Pic From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Down_Id>"&pId&" And Fk_Module_Id="&TempModule&" Order By Fk_Down_Order Asc,Fk_Down_Id Asc"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextTitle$}",Rs("Fk_Down_Title"))
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextUrl$}",GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Down_Id"),Rs("Fk_Down_FileName")))
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextPic$}",Rs("Fk_Down_Pic"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextTitle$}","无下一篇")
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextUrl$}","#")
				TemplateCode=ReplaceTag(TemplateCode,"{$DownNextPic$}","")
			End If
			Rs.Close
		End If
		If Instr(TemplateCode,"{$DownAdmin$}")>0 Then
			Sqlstr="Select Fk_Admin_Name From [Fk_Admin] Where Fk_Admin_Id=" & TempAdmin
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$DownAdmin$}",Rs("Fk_Admin_Name"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$DownAdmin$}","")
			End If
			Rs.Close
		End If
		DownChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：SubjectChange
	'作    用：专题页参数
	'参    数：
	'==============================
	Public Function SubjectChange(TemplateCode,pId)
		Sqlstr="Select Fk_Subject_Name From [Fk_Subject] Where Fk_Subject_Id=" & pId
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TemplateCode=ReplaceTag(TemplateCode,"{$SubjectId$}",pId)
			TemplateCode=ReplaceTag(TemplateCode,"{$SubjectName$}",Rs("Fk_Subject_Name"))
		Else
			Call FKFun.ShowErr("专题未找到！",0)
		End If
		Rs.Close
		SubjectChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：SearchChange
	'作    用：搜索页参数
	'参    数：
	'==============================
	Public Function SearchChange(TemplateCode)
		TemplateCode=ReplaceTag(TemplateCode,"{$SearchStr$}",SearchStr)
		TemplateCode=ReplaceTag(TemplateCode,"{$SearchType$}",SearchType)
		SearchChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：FileChange
	'作    用：替换模板模块参数
	'参    数：
	'TemplateCode  要处理的字符串
	'==============================
	Public Function FileChange(TemplateCode)
		Dim TemplateTemps
		'替换内部文件调用
		While Instr(TemplateCode,"{$File(")
			Temp=Split(Split(TemplateCode,"{$File(")(1),")$}")(0)
			If Fk_Site_SkinTest=1 Then
				TemplateTemps=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&Temp&".html")
			Else
				Temp=FKFun.HTMLEncode(Temp)
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Temp&"'"
				Rs.Open Sqlstr,Conn,1,1
				If Not Rs.Eof Then
					TemplateTemps=Rs("Fk_Template_Content")
				Else
					TemplateTemps=""
				End If
				Rs.Close
			End If
			TemplateCode=ReplaceTag(TemplateCode,"{$File("&Temp&")$}",TemplateTemps)
		Wend
		'自定义标签加标识
		TemplateCode=ReplaceTag(TemplateCode,"{$GBookList","{$aaaaaaaaaaaaaaaaaabbbbbbGBookList")
		TemplateCode=ReplaceTag(TemplateCode,"{$Field_","{$aaaaaaaaaaaaaaaaaabbbbbbField_")
		TemplateCode=ReplaceTag(TemplateCode,"{$FieldList_","{$aaaaaaaaaaaaaaaaaabbbbbbFieldList_")
		'替换独立信息
		While Instr(TemplateCode,"{$Info(")
			Temp=Clng(Split(Split(TemplateCode,"{$Info(")(1),")$}")(0))
			Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id="&Temp&""
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				TemplateCode=ReplaceTag(TemplateCode,"{$Info("&Temp&")$}",Rs("Fk_Info_Content"))
			Else
				TemplateCode=ReplaceTag(TemplateCode,"{$Info("&Temp&")$}","")
			End If
			Rs.Close
		Wend
		FileChange=TemplateCode
	End Function
	
'========================For标签处理===========================	

	'==============================
	'函 数 名：FkNav
	'作    用：菜单标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkNav(BCode,BPar)
		Dim NavUrl,tempFor,t_TempArr,t_Temp,ii,tRs
		Set tRs=Server.Createobject("Adodb.RecordSet")
		ListNo=1
		tempFor=""
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>3 Then
			Call FKFun.ShowErr("Nav标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(VauleTemp) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		'截取多级菜单FOR嵌套
		If Instr(BCode,"{$For(Nav")>0 Then
			t_Temp=Split(BCode,"{$For(Nav")(0)
			t_Temp=Replace(BCode,t_Temp,"",1,1)
			t_TempArr=Split(t_Temp,"{$Next$}")
			tempFor=t_TempArr(0)&"{$Next$}"
			ii=1
			While GetCount(tempFor,"{$For")<>GetCount(tempFor,"{$Next$}")
				tempFor=tempFor&t_TempArr(ii)&"{$Next$}"
				ii=ii+1
			Wend
			t_Temp=Split(tempFor,")$}")(0)
			tempFor=Right(tempFor,Len(tempFor)-Len(t_Temp)-3)
			tempFor=Left(tempFor,Len(tempFor)-8)
			BCode=Replace(BCode,t_Temp&")$}"&tempFor&"{$Next$}","{$InsertForStart$}{$FangkaFor$}{$InsertForEnd$}")
			BCode=Replace(BCode,"{$InsertForStart$}",t_Temp&")$}")
			BCode=Replace(BCode,"{$InsertForEnd$}","{$Next$}")
		End If
		'回溯操作
		If IsNumeric(VauleArr(3)) And VauleArr(3)<0 Then
			While VauleArr(3)<0 And VauleArr(1)>0
				Sqlstr="Select Fk_Module_Level From [Fk_Module] Where Fk_Module_Id="&VauleArr(1)&""
				Rs.Open Sqlstr,Conn,1,1
				If Not Rs.Eof Then
					VauleArr(1)=Rs("Fk_Module_Level")
					VauleArr(3)=VauleArr(3)+1
				Else
					VauleArr(1)=0
				End If
				Rs.Close
			Wend
		End If
		'替换操作
		Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type,Fk_Module_MUrl,Fk_Module_UrlType,Fk_Module_Url,Fk_Module_Subhead,Fk_Module_Pic,Fk_Module_Description From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_MenuShow=1 And Fk_Module_Menu="&VauleArr(0)&" And Fk_Module_Level="&VauleArr(1)&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
			If Rs("Fk_Module_Type")=5 And Rs("Fk_Module_UrlType")=0 Then
				NavUrl=Rs("Fk_Module_Url")
			Else
				NavUrl=GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))
			End If
			if City <>"" Then
						NavUrl= "/?"& City & "/" & Split(NavUrl,"?")(1)
			End If
			FkNav=FkNav&BCode
			FkNav=ReplaceTag(FkNav,"{$NavNo$}",ListNo)
			FkNav=ReplaceTag(FkNav,"{$NavId$}",Rs("Fk_Module_Id"))
			FkNav=ReplaceTag(FkNav,"{$NavName$}",Rs("Fk_Module_Name"))
			FkNav=ReplaceTag(FkNav,"{$NavDescription$}",Rs("Fk_Module_Description"))
			FkNav=ReplaceTag(FkNav,"{$NavSubhead$}",Rs("Fk_Module_Subhead"))
			FkNav=ReplaceTag(FkNav,"{$NavPic$}",Rs("Fk_Module_Pic"))
			FkNav=ReplaceTag(FkNav,"{$NavUrl$}",NavUrl)
			FkNav=ReplaceTag(FkNav,"{$NavType$}",Rs("Fk_Module_Type"))
			If Instr(FkNav,"{$NavCount$}")>0 Then
				If Rs("Fk_Module_Type")=1 Then
					Sqlstr="Select Count(Fk_Article_Id) From [Fk_ArticleList] Where Fk_Article_Show=1 And (Fk_Article_Module="&Rs("Fk_Module_Id")&" Or Fk_Module_LevelList Like '%%,"&Rs("Fk_Module_Id")&",%%')"
					tRs.Open Sqlstr,Conn,1,1
					If tRs(0)<>"" Then
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",tRs(0))
					Else
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",0)
					End If
					tRs.Close
				ElseIf Rs("Fk_Module_Type")=2 Then
					Sqlstr="Select Count(Fk_Product_Id) From [Fk_ProductList] Where Fk_Product_Show=1 And (Fk_Product_Module="&Rs("Fk_Module_Id")&" Or Fk_Module_LevelList Like '%%,"&Rs("Fk_Module_Id")&",%%')"
					tRs.Open Sqlstr,Conn,1,1
					If tRs(0)<>"" Then
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",tRs(0))
					Else
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",0)
					End If
					tRs.Close
				ElseIf Rs("Fk_Module_Type")=7 Then
					Sqlstr="Select Count(Fk_Down_Id) From [Fk_DownList] Where Fk_Down_Show=1 And (Fk_Down_Module="&Rs("Fk_Module_Id")&" Or Fk_Module_LevelList Like '%%,"&Rs("Fk_Module_Id")&",%%')"
					tRs.Open Sqlstr,Conn,1,1
					If tRs(0)<>"" Then
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",tRs(0))
					Else
						FkNav=ReplaceTag(FkNav,"{$NavCount$}",0)
					End If
					tRs.Close
				Else
					FkNav=ReplaceTag(FkNav,"{$NavCount$}",0)
				End If
			End If
			If VauleArr(2)>1 And Instr(BCode,"{$NavSub$}")>0 Then
				FkNav=ReplaceTag(FkNav,"{$NavSub$}",FkNavs(Rs("Fk_Module_Id"),Clng(VauleArr(2))-1))
			End If
			Rs.MoveNext
			ListNo=ListNo+1
		Wend
		Rs.Close
		FkNav=ReplaceTag(FkNav,"{$FangkaFor$}",tempFor)
	End Function

	'==============================
	'函 数 名：FkNavs
	'作    用：读取多级菜单操作
	'参    数：当前父ID GetId，还要读取级数GetCount
	'GetId      读取菜单父ID
	'GetCount   读取级数
	'==============================
	Private Function FkNavs(GetId,GetCount)
		Dim NavUrl,sRs
		Set sRs=Server.Createobject("Adodb.RecordSet")
		Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type,Fk_Module_MUrl,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_MenuShow=1 And Fk_Module_Level="&GetId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		sRs.Open Sqlstr,Conn,1,1
		If Not sRs.Eof Then
			FkNavs="<ul class=""sub"">" & vbCrLf
			While Not sRs.Eof
				If sRs("Fk_Module_Type")=5 And sRs("Fk_Module_UrlType")=0 Then
					NavUrl=sRs("Fk_Module_Url")
				Else
					NavUrl=GetModuleUrl(sRs("Fk_Module_MUrl"),sRs("Fk_Module_Type"),sRs("Fk_Module_Id"))
				End If
				if City <>"" Then
						NavUrl= "/?"& City & "/" & Split(NavUrl,"?")(1)
				End If
				FkNavs=FkNavs&"<li><a href="""&NavUrl&""" title="""&sRs("Fk_Module_Name")&""""
				If sRs("Fk_Module_Type")=5 Then
					FkNavs=FkNavs&" target=""_blank"""
				End If
				FkNavs=FkNavs&">"&sRs("Fk_Module_Name")&"</a>"
				If GetCount>1 Then
					FkNavs=FkNavs&FkNavs(sRs("Fk_Module_Id"),GetCount-1)
				End If
				FkNavs=FkNavs&"</li>" & vbCrLf
				sRs.MoveNext
			Wend
			FkNavs=FkNavs&"</ul>" & vbCrLf
		End If
		sRs.Close
		Set sRs=Nothing
	End Function
	
	'==============================
	'函 数 名：FkArticleList
	'作    用：文章列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkArticleList(BCode,BPar)
		Dim TempTitle,TempTitleAll,TempContent,TempTime
		Dim t_TempArr,t_Temp
		Dim time_TempArr,time_Temp
		Dim content_TempArr,content_Temp
		Dim search_TempArr,search_Temp
		ListNo=1
		ListPageNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>8 Then
			Call FKFun.ShowErr("ArticleList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(Replace(VauleTemp,"+","")) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		'截取要替换的时间和字符段
		time_TempArr=Split(FKFun.RegExpTest("\{\$ArticleListTime\(.*?\)\$\}",BCode),"|-_-|")
		content_TempArr=Split(FKFun.RegExpTest("\{\$ArticleListContent\(.*?\)\$\}",BCode),"|-_-|")
		'组合SQL
		Sqlstr="Select"
		If VauleArr(3)>0 And VauleArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(3)&""
		End If
		Sqlstr=Sqlstr&" Fk_Article_Id,Fk_Article_Title,Fk_Article_Description,Fk_Article_Color,Fk_Article_Content,Fk_Article_Field,Fk_Article_FileName,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_PicList,Fk_Article_Click,Fk_Article_Url,Fk_Article_Time,Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_Type,Fk_Module_Pic From [Fk_ArticleList] Where Fk_Article_Show=1"
		If Clng(VauleArr(0))>0 Then
			Sqlstr=Sqlstr&" And Fk_Module_Menu="&Clng(VauleArr(0))
		End If
		If Instr(VauleArr(1),"+")>0 Then
			VauleArr(1)=Replace(VauleArr(1),"+",",")
			Sqlstr=Sqlstr&" And Fk_Article_Module In ("&VauleArr(1)&")"
		ElseIf Clng(VauleArr(1))>0 Then
			Sqlstr=Sqlstr&" And (Fk_Article_Module="&VauleArr(1)&" Or Fk_Module_LevelList Like '%%,"&VauleArr(1)&",%%')"
		End If
		If VauleArr(4)=0 Then
			If Clng(VauleArr(5))>0 Then
				Sqlstr=Sqlstr&" And Fk_Article_Recommend Like '%%,"&VauleArr(5)&",%%'"
			ElseIf Clng(VauleArr(5))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Article_Recommend=',,'"
			End If
			If Clng(VauleArr(6))>0 Then
				Sqlstr=Sqlstr&" And Fk_Article_Subject Like '%%,"&VauleArr(6)&",%%'"
			ElseIf Clng(VauleArr(6))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Article_Subject=',,'"
			End If
			If Clng(VauleArr(8))=1 Then
				Sqlstr=Sqlstr&" And Fk_Article_Pic<>''"
			End If
		ElseIf SearchStr<>"" Then
			Sqlstr=Sqlstr&" And (Fk_Article_Title Like '%%"&SearchStr&"%%'"
			If SearchField<>"" Then
				search_TempArr=Split(SearchField,",")
				For Each search_Temp In search_TempArr
					If search_Temp="IsContent" Then
						Sqlstr=Sqlstr&" Or Fk_Article_Content Like '%%"&SearchStr&"%%'"
					Else
						Sqlstr=Sqlstr&" Or Fk_Article_Field Like '%%"&search_Temp&"|-Fangka_Field-|%%"&SearchStr&"%%'"
					End If
				Next
			End If
			Sqlstr=Sqlstr&")"
			If SearchFieldList<>"" Then
				search_TempArr=Split(SearchFieldList,",")
				For Each search_Temp In search_TempArr
					If Instr(search_Temp,"||")>0 Then
						Sqlstr=Sqlstr&" And Fk_Article_Field Like '%%"&Split(search_Temp,"||")(0)&"|-Fangka_Field-|%%"&Split(search_Temp,"||")(1)&"%%'"
					End If
				Next
			End If
		End If
		Select Case Clng(VauleArr(2))
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Article_Order Asc,Fk_Article_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Article_Time Desc,Fk_Article_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Article_Click Desc,Fk_Article_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Article_Order Desc,Fk_Article_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Article_Time Asc,Fk_Article_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Article_Click Asc,Fk_Article_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If VauleArr(4)=0 Then
				If Clng(VauleArr(3))>0 Then
					ReadCount=Clng(VauleArr(3))
				Else
					ReadCount=50
				End If
			Else
				If TempPageSize="" Then
					Response.Write("非列表页不能调用分页的数据，请在模板标签生成器中生成相应的不分页的For标签！")
					Rs.Close
					Response.End()
				End If
				ListPageNo=(PageNow-1)*TempPageSize+1
				ReadCount=TempPageSize
				Rs.PageSize=TempPageSize
				PageCounts=Rs.PageCount
				PageAll=Rs.RecordCount
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				Rs.AbsolutePage=PageNow
			End If
			While (Not Rs.Eof) And i<ReadCount+1
				TempTitleAll=Rs("Fk_Article_Title")
				TempContent=RemoveHTML(Rs("Fk_Article_Content"))
				TempTime=Rs("Fk_Article_Time")
				TempTitle=TempTitleAll
				If Len(TempTitle)>Clng(VauleArr(7)) And Clng(VauleArr(7))>0 Then
					TempTitle=Left(TempTitle,Clng(VauleArr(7)))&"..."
				End If
				If Rs("Fk_Article_Color")<>"" Then
					TempTitle="<span style='color:"&Rs("Fk_Article_Color")&"'>"&TempTitle&"</span>"
				End If
				FkArticleList=FkArticleList&BCode 
				FkArticleList=ReplaceTag(FkArticleList,"{$ListNo$}",ListNo)
				FkArticleList=ReplaceTag(FkArticleList,"{$ListPageNo$}",ListPageNo)
				FkArticleList=ReplaceTag(FkArticleList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ModuleListUrl$}",GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")))
				FkArticleList=ReplaceTag(FkArticleList,"{$ModulePic$}",Rs("Fk_Module_Pic"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListId$}",Rs("Fk_Article_Id"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListTitle$}",CityName&TempTitle)
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListTitleAll$}",TempTitleAll)
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListDescription$}",Rs("Fk_Article_Description"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListContent$}",Rs("Fk_Article_Content"))
				Dim ori 
				If Rs("Fk_Article_Url")<>"" Then
					ori=	Rs("Fk_Article_Url")
					if City <>"" Then
						ori= "/?"& City & "/" & Split(ori,"?")(1)
					End If
					FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListUrl$}",ori)
				Else
					ori= GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Article_Id"),Rs("Fk_Article_FileName"))
					if City <>"" Then
						ori= "/?"& City & "/" & Split(ori,"?")(1)
					End If
					FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListUrl$}",ori)
				End If
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListPic$}",Rs("Fk_Article_Pic"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListPicBig$}",Rs("Fk_Article_PicBig"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListPicList$}",Rs("Fk_Article_PicList"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListClick$}",Rs("Fk_Article_Click"))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListNew$}",DateDiff("d",TempTime,Now()))
				FkArticleList=ReplaceTag(FkArticleList,"{$ArticleListTime$}",TempTime)
				For Each content_Temp In content_TempArr
					If content_Temp<>"" Then
						FkArticleList=ReplaceTag(FkArticleList,content_Temp,Left(TempContent,Clng(Split(Split(content_Temp,"(")(1),")")(0))))
					End If
				Next
				For Each time_Temp In time_TempArr
					If time_Temp<>"" Then
						FkArticleList=ReplaceTag(FkArticleList,time_Temp,ChangeTime(Split(Split(time_Temp,"(")(1),")")(0),TempTime))
					End If
				Next
				If Rs("Fk_Article_Field")<>"" Then
					t_TempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
					For Each t_Temp In t_TempArr
						FkArticleList=ReplaceTag(FkArticleList,"{$aaaaaaaaaaaaaaaaaabbbbbbFieldList_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
					Next
				End If
				Rs.MoveNext
				ListNo=ListNo+1
				ListPageNo=ListPageNo+1
				i=i+1
			Wend
		End If
		Rs.Close
	End Function

	'==============================
	'函 数 名：FkProductList
	'作    用：产品列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkProductList(BCode,BPar)
		Dim TempTitle,TempTitleAll,TempContent,TempTime
		Dim t_TempArr,t_Temp
		Dim time_TempArr,time_Temp
		Dim content_TempArr,content_Temp
		Dim search_TempArr,search_Temp
		Dim ori
		ListNo=1
		ListPageNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>8 Then
			Call FKFun.ShowErr("ProductList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(Replace(VauleTemp,"+","")) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		'截取要替换的时间和字符段
		time_TempArr=Split(FKFun.RegExpTest("\{\$ProductListTime\(.*?\)\$\}",BCode),"|-_-|")
		content_TempArr=Split(FKFun.RegExpTest("\{\$ProductListContent\(.*?\)\$\}",BCode),"|-_-|")
		'组合SQL
		Sqlstr="Select"
		If VauleArr(3)>0 And VauleArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(3)&""
		End If
		Sqlstr=Sqlstr&" Fk_Product_Id,Fk_Product_Title,Fk_Product_Description,Fk_Product_Color,Fk_Product_Content,Fk_Product_Field,Fk_Product_FileName,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_PicList,Fk_Product_Click,Fk_Product_Url,Fk_Product_Time,Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_Type,Fk_Module_Pic From [Fk_ProductList] Where Fk_Product_Show=1"
		If Clng(VauleArr(0))>0 Then
			Sqlstr=Sqlstr&" And Fk_Module_Menu="&Clng(VauleArr(0))
		End If
		If Instr(VauleArr(1),"+")>0 Then
			VauleArr(1)=Replace(VauleArr(1),"+",",")
			Sqlstr=Sqlstr&" And Fk_Product_Module In ("&VauleArr(1)&")"
		ElseIf Clng(VauleArr(1))>0 Then
			Sqlstr=Sqlstr&" And (Fk_Product_Module="&VauleArr(1)&" Or Fk_Module_LevelList Like '%%,"&VauleArr(1)&",%%')"
		End If
		If VauleArr(4)=0 Then
			If Clng(VauleArr(5))>0 Then
				Sqlstr=Sqlstr&" And Fk_Product_Recommend Like '%%,"&VauleArr(5)&",%%'"
			ElseIf Clng(VauleArr(5))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Product_Recommend=',,'"
			End If
			If Clng(VauleArr(6))>0 Then
				Sqlstr=Sqlstr&" And Fk_Product_Subject Like '%%,"&VauleArr(6)&",%%'"
			ElseIf Clng(VauleArr(6))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Product_Subject=',,'"
			End If
			If Clng(VauleArr(8))=1 Then
				Sqlstr=Sqlstr&" And Fk_Product_Pic<>''"
			End If
		ElseIf SearchStr<>"" Then
			Sqlstr=Sqlstr&" And (Fk_Product_Title Like '%%"&SearchStr&"%%'"
			If SearchField<>"" Then
				search_TempArr=Split(SearchField,",")
				For Each search_Temp In search_TempArr
					If search_Temp="IsContent" Then
						Sqlstr=Sqlstr&" Or Fk_Product_Content Like '%%"&SearchStr&"%%'"
					Else
						Sqlstr=Sqlstr&" Or Fk_Product_Field Like '%%"&search_Temp&"|-Fangka_Field-|%%"&SearchStr&"%%'"
					End If
				Next
			End If
			Sqlstr=Sqlstr&")"
			If SearchFieldList<>"" Then
				search_TempArr=Split(SearchFieldList,",")
				For Each search_Temp In search_TempArr
					If Instr(search_Temp,"||")>0 Then
						Sqlstr=Sqlstr&" And Fk_Product_Field Like '%%"&Split(search_Temp,"||")(0)&"|-Fangka_Field-|%%"&Split(search_Temp,"||")(1)&"%%'"
					End If
				Next
			End If
		End If
		Select Case VauleArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Product_Order Asc,Fk_Product_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Product_Time Desc,Fk_Product_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Product_Click Desc,Fk_Product_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Product_Order Desc,Fk_Product_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Product_Time Asc,Fk_Product_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Product_Click Asc,Fk_Product_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If VauleArr(4)=0 Then
				If Clng(VauleArr(3))>0 Then
					ReadCount=Clng(VauleArr(3))
				Else
					ReadCount=50
				End If
			Else
				If TempPageSize="" Then
					Response.Write("非列表页不能调用分页的数据，请在模板标签生成器中生成相应的不分页的For标签！")
					Rs.Close
					Response.End()
				End If
				ListPageNo=(PageNow-1)*TempPageSize+1
				ReadCount=TempPageSize
				Rs.PageSize=TempPageSize
				PageCounts=Rs.PageCount
				PageAll=Rs.RecordCount
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				Rs.AbsolutePage=PageNow
			End If
			While (Not Rs.Eof) And i<ReadCount+1
				TempTitleAll=Rs("Fk_Product_Title")
				TempContent=RemoveHTML(Rs("Fk_Product_Content"))
				TempTime=Rs("Fk_Product_Time")
				TempTitle=TempTitleAll
				If Len(TempTitleAll)>Clng(VauleArr(7)) And Clng(VauleArr(7))>0 Then
					TempTitle=Left(TempTitleAll,Clng(VauleArr(7)))&"..."
				End If
				If Rs("Fk_Product_Color")<>"" Then
					TempTitle="<span style='color:"&Rs("Fk_Product_Color")&"'>"&TempTitle&"</span>"
				End If
				FkProductList=FkProductList&BCode
				FkProductList=ReplaceTag(FkProductList,"{$ListNo$}",ListNo)
				FkProductList=ReplaceTag(FkProductList,"{$ListPageNo$}",ListPageNo)
				FkProductList=ReplaceTag(FkProductList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
				FkProductList=ReplaceTag(FkProductList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
				FkProductList=ReplaceTag(FkProductList,"{$ModuleListUrl$}",GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")))
				FkProductList=ReplaceTag(FkProductList,"{$ModulePic$}",Rs("Fk_Module_Pic"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListId$}",Rs("Fk_Product_Id"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListTitle$}",CityName&TempTitle)
				FkProductList=ReplaceTag(FkProductList,"{$ProductListTitleAll$}",TempTitleAll)
				FkProductList=ReplaceTag(FkProductList,"{$ProductListDescription$}",Rs("Fk_Product_Description"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListContent$}",Rs("Fk_Product_Content"))
				If Rs("Fk_Product_Url")<>"" Then
					ori=Rs("Fk_Product_Url")
					if City <>"" Then
						ori= "/?"& City & "/" & Split(ori,"?")(1)
					End If
					FkProductList=ReplaceTag(FkProductList,"{$ProductListUrl$}",ori)
				Else
					
					ori=GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Product_Id"),Rs("Fk_Product_FileName"))
					if City <>"" Then
						ori= "/?"& City & "/" & Split(ori,"?")(1)
					End If
					FkProductList=ReplaceTag(FkProductList,"{$ProductListUrl$}",ori)
					
				End If
				FkProductList=ReplaceTag(FkProductList,"{$ProductListPic$}",Rs("Fk_Product_Pic"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListPicBig$}",Rs("Fk_Product_PicBig"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListPicList$}",Rs("Fk_Product_PicList"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListClick$}",Rs("Fk_Product_Click"))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListNew$}",DateDiff("d",TempTime,Now()))
				FkProductList=ReplaceTag(FkProductList,"{$ProductListTime$}",TempTime)
				For Each content_Temp In content_TempArr
					If content_Temp<>"" Then
						FkProductList=ReplaceTag(FkProductList,content_Temp,Left(TempContent,Clng(Split(Split(content_Temp,"(")(1),")")(0))))
					End If
				Next
				For Each time_Temp In time_TempArr
					If time_Temp<>"" Then
						FkProductList=ReplaceTag(FkProductList,time_Temp,ChangeTime(Split(Split(time_Temp,"(")(1),")")(0),TempTime))
					End If
				Next
				If Rs("Fk_Product_Field")<>"" Then
					t_TempArr=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
					For Each t_Temp In t_TempArr
						FkProductList=ReplaceTag(FkProductList,"{$aaaaaaaaaaaaaaaaaabbbbbbFieldList_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
					Next
				End If
				Rs.MoveNext
				ListNo=ListNo+1
				ListPageNo=ListPageNo+1
				i=i+1
			Wend
		End If
		Rs.Close
	End Function

	'==============================
	'函 数 名：FkDownList
	'作    用：下载列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkDownList(BCode,BPar)
		Dim TempTitle,TempTitleAll,TempContent,TempTime
		Dim t_TempArr,t_Temp
		Dim time_TempArr,time_Temp
		Dim content_TempArr,content_Temp
		Dim search_TempArr,search_Temp
		ListNo=1
		ListPageNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>8 Then
			Call FKFun.ShowErr("DownList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(Replace(VauleTemp,"+","")) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		'截取要替换的时间和字符段
		time_TempArr=Split(FKFun.RegExpTest("\{\$DownListTime\(.*?\)\$\}",BCode),"|-_-|")
		content_TempArr=Split(FKFun.RegExpTest("\{\$DownListContent\(.*?\)\$\}",BCode),"|-_-|")
		'组合SQL
		Sqlstr="Select"
		If VauleArr(3)>0 And VauleArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(3)&""
		End If
		Sqlstr=Sqlstr&" Fk_Down_Id,Fk_Down_Title,Fk_Down_Description,Fk_Down_Color,Fk_Down_Content,Fk_Down_Field,Fk_Down_FileName,Fk_Down_Pic,Fk_Down_PicBig,Fk_Down_PicList,Fk_Down_Click,Fk_Down_System,Fk_Down_Language,Fk_Down_Url,Fk_Down_Time,Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_Type,Fk_Module_Pic From [Fk_DownList] Where Fk_Down_Show=1"
		If Clng(VauleArr(0))>0 Then
			Sqlstr=Sqlstr&" And Fk_Module_Menu=" & Clng(VauleArr(0))
		End If
		If Instr(VauleArr(1),"+")>0 Then
			VauleArr(1)=Replace(VauleArr(1),"+",",")
			Sqlstr=Sqlstr&" And Fk_Down_Module In ("&VauleArr(1)&")"
		ElseIf Clng(VauleArr(1))>0 Then
			Sqlstr=Sqlstr&" And (Fk_Down_Module="&VauleArr(1)&" Or Fk_Module_LevelList Like '%%,"&VauleArr(1)&",%%')"
		End If
		If VauleArr(4)=0 Then
			If Clng(VauleArr(5))>0 Then
				Sqlstr=Sqlstr&" And Fk_Down_Recommend Like '%%,"&VauleArr(5)&",%%'"
			ElseIf Clng(VauleArr(5))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Down_Recommend=',,'"
			End If
			If Clng(VauleArr(6))>0 Then
				Sqlstr=Sqlstr&" And Fk_Down_Subject Like '%%,"&VauleArr(6)&",%%'"
			ElseIf Clng(VauleArr(6))=-1 Then
				Sqlstr=Sqlstr&" And Fk_Down_Subject=',,'"
			End If
			If Clng(VauleArr(8))=1 Then
				Sqlstr=Sqlstr&" And Fk_Down_Pic<>''"
			End If
		ElseIf SearchStr<>"" Then
			Sqlstr=Sqlstr&" And (Fk_Down_Title Like '%%"&SearchStr&"%%'"
			If SearchField<>"" Then
				search_TempArr=Split(SearchField,",")
				For Each search_Temp In search_TempArr
					If search_Temp="IsContent" Then
						Sqlstr=Sqlstr&" Or Fk_Down_Content Like '%%"&SearchStr&"%%'"
					Else
						Sqlstr=Sqlstr&" Or Fk_Down_Field Like '%%"&search_Temp&"|-Fangka_Field-|%%"&SearchStr&"%%'"
					End If
				Next
			End If
			Sqlstr=Sqlstr&")"
			If SearchFieldList<>"" Then
				search_TempArr=Split(SearchFieldList,",")
				For Each search_Temp In search_TempArr
					If Instr(search_Temp,"||")>0 Then
						Sqlstr=Sqlstr&" And Fk_Down_Field Like '%%"&Split(search_Temp,"||")(0)&"|-Fangka_Field-|%%"&Split(search_Temp,"||")(1)&"%%'"
					End If
				Next
			End If
		End If
		Select Case VauleArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Down_Order Asc,Fk_Down_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Down_Time Desc,Fk_Down_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Down_Click Desc,Fk_Down_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Down_Order Desc,Fk_Down_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Down_Time Asc,Fk_Down_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Down_Click Asc,Fk_Down_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If VauleArr(4)=0 Then
				If Clng(VauleArr(3))>0 Then
					ReadCount=Clng(VauleArr(3))
				Else
					ReadCount=50
				End If
			Else
				If TempPageSize="" Then
					Response.Write("非列表页不能调用分页的数据，请在模板标签生成器中生成相应的不分页的For标签！")
					Rs.Close
					Response.End()
				End If
				ListPageNo=(PageNow-1)*TempPageSize+1
				ReadCount=TempPageSize
				Rs.PageSize=TempPageSize
				PageCounts=Rs.PageCount
				PageAll=Rs.RecordCount
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				Rs.AbsolutePage=PageNow
			End If
			While (Not Rs.Eof) And i<ReadCount+1
				TempTitleAll=Rs("Fk_Down_Title")
				TempContent=RemoveHTML(Rs("Fk_Down_Content"))
				TempTime=Rs("Fk_Down_Time")
				TempTitle=TempTitleAll
				If Len(TempTitleAll)>Clng(VauleArr(7)) And Clng(VauleArr(7))>0 Then
					TempTitle=Left(TempTitleAll,Clng(VauleArr(7)))&"..."
				End If
				If Rs("Fk_Down_Color")<>"" Then
					TempTitle="<span style='color:"&Rs("Fk_Down_Color")&"'>"&TempTitle&"</span>"
				End If
				FkDownList=FkDownList&BCode
				FkDownList=ReplaceTag(FkDownList,"{$ListNo$}",ListNo)
				FkDownList=ReplaceTag(FkDownList,"{$ListPageNo$}",ListPageNo)
				FkDownList=ReplaceTag(FkDownList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
				FkDownList=ReplaceTag(FkDownList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
				FkDownList=ReplaceTag(FkDownList,"{$ModuleListUrl$}",GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")))
				FkDownList=ReplaceTag(FkDownList,"{$ModulePic$}",Rs("Fk_Module_Pic"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListId$}",Rs("Fk_Down_Id"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListTitle$}",TempTitle)
				FkDownList=ReplaceTag(FkDownList,"{$DownListTitleAll$}",TempTitleAll)
				FkDownList=ReplaceTag(FkDownList,"{$DownListDescription$}",Rs("Fk_Down_Description"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListContent$}",Rs("Fk_Down_Content"))
				If Rs("Fk_Down_Url")<>"" Then
					FkDownList=ReplaceTag(FkDownList,"{$DownListUrl$}",Rs("Fk_Down_Url"))
				Else
					FkDownList=ReplaceTag(FkDownList,"{$DownListUrl$}",GetContentUrl(GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id")),Rs("Fk_Down_Id"),Rs("Fk_Down_FileName")))
				End If
				FkDownList=ReplaceTag(FkDownList,"{$DownListPic$}",Rs("Fk_Down_Pic"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListPicBig$}",Rs("Fk_Down_PicBig"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListPicList$}",Rs("Fk_Down_PicList"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListClick$}",Rs("Fk_Down_Click"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListSystem$}",Rs("Fk_Down_System"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListLanguage$}",Rs("Fk_Down_Language"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))
				FkDownList=ReplaceTag(FkDownList,"{$DownListNew$}",DateDiff("d",TempTime,Now()))
				FkDownList=ReplaceTag(FkDownList,"{$DownListTime$}",TempTime)
				For Each content_Temp In content_TempArr
					If content_Temp<>"" Then
						FkDownList=ReplaceTag(FkDownList,content_Temp,Left(TempContent,Clng(Split(Split(content_Temp,"(")(1),")")(0))))
					End If
				Next
				For Each time_Temp In time_TempArr
					If time_Temp<>"" Then
						FkDownList=ReplaceTag(FkDownList,time_Temp,ChangeTime(Split(Split(time_Temp,"(")(1),")")(0),TempTime))
					End If
				Next
				If Rs("Fk_Down_Field")<>"" Then
					t_TempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
					For Each t_Temp In t_TempArr
						FkDownList=ReplaceTag(FkDownList,"{$aaaaaaaaaaaaaaaaaabbbbbbFieldList_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
					Next
				End If
				Rs.MoveNext
				ListNo=ListNo+1
				ListPageNo=ListPageNo+1
				i=i+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkGBookList
	'作    用：留言列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkGBookList(BCode,BPar)
		Dim t_TempArr,t_Temp
		dim t_Fk_Module_MUrl,t_Fk_Module_Type
		ListNo=1
		ListPageNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>4 Then
			Call FKFun.ShowErr("GBookList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(VauleTemp) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		'组合SQL
		Sqlstr="Select Fk_Module_MUrl,Fk_Module_Type From [Fk_Module] Where Fk_Module_Type=4 And Fk_Module_Id="&VauleArr(1)
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			t_Fk_Module_MUrl=Rs("Fk_Module_MUrl")
			t_Fk_Module_Type=Rs("Fk_Module_Type")
		End If
		Rs.Close
		Sqlstr="Select"
		If VauleArr(2)>0 And VauleArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(2)&""
		End If
		Sqlstr=Sqlstr&" Fk_GBook_Id,Fk_GBook_Content,Fk_GBook_Time,Fk_GBook_ReContent,Fk_GBook_ReTime From [Fk_GBook] Where Fk_GBook_Module=" & VauleArr(1)
		If VauleArr(3)=1 Then
			Sqlstr=Sqlstr&" And Fk_GBook_ReContent<>''"
		End If
		Sqlstr=Sqlstr&" Order By Fk_GBook_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If VauleArr(4)=0 Then
				If Clng(VauleArr(2))>0 Then
					ReadCount=Clng(VauleArr(2))
				Else
					ReadCount=50
				End If
			Else
				If TempPageSize="" Then
					Response.Write("非列表页不能调用分页的数据，请在模板标签生成器中生成相应的不分页的For标签！")
					Rs.Close
					Response.End()
				End If
				ListPageNo=(PageNow-1)*TempPageSize+1
				ReadCount=TempPageSize
				Rs.PageSize=TempPageSize
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				PageAll=Rs.RecordCount
				Rs.AbsolutePage=PageNow
				Rs.PageSize=TempPageSize
			End If
			While (Not Rs.Eof) And i<ReadCount+1
				FkGBookList=FkGBookList&BCode
				FkGBookList=Replace(FkGBookList,"{$ListNo$}",ListNo)
				FkGBookList=Replace(FkGBookList,"{$ListPageNo$}",ListPageNo)
				t_TempArr=Split(Rs("Fk_GBook_Content"),"|-_-|Fangka|-_-|")
				For Each t_Temp In t_TempArr
					FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookList"&Split(t_Temp,"|-^-|Fangka|-^-|")(2)&"$}",Split(t_Temp,"|-^-|Fangka|-^-|")(1))
				Next
				FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookListTime$}",Rs("Fk_GBook_Time"))
				If Rs("Fk_GBook_ReContent")<>"" Then
					FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookListReContent$}",Rs("Fk_GBook_ReContent"))
					FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookListReTime$}",Rs("Fk_GBook_ReTime"))
				Else
					FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookListReContent$}","待回复")
					FkGBookList=ReplaceTag(FkGBookList,"{$aaaaaaaaaaaaaaaaaabbbbbbGBookListReTime$}","")
				End If
				Rs.MoveNext
				ListNo=ListNo+1
				ListPageNo=ListPageNo+1
				i=i+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkFriendsList
	'作    用：友情链接列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkFriendsList(BCode,BPar)
		ListNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>2 Then
			Call FKFun.ShowErr("FriendsList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(VauleTemp) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		Sqlstr="Select"
		If VauleArr(2)>0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(2)&""
		End If
		Sqlstr=Sqlstr&" Fk_Friends_Name,Fk_Friends_Url,Fk_Friends_About,Fk_Friends_Logo From [Fk_Friends] Where 1=1"
		If VauleArr(0)>0 Then
			Sqlstr=Sqlstr&" And Fk_Friends_FriendsType="&VauleArr(0)&""
		End If
		If VauleArr(1)=1 Then
			Sqlstr=Sqlstr&" And Fk_Friends_ShowType=1"
		Else
			Sqlstr=Sqlstr&" And Fk_Friends_ShowType=2"
		End If
		Sqlstr=Sqlstr&" Order By Fk_Friends_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				FkFriendsList=FkFriendsList&BCode
				FkFriendsList=ReplaceTag(FkFriendsList,"{$ListNo$}",ListNo)
				FkFriendsList=ReplaceTag(FkFriendsList,"{$FriendsName$}",Rs("Fk_Friends_Name"))
				FkFriendsList=ReplaceTag(FkFriendsList,"{$FriendsUrl$}",Rs("Fk_Friends_Url"))
				FkFriendsList=ReplaceTag(FkFriendsList,"{$FriendsAbout$}",Rs("Fk_Friends_About"))
				FkFriendsList=ReplaceTag(FkFriendsList,"{$FriendsLogo$}",Rs("Fk_Friends_Logo"))
				Rs.MoveNext
				ListNo=ListNo+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkJobList
	'作    用：招聘列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkJobList(BCode,BPar)
		Dim t_TempArr,t_Temp
		ListNo=1
		'判断合法性
		VauleArr=Split(BPar,"/")
		If UBound(VauleArr)<>3 Then
			Call FKFun.ShowErr("JobList标签参数个数不正常！",0)
		End If
		For Each VauleTemp In VauleArr
			If Not IsNumeric(VauleTemp) Then
				Call FKFun.ShowErr("出错了：“"&BPar&"”出现非数字参数！",0)
			End If
		Next
		Sqlstr="Select"
		If VauleArr(2)>0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(0)&""
		End If
		Sqlstr=Sqlstr&" Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date,Fk_Job_Time,Fk_Job_Field From [Fk_Job] Where Fk_Job_Module=" & VauleArr(1)
		If VauleArr(3)=1 Then
			Sqlstr=Sqlstr&" And DateAdd('d',Fk_Job_Date,Fk_Job_Time)<=#"&Now()&"#"
		End If
		If VauleArr(3)=2 Then
			Sqlstr=Sqlstr&" And DateAdd('d',Fk_Job_Date,Fk_Job_Time)>#"&Now()&"#"
		End If
		Sqlstr=Sqlstr&" Order By Fk_Job_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				FkJobList=FkJobList&BCode
				FkJobList=ReplaceTag(FkJobList,"{$ListNo$}",ListNo)
				FkJobList=ReplaceTag(FkJobList,"{$JobName$}",Rs("Fk_Job_Name"))
				FkJobList=ReplaceTag(FkJobList,"{$JobCount$}",Rs("Fk_Job_Count"))
				FkJobList=ReplaceTag(FkJobList,"{$JobAbout$}",Rs("Fk_Job_About"))
				FkJobList=ReplaceTag(FkJobList,"{$JobArea$}",Rs("Fk_Job_Area"))
				If Rs("Fk_Job_Date")=0 Then
					FkJobList=ReplaceTag(FkJobList,"{$JobDate$}","长期有效")
				Else
					FkJobList=ReplaceTag(FkJobList,"{$JobDate$}",Rs("Fk_Job_Date")&"天")
				End If
				FkJobList=ReplaceTag(FkJobList,"{$JobTime$}",Rs("Fk_Job_Time"))
				If Rs("Fk_Job_Field")<>"" Then
					t_TempArr=Split(Rs("Fk_Job_Field"),"[-Fangka_Field-]")
					For Each t_Temp In t_TempArr
						FkJobList=ReplaceTag(FkJobList,"{$aaaaaaaaaaaaaaaaaabbbbbbFieldList_"&Split(t_Temp,"|-Fangka_Field-|")(0)&"$}",Split(t_Temp,"|-Fangka_Field-|")(1))
					Next
				End If
				Rs.MoveNext
				ListNo=ListNo+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkSubjectList
	'作    用：专题列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkSubjectList(BCode,BPar)
		Dim SubjectUrl
		ListNo=1
		VauleArr=Split(BPar,"/")
		Sqlstr="Select"
		If Clng(VauleArr(0))>0 Then
			Sqlstr=Sqlstr&" Top "&VauleArr(0)&""
		End If
		Sqlstr=Sqlstr&" Fk_Subject_Id,Fk_Subject_Name,Fk_Subject_Pic,Fk_Subject_Dir From [Fk_Subject] Order By Fk_Subject_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				If Fk_Site_Html=0 Then
					SubjectUrl=SiteDir&"Subject/Index.asp?Id="&Rs("Fk_Subject_Id")
				Else
					SubjectUrl=SiteDir&"Subject/"&Rs("Fk_Subject_Dir")&GetHtmlSuffix()
				End If
				FkSubjectList=FkSubjectList&BCode
				FkSubjectList=ReplaceTag(FkSubjectList,"{$ListNo$}",ListNo)
				FkSubjectList=ReplaceTag(FkSubjectList,"{$SubjectListName$}",Rs("Fk_Subject_Name"))
				FkSubjectList=ReplaceTag(FkSubjectList,"{$SubjectListPic$}",Rs("Fk_Subject_Pic"))
				FkSubjectList=ReplaceTag(FkSubjectList,"{$SubjectListUrl$}",SubjectUrl)
				Rs.MoveNext
				ListNo=ListNo+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkPicList
	'作    用：题图列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkPicList(BCode,BPar)
		Dim ListNo,s_TempArr,s_Temp,PicTextTemp
		ListNo=1
		If BPar="" Or Instr(BPar,"||")=0 Then
			Exit Function
		End If
		s_TempArr=Split(BPar,"|-_-|")
		For Each s_Temp In s_TempArr
			If s_Temp<>"" And Instr(s_Temp,"||")>0 Then
				If UBound(Split(s_Temp,"||"))>=2 Then
					PicTextTemp=Split(s_Temp,"||")(2)
				Else
					PicTextTemp=""
				End If
				FkPicList=FkPicList&BCode
				FkPicList=ReplaceTag(FkPicList,"{$ListNo$}",ListNo)
				FkPicList=ReplaceTag(FkPicList,"{$PicListNo$}",ListNo)
				FkPicList=ReplaceTag(FkPicList,"{$PicListSmall$}",Split(s_Temp,"||")(0))
				FkPicList=ReplaceTag(FkPicList,"{$PicListBig$}",Split(s_Temp,"||")(1))
				FkPicList=ReplaceTag(FkPicList,"{$PicListText$}",PicTextTemp)
				ListNo=ListNo+1
			End If
		Next
	End Function
	
	'==============================
	'函 数 名：FkNumList
	'作    用：循环列表标签操作
	'参    数：
	'BCode  标签内容
	'BPar   标签参数
	'==============================
	Private Function FkNumList(BCode,BPar)
		Dim ListNo
		ListNo=1
		If BPar="" Or Instr(BPar,"/")=0 Then
			Exit Function
		End If
		TempArr=Split(BPar,"/")
		If Not IsNumeric(TempArr(0)) Or Not IsNumeric(TempArr(1)) Then
			Exit Function
		End If
		For ListNo=Clng(TempArr(0)) To Clng(TempArr(1))
			FkNumList=FkNumList&BCode
			FkNumList=Replace(FkNumList,"{$ListNo$}",ListNo)
		Next
	End Function

'========================其他函数区===========================
	'==============================
	'函 数 名：GetPageCrumbs
	'作    用：输出面包屑菜单
	'参    数：
	'==============================
	Private Function GetPageCrumbs(ModuleLevelList,ModuleUrl,ModuleName)
		If Instr(ModuleLevelList,",")>0 Then
			TempArr=Split(ModuleLevelList,",")
			GetPageCrumbs=""
			For Each Temp In TempArr
				If Temp<>"" Then
					Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_MUrl,Fk_Module_Type,Fk_Module_UrlType,Fk_Module_Url From [Fk_Module] Where Fk_Module_Id=" & Temp
					Rs.Open Sqlstr,Conn,1,1
					If Not Rs.Eof Then
						If Rs("Fk_Module_Type")=5 And Rs("Fk_Module_UrlType")=0 Then
							GetPageCrumbs=GetPageCrumbs&"&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&Rs("Fk_Module_Url")&""" title="""&Rs("Fk_Module_Name")&""">"&Rs("Fk_Module_Name")&"</a>"
						Else
							GetPageCrumbs=GetPageCrumbs&"&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&GetModuleUrl(Rs("Fk_Module_MUrl"),Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&""" title="""&Rs("Fk_Module_Name")&""">"&Rs("Fk_Module_Name")&"</a>"
						End If
					End If
					Rs.Close
				End If
			Next
		End If
		GetPageCrumbs=GetPageCrumbs&"&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&ModuleUrl&""" title="""&ModuleName&""">"&ModuleName&"</a>"
	End Function

	'==============================
	'函 数 名：GetWordUrl
	'作    用：替换站内链接
	'参    数：
	'==============================
	Private Function GetWordUrl(Str)
		If Str<>"" Then
			Sqlstr="Select Fk_Word_Name,Fk_Word_Url,Fk_Word_RNum From [Fk_Word] Order By Fk_Word_Id Desc"
			Rs.Open Sqlstr,Conn,1,1
			While Not Rs.Eof
				If Clng(Rs("Fk_Word_RNum"))>0 Then
					Str=Replace(Str,Rs("Fk_Word_Name"),"<a href="""&Rs("Fk_Word_Url")&""" target=""_blank"" title="""&Rs("Fk_Word_Name")&""">"&Rs("Fk_Word_Name")&"</a>",1,Rs("Fk_Word_RNum"))
				Else
					Str=Replace(Str,Rs("Fk_Word_Name"),"<a href="""&Rs("Fk_Word_Url")&""" target=""_blank"" title="""&Rs("Fk_Word_Name")&""">"&Rs("Fk_Word_Name")&"</a>")
				End If
				Rs.MoveNext
			Wend
			Rs.Close
			GetWordUrl=Str
		Else
			GetWordUrl="内容建设中"
		End If
	End Function

	'==============================
	'函 数 名：PageCodeChange
	'作    用：页码参数参数
	'参    数：
	'==============================
	Private Function PageCodeChange(Str)
		Str=Replace(Str,"{$PageFirst$}",PageFirst)
		Str=Replace(Str,"{$PagePrev$}",PagePrev)
		Str=Replace(Str,"{$PageNext$}",PageNext)
		Str=Replace(Str,"{$PageLast$}",PageLast)
		Str=Replace(Str,"{$PageNow$}",PageNow)
		Str=Replace(Str,"{$PageCount$}",PageCounts)
		Str=Replace(Str,"{$PageRecordCount$}",PageAll)
		Str=Replace(Str,"{$PageSize$}",TempPageSize)
		PageCodeChange=Str
	End Function

	'==============================
	'函 数 名：ReplaceTag
	'作    用：替换标签
	'参    数：
	'==============================
	Private Function ReplaceTag(MyStr,MyTag,MyChang)
		If IsNull(MyChang) Then
			ReplaceTag=Replace(MyStr,MyTag,"")
		Else
			ReplaceTag=Replace(MyStr,MyTag,MyChang)
		End If
	End Function

	'==============================
	'函 数 名：GetHtmlSuffix
	'作    用：获取生成后缀
	'参    数：
	'==============================
	Public Function GetHtmlSuffix()
		Select Case Fk_Site_HtmlSuffix
			Case 0
				GetHtmlSuffix=".html"
			Case 1
				GetHtmlSuffix=".htm"
			Case 2
				GetHtmlSuffix=".shtml"
			Case 3
				GetHtmlSuffix=".xml"
		End Select
	End Function

	'==============================
	'函 数 名：ChangeTime
	'作    用：处理时间
	'参    数：
	'==============================
	Private Function ChangeTime(str,sTime)
		str=Replace(str,"yyyy",Year(sTime))
		str=Replace(str,"yy",Right(Year(sTime),2))
		str=Replace(str,"mm",Month(sTime))
		str=Replace(str,"dd",Day(sTime))
		str=Replace(str,"hh",Hour(sTime))
		str=Replace(str,"nn",Minute(sTime))
		str=Replace(str,"ss",Second(sTime))
		ChangeTime=str
	End Function
	
	'==============================
	'函 数 名：GetModuleUrl
	'作    用：获取模块完整地址
	'参    数：
	'm_MUrl  模块路径
	'm_Type  模块类型
	'm_Id    模块ID
	'==============================
	Public Function GetModuleUrl(m_MUrl,m_Type,m_Id)
		Dim sRs
		If m_Type=8 Then
			Set sRs=Server.Createobject("Adodb.RecordSet")
			Sqlstr="Select Top 1 Fk_Module_Id,Fk_Module_MUrl,Fk_Module_Type From [Fk_Module] Where Fk_Module_Level="&m_Id&" And Fk_Module_Type<>8 Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			sRs.Open Sqlstr,Conn,1,1
			If Not sRs.Eof Then
				GetModuleUrl=GetModuleUrl(sRs("Fk_Module_MUrl"),sRs("Fk_Module_Type"),sRs("Fk_Module_Id"))
			Else
				GetModuleUrl="#"
			End If
			sRs.Close
			Set sRs=Nothing
			Exit Function
		ElseIf Fk_Site_Html=1 Then
			GetModuleUrl=SiteDir&"Index.asp?Type="&m_Type&"&Module="&m_Id
			Exit Function
		Else
			GetModuleUrl=m_MUrl
		End If
		If Fk_Site_Html=0 Then
			If Fk_Site_Sign<>"" Then
				GetModuleUrl=Replace(GetModuleUrl,"/",Fk_Site_Sign)
				If Right(GetModuleUrl,1)=Fk_Site_Sign Then
					GetModuleUrl=Left(GetModuleUrl,Len(GetModuleUrl)-1)
				End If
			End If
			GetModuleUrl=UrlIndexIn&"?"&GetModuleUrl
		End If
		GetModuleUrl=SiteDir&GetModuleUrl
	End Function
	
	'==============================
	'函 数 名：GetContentUrl
	'作    用：获取内容页完整地址
	'参    数：
	'm_MUrl        模块路径
	'm_Id          内容ID
	'm_FileName    内容文件名
	'==============================
	Public Function GetContentUrl(m_MUrl,m_Id,m_FileName)
		If Fk_Site_Html=1 Then
			GetContentUrl=m_MUrl&"&Id="&m_Id
			Exit Function
		Else
			If Fk_Site_Sign<>"" And Fk_Site_Html=0 Then
				GetContentUrl=Fk_Site_Sign
			Else
				GetContentUrl=""
			End If
			If m_FileName<>"" Then
				GetContentUrl=m_MUrl&GetContentUrl&m_FileName&GetHtmlSuffix()
			Else
				GetContentUrl=m_MUrl&GetContentUrl&m_Id&GetHtmlSuffix()
			End If
		End If
	End Function
	
	'==============================
	'函 数 名：GetTemplate
	'作    用：获取模板代码
	'参    数：
	's_FileName      获取的模板文件
	's_TemplateId    获取类型
	's_IsIndex       是否菜单首页
	's_MenuTepmlate  菜单模板目录
	'==============================
	Public Function GetTemplate(s_FileName,s_TemplateId,s_IsIndex,s_MenuTepmlate)
		Dim TempFileName
		If s_MenuTepmlate<>"" Then
			s_MenuTepmlate=s_MenuTepmlate&"/"
		End If
		If s_IsIndex=1 And s_MenuTepmlate<>"" Then  '子菜单中的首页模块
			If Fk_Site_SkinTest=1 Then
				GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&s_MenuTepmlate&"index.html")
				Exit Function
			End If
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&s_MenuTepmlate&"index'"
		ElseIf s_TemplateId=0 Then  '默认模板
			If Fk_Site_SkinTest=1 Then
				GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&s_MenuTepmlate&s_FileName&".html")
				Exit Function
			End If
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&s_MenuTepmlate&s_FileName&"'"
		Else  '自定义模板
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Id=" & s_TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempFileName=Rs("Fk_Template_Name")
			GetTemplate=Rs("Fk_Template_Content")
		Else
			Call FKFun.ShowErr("模板未找到！",0)
		End If
		Rs.Close
		If Fk_Site_SkinTest=1 Then
			GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&TempFileName&".html")
		End If
	End Function
	
	'==============================
	'函 数 名：ReChangeField
	'作    用：多余标签清理
	'参    数：
	'TemplateCode  要处理的字符串
	'==============================
	Public Function ReChangeField(TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbGBookList.*?\$\}","",TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbField\_.*?\$\}","",TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbFieldList\_.*?\$\}","",TemplateCode)
		ReChangeField=TemplateCode
	End Function
	
	'==============================
	'函 数 名：RemoveHTML
	'作    用：过滤HTML
	'参    数：
	'strHTML   要处理的字符串
	'==============================
	Private Function RemoveHTML(strHTML)
		Dim objRegExp, Match, Matches 
		Set objRegExp = New Regexp 
		objRegExp.IgnoreCase = True 
		objRegExp.Global = True 
		objRegExp.Pattern = "<.+?>" 
		Set Matches = objRegExp.Execute(strHTML) 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		objRegExp.Pattern = "\&.+?;" 
		Set Matches = objRegExp.Execute(strHTML) 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		RemoveHTML=strHTML 
		Set objRegExp = Nothing 
	End Function

'========================模板引擎===========================
	'==============================
	'函 数 名：TemplateDo
	'作    用：获取优先处理函数
	'参    数：
	'==============================
	Public Function TemplateDo(TemplateCode)
		Dim ForI,IfI
		ForI=Instr(TemplateCode,"{$For")
		IfI=Instr(TemplateCode,"{$If")
		If ForI=0 And IfI=0 Then
			TemplateDo=TemplateCode
			Exit Function
		End If
		If ForI>0 And IfI>0 Then
			If ForI<IfI Then
				TemplateCode=TemplateFor(TemplateCode)
			Else
				TemplateCode=TemplateIf(TemplateCode)
			End If
		ElseIf ForI>0 Then
			TemplateCode=TemplateFor(TemplateCode)
		ElseIf IfI>0 Then
			TemplateCode=TemplateIf(TemplateCode)
		ELse
			TemplateDo=TemplateCode
			Exit Function
		End If
		Call TemplateDo(TemplateCode)
		TemplateDo=TemplateCode
	End Function

	'==============================
	'函 数 名：TemplateFor
	'作    用：处理For
	'参    数：
	'==============================	
	Private Function TemplateFor(TemplateCode)
		Temp=GetFor(TemplateCode)
		TemplateTag=Split(Split(Temp,"{$For(")(1),",")(0)
		TemplatePar=Split(Split(Temp,",")(1),")")(0)
		TemplateBCode=Right(Temp,Len(Temp)-Len("{$For("&TemplateTag&","&TemplatePar&")$}"))
		TemplateBCode=Left(TemplateBCode,Len(TemplateBCode)-8)
		Select Case TemplateTag
			Case "Nav"
				TemplateFor=Replace(TemplateCode,Temp,FkNav(TemplateBCode,TemplatePar))
			Case "ArticleList"
				TemplateFor=Replace(TemplateCode,Temp,FkArticleList(TemplateBCode,TemplatePar))
			Case "ProductList"
				TemplateFor=Replace(TemplateCode,Temp,FkProductList(TemplateBCode,TemplatePar))
			Case "DownList"
				TemplateFor=Replace(TemplateCode,Temp,FkDownList(TemplateBCode,TemplatePar))
			Case "FriendsList"
				TemplateFor=Replace(TemplateCode,Temp,FkFriendsList(TemplateBCode,TemplatePar))
			Case "JobList"
				TemplateFor=Replace(TemplateCode,Temp,FkJobList(TemplateBCode,TemplatePar))
			Case "SubjectList"
				TemplateFor=Replace(TemplateCode,Temp,FkSubjectList(TemplateBCode,TemplatePar))
			Case "GBookList"
				TemplateFor=Replace(TemplateCode,Temp,FkGBookList(TemplateBCode,TemplatePar))
			Case "PicList"
				TemplateFor=Replace(TemplateCode,Temp,FkPicList(TemplateBCode,TemplatePar))
			Case "NumList"
				TemplateFor=Replace(TemplateCode,Temp,FkNumList(TemplateBCode,TemplatePar))
			Case Else
				TemplateFor=Replace(TemplateCode,Temp,"")
		End Select
	End Function

	'==============================
	'函 数 名：TemplateIf
	'作    用：处理If
	'参    数：
	'==============================	
	Private Function TemplateIf(TemplateCode)
		Dim Check1,Check2
		Dim t_TempArr,t_Temp,tempIf,myChange,myTemp
		Temp=GetIf(TemplateCode)
		myTemp=Temp
		myChange=0
		'截取IF嵌套
		If GetCount(Temp,"{$If")>1 Then
			myChange=1
			t_Temp=Right(Temp,Len(Temp)-5)
			t_Temp=Left(t_Temp,Len(t_Temp)-10)
			t_TempArr=Split(FKFun.RegExpTest("\{\$If\((.|\n)*?\{\$End If\$\}",t_Temp),"|-_-|")
			For Each t_Temp In t_TempArr
				Temp=Replace(Temp,t_Temp,"{FangkaIF}")
			Next
		End If
		TemplatePar=Split(Split(Temp,"{$If(")(1),")")(0)
		If1=GetIfOne(Temp,"{$If("&TemplatePar&")$}")
		If2=Replace(Temp,If1,"")
		If2=Replace(If2,"{$If("&TemplatePar&")$}","")
		If2=Left(If2,Len(If2)-10)
		If2=Right(If2,Len(If2)-8)
		If If2="{$Null$}" Then
			If2=""
		End If
		If myChange=1 Then
			Temp=MyTemp
			MyTemp=If1&"|-_无聊的间隔_-|"&If2
			For Each t_Temp In t_TempArr
				MyTemp=Replace(MyTemp,"{FangkaIF}",t_Temp,1,1)
			Next
			If1=Split(MyTemp,"|-_无聊的间隔_-|")(0)
			If2=Split(MyTemp,"|-_无聊的间隔_-|")(1)
		End If
		TempArr=Split(TemplatePar,",")
		If IsNumeric(TempArr(0)) And IsNumeric(TempArr(1)) Then
			Check1=CDBl(TempArr(0))
			Check2=CDBl(TempArr(1))
		Else
			Check1=TempArr(0)
			Check2=TempArr(1)
		End If
		Select Case TempArr(2)
			Case ">"
				If Check1>Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<"
				If Check1<Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "="
				If Check1=Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<>"
				If Check1<>Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<="
				If Check1<=Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case ">="
				If Check1>=Check2 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "Mod"
				If (Check1 Mod Check2)=0 Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case Else
				TemplateIf=Replace(TemplateCode,Temp,"")
		End Select
	End Function
	
	'==============================
	'函 数 名：GetFor
	'作    用：获取For字符串
	'参    数：
	'==============================	
	Private Function GetFor(TemplateCode)
		Temp=Split(TemplateCode,"{$For")(0)
		Temp=Replace(TemplateCode,Temp,"")
		TempArr=Split(Temp,"{$Next$}")
		GetFor=TempArr(0)&"{$Next$}"
		i=1
		While GetCount(GetFor,"{$For")<>GetCount(GetFor,"{$Next$}")
			GetFor=GetFor&TempArr(i)&"{$Next$}"
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetIf
	'作    用：获取If字符串
	'参    数：
	'==============================	
	Private Function GetIf(TemplateCode)
		Temp=Split(TemplateCode,"{$If")(0)
		Temp=Replace(TemplateCode,Temp,"")
		TempArr=Split(Temp,"{$End If$}")
		GetIf=TempArr(0)&"{$End If$}"
		i=1
		While GetCount(GetIf,"{$If")<>GetCount(GetIf,"{$End If$}")
			GetIf=GetIf&TempArr(i)&"{$End If$}"
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetIfOne
	'作    用：获取If字符串Else前
	'参    数：
	'==============================	
	Private Function GetIfOne(TemplateCode,IfCode)
		TempArr=Split(TemplateCode,"{$Else$}")
		GetIfOne=Replace(TempArr(0),IfCode,"")
		i=1
		While GetCount(GetIfOne,"{$If")<>GetCount(GetIfOne,"{$End If$}")
			GetIfOne=GetIfOne&TempArr(i)
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetCount
	'作    用：判断字符串中相同字符的个数
	'参    数：
	'==============================	
	Private Function GetCount(Strs,Word)
		Dim N1,N2,N3
		N1=Len(Strs)
		N2=Len(Replace(Strs,Word,""))
		N3=Len(Word)
		GetCount=Clng(((N1-N2)/N3))
	End Function 
End Class
%>
