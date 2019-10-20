<!--#Include File="../Include.asp"--><%
'==========================================
'文 件 名：Search/Index.asp
'文件用途：搜索页
'版权所有：
'==========================================

'定义变量
Dim PageCode
Set FKTemplate=New Cls_Template
Set FKPageCode=New Cls_PageCode

'获取参数
SearchStr=FKFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
SearchType=Trim(Request.QueryString("SearchType"))
SearchTemplate=FKFun.HTMLEncode(Trim(Request.QueryString("SearchTemplate")))
SearchField=FKFun.HTMLEncode(Trim(Replace(Request.QueryString("SearchField")," ","")))
SearchFieldList=FKFun.HTMLEncode(Trim(Replace(Request.QueryString("SearchFieldList")," ","")))
If SearchType<>"" Then
	SearchType=Clng(SearchType)
Else
	SearchType=0
End If
Call FKFun.AlertString(SearchStr,1,50,0,"请输入搜索关键字！","搜索关键字不能大于50个字符！")
PageNow=Trim(Request.QueryString("Page"))
If PageNow="" Then
	PageNow=1
Else
	PageNow=Clng(PageNow)
End If
TempPageSize=Fk_Site_PageSize
PageCode=FKPageCode.cSearch(SearchTemplate)
Response.Write(PageCode)
%>
<!--#Include File="../Code.asp"-->
