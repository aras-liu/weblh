<%
'==========================================
'文 件 名：Class/Cls_Fun.asp
'文件用途：常规函数类
'版权所有：
'==========================================

Class Cls_Fun
	Private x,y,ii
	'==============================
	'函 数 名：AlertInfo
	'作    用：错误显示函数
	'参    数：错误提示内容InfoStr，转向页面GoUrl
	'==============================
	Public Function AlertInfo(InfoStr,GoUrl)
		If GoUrl="1" Then
			Response.Write "<Script>alert('"& InfoStr &"');location.href='javascript:history.go(-1)';</Script>"
		Else
			Response.Write "<Script>alert('"& InfoStr &"');location.href='"& GoUrl &"';</Script>"
		End If
		Session.CodePage=936
		Response.End()
	End Function
	
	'==============================
	'函 数 名：HTMLEncode
	'作    用：字符转换函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString=replace(fString, ">", "&gt;")
			fString=replace(fString, "<", "&lt;")
			fString=Replace(fString, CHR(32), " ")		
			fString=Replace(fString, CHR(34), "&quot;")
			fString=Replace(fString, CHR(39), "&#39;")
			fString=Replace(fString, CHR(9), "&nbsp;")
			fString=Replace(fString, CHR(13), "")
			fString=Replace(fString, CHR(10) & CHR(10), "<p></p>")
			fString=Replace(fString, CHR(10), "<br />")
			HTMLEncode=fString
		End If
	End Function
	
	'==============================
	'函 数 名：HTMLDncode
	'作    用：字符转回函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLDncode(fString)
		If Not IsNull(fString) Then
			fString=Replace(fString, "&gt;",">" )
			fString=Replace(fString, "&lt;", "<")
			fString=Replace(fString, " ", CHR(32))
			fString=Replace(fString, "&nbsp;", CHR(9))
			fString=Replace(fString, "&quot;", CHR(34))
			fString=Replace(fString, "&#39;", CHR(39))
			fString=Replace(fString, "", CHR(13))
			fString=Replace(fString, "<p></p>",CHR(10) & CHR(10) )
			fString=Replace(fString, "<br />",CHR(10) )
			HTMLDncode=fString
		End If
	End Function
	
	'==============================
	'函 数 名：AlertNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function AlertNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Call AlertInfo(ErrStr,"1")
		End If
	End Function

	'==============================
	'函 数 名：AlertString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function AlertString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Call AlertInfo(ShortErr,"1")
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Call AlertInfo(LongErr,"1")
		End If
	End Function
	
	'==============================
	'函 数 名：ShowNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function ShowNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Response.Write(ErrStr)
			Call FKDB.DB_Close()
			Session.CodePage=936
			Response.End()
		End If
	End Function

	'==============================
	'函 数 名：ShowString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function ShowString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Response.Write(ShortErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Response.Write(LongErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
	End Function
	
	'==============================
	'函 数 名：StringLength
	'作    用：判断字符串长度
	'参    数：需进行判断的文本Txt
	'==============================
	Public Function StringLength(Txt)
		Txt=Trim(Txt)
		x=Len(Txt)
		y=0
		For ii=1 To x
			If Asc(Mid(Txt,ii,1))<=2 or Asc(Mid(Txt,ii,1))>255 Then
				y=y + 2
			Else
				y=y + 1
			End If
		Next
		StringLength=y
	End Function
	
	'==============================
	'函 数 名：BeSelect
	'作    用：判断select选项选中
	'参    数：Select1,Select2
	'==============================
	Public Function BeSelect(Select1,Select2)
		If Select1=Select2 Then
			BeSelect=" selected='selected'"
		End If
	End Function
	
	'==============================
	'函 数 名：BeCheck
	'作    用：判断Check选项选中
	'参    数：Check1,Check2
	'==============================
	Public Function BeCheck(Check1,Check2)
		If Check1=Check2 Then
			BeCheck=" checked='checked'"
		End If
	End Function
	
	'==============================
	'函 数 名：CheckModule
	'作    用：判断模块类型，输出名称
	'参    数：要判断的类型ModuleId
	'==============================
	Public Function CheckModule(ModuleId)
		For i=0 To UBound(FKModuleId)
			If ModuleId=Clng(FKModuleId(i)) Then
				CheckModule=FKModuleName(i)
				Exit Function
			End If
		Next
	End Function	

	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCode(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&"1');"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows-1)&"');"">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows+1)&"');"">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&AllPage&"');"">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""SetRContent('MainRight','"&PageUrl&"'+this.options[this.selectedIndex].value);"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function

	'==============================
	'函 数 名：GetNowUrl
	'作    用：返回当前网址
	'参    数：
	'==============================
	Public Function GetNowUrl()
		GetNowUrl=Request.ServerVariables("Script_Name")&"?"&Request.ServerVariables("QUERY_STRING")
	End Function

	'==============================
	'函 数 名：ReplaceTest
	'作    用：正则表达式，替换字符串
	'参    数：规则patrn，要替换的字符串Str，替换为字符串replStr
	'==============================
	Public Function ReplaceTest(patrn,replStr,Str)
		Dim regEx
		Set regEx=New RegExp
		regEx.Pattern=patrn
		regEx.IgnoreCase=True
		regEx.Global=True 
		ReplaceTest=regEx.Replace(Str,replStr)
	End Function 

	'==============================
	'函 数 名：RegExpTest
	'作    用：正则表达式，获取字符串
	'参    数：源字符串patrn，规则strng
	'==============================
	Public Function RegExpTest(patrn, strng)
		Dim regEx, Matchs, Matches, RetStr
		Set regEx=New RegExp 
		regEx.Pattern=patrn 
		regEx.IgnoreCase=True
		regEx.Global=True
		Set Matches=regEx.Execute(strng) 
		For Each Matchs in Matches
			RetStr=RetStr & Matchs.Value & "|-_-|"
		Next 
		RegExpTest=RetStr 
	End Function 
	
	'==============================
	'函 数 名：NoTrash
	'作    用：垃圾信息强力判断
	'参    数：
	'TryStr      要判断的信息
	'CheckType   判断类型
	'ErrStr      垃圾信息提示
	'==============================
	Public Function NoTrash(TryStr,CheckType,ErrStr)
		Dim HttpCount,TryStr2,TryStr3
		TryStr2=LCase(TryStr)
		TryStr3=Replace(TryStr2,"|-^-|fangka|-^-|","")
		TryStr3=Replace(TryStr3,"|-_-|fangka|-_-|","")
		HttpCount=Clng(((Len(TryStr2)-Len(Replace(TryStr2,"http://","")))/7))
		If HttpCount>3 Then
			Call AlertInfo(ErrStr,"1")
		End If
		If CheckType=1 Then
			If StringLength(TryStr3)<=(Len(TryStr3)*1.2) Then
				Call AlertInfo(ErrStr,"1")
			End If
		End If
	End Function

	'==============================
	'函 数 名：GetHttpPage
	'作    用：获取页面源代码函数
	'参    数：网址HttpUrl，编码Cset
	'==============================
	Public Function GetHttpPage(HttpUrl,Cset)
		If IsNull(HttpUrl)=True Or HttpUrl="" Then
			GetHttpPage2="ERR！"
			Exit Function
		End If
		On Error Resume Next
		Dim Http
		Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
		Http.open "GET",HttpUrl,False
		Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing
			GetHttpPage="ERR！"
			Exit function
		End if
		GetHttpPage2=BytesToBSTR(Http.responseBody,Cset)
		Set Http=Nothing
		If Err.number<>0 then
			Err.Clear
			GetHttpPage2="ERR！"
			Exit function
		End If
	End Function

	'==============================
	'函 数 名：BytesToBstr
	'作    用：转换编码函数
	'参    数：字符串Body，编码Cset
	'==============================
	Private Function BytesToBstr(Body,Cset)
		Dim Objstream
		Set Objstream=Server.CreateObject("ado"&"d"&"b.st"&"re"&"am")
		Objstream.Type=1
		Objstream.Mode =3
		Objstream.Open
		Objstream.Write body
		Objstream.Position=0
		Objstream.Type=2
		Objstream.Charset=Cset
		BytesToBstr=Objstream.ReadText 
		Objstream.Close
		set Objstream=nothing
	End Function
	
	'==============================
	'函 数 名：HtmlToJs
	'作    用：HTML转JS
	'参    数：字符串CStrs
	'==============================
	Public Function HtmlToJs(CStrs)
		Dim ToJs
		CStrs=Replace(CStrs,Chr(10),"") 
		CStrs=Replace(CStrs,Chr(32)&Chr(32),"") 
		CStrs=Split(CStrs,Chr(13))
		ToJs=""
		For i=0 To UBound(CStrs) 
		If Trim(CStrs(i)) <> "" Then 
			CStrs(i)= Replace(CStrs(i),Chr(34),Chr(39)) 
			ToJs=ToJs&"document.write("&Chr(34)&CStrs(i)&Chr(34)&");"&Chr(10) 
		End If 
		Next
		HtmlToJs=ToJs
	End Function
		
	'==============================
	'函 数 名：UnEscape
	'作    用：js escape解码
	'参    数：字符串Str
	'==============================
	Public Function UnEscape(Str)
		dim r,s,c 
		s="" 
		For r=1 to Len(Str) 
			c=Mid(Str,r,1) 
			If Mid(Str,r,2)="%u" and r<=Len(Str)-5 Then 
				If IsNumeric("&H" & Mid(Str,r+2,4)) Then 
					s=s & CHRW(CInt("&H" & Mid(Str,r+2,4))) 
					r=r+5 
				Else 
					s=s & c 
				End If 
			ElseIf c="%" and r<=Len(Str)-2 Then 
				If IsNumeric("&H" & Mid(Str,r+1,2)) Then 
					s=s & CHRW(CInt("&H" & Mid(Str,r+1,2))) 
					r=r+2 
				Else 
					s=s & c 
				End If 
			Else 
				s=s & c 
			End If 
		Next 
		UnEscape=s 
	End Function 
	
	'==============================
	'函 数 名：RemoveHTML
	'作    用：过滤HTML
	'参    数：
	'==============================
	Public Function RemoveHTML(strHTML)
		Dim objRegExp, Match, Matches 
		Set objRegExp=New Regexp 
		objRegExp.IgnoreCase=True 
		objRegExp.Global=True 
		'取闭合的<> 
		objRegExp.Pattern="<.+?>" 
		'进行匹配 
		Set Matches=objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		'取特殊字符
		objRegExp.Pattern="\&.+?;" 
		'进行匹配 
		Set Matches=objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		RemoveHTML=strHTML 
		Set objRegExp=Nothing 
	End Function
	
	'==============================
	'函 数 名：IsObjInstalled
	'作    用：判断组件是否安装了
	'参    数：组件名strClassString
	'==============================
	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled=False 
		Err.Clear
		Dim xTestObj
		Set xTestObj=Server.CreateObject(strClassString)
		If 0=Err Then IsObjInstalled=True 
		Set xTestObj=Nothing
		Err.Clear
	End Function
	
	'==============================
	'函 数 名：ReplaceExt
	'作    用：过滤敏感后缀
	'参    数：要替换的字符串ReStr
	'==============================
	Public Function ReplaceExt(ReStr)
		ReStr=Replace(ReStr,"aspx","")
		ReStr=Replace(ReStr,"asp","")
		ReStr=Replace(ReStr,"jsp","")
		ReStr=Replace(ReStr,"cer","")
		ReStr=Replace(ReStr,"asa","")
		ReStr=Replace(ReStr,"cgi","")
		ReplaceExt=ReStr
	End Function
	
	'==============================
	'函 数 名：GetNumeric
	'作    用：获取处理数字参数
	'参    数：
	'GetTag      获取的值
	'GetDefault  默认值
	'==============================
	Public Function GetNumeric(GetTag,GetDefault)
		GetNumeric=Trim(Request.QueryString(GetTag))
		If GetNumeric<>"" Then
			GetNumeric=Clng(GetNumeric)
		Else
			GetNumeric=GetDefault
		End If
	End Function
	
	'==============================
	'函 数 名：ShowErr
	'作    用：错误提示函数
	'参    数：错误说明ErrStr
	'         错误类型ErrType(0前台，1后台)
	'==============================
	Public Function ShowErr(ErrStr,ErrType)
		If ErrType=0 Then
			Temp="<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
			Temp=Temp&"<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Temp=Temp&"<head>"
			Temp=Temp&"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
			Temp=Temp&"<title>系统提示</title>"
			Temp=Temp&"<style type=""text/css"">"
			Temp=Temp&"body,td,th {font-size: 12px;}"
			Temp=Temp&"body {margin-left:0px;margin-top:50px;margin-right: 0px;margin-bottom: 0px;}"
			Temp=Temp&"#errBox {width:500px;margin:0 auto;border:1px solid #CCC;}"
			Temp=Temp&"#errTitle {border-bottom:1px solid #CCC;line-height:25px;text-align:center;}"
			Temp=Temp&"#errStr {padding:10px;line-height:22px;text-indent:24px;}"
			Temp=Temp&"</style>"
			Temp=Temp&"</head>"
			Temp=Temp&"<body>"
			Temp=Temp&"<div id=""errBox"">"
			Temp=Temp&"	<div id=""errTitle"">系统提示</div>"
			Temp=Temp&"    <div id=""errStr"">"&ErrStr&"</div>"
			Temp=Temp&"</div>"
			Temp=Temp&"</body>"
			Temp=Temp&"</html>"
		End If
		If ErrType=1 Then
			Temp="<div id=""errBox"">"
			Temp=Temp&"	<div id=""errTitle"">系统提示</div>"
			Temp=Temp&"    <div id=""errStr"">"&ErrStr&"</div>"
			Temp=Temp&"</div>"
			Temp=Temp&"</body>"
		End If
		If ErrType=2 Then
			Temp=ErrStr
		End If
		Response.Write(Temp)
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End Function
	
	'==============================
	'函 数 名：Jmail
	'作    用：邮件发送
	'参    数：发送到mailTo，邮件主题mailTopic，邮件内容mailBody，编码mailCharset，正文格式mailContentType
	'==============================
	Public Function Jmail(mailTo,mailTopic,mailBody,mailCharset,mailContentType) 
		If IsObjInstalled("JMail.Message") Then
			Dim myJmail 
			Set myJmail=Server.CreateObject("JMail.Message") 
			myJmail.Charset=mailCharset
			myJmail.silent=true 
			myJmail.ContentType=mailContentType
			myJmail.MailServerUserName=Mail_Name 
			myJmail.MailServerPassWord=Mail_Pass 
			myJmail.AddRecipient mailTo,""
			myJmail.Subject=mailTopic 
			myJmail.Body=mailBody
			myJmail.FromName="" 
			myJmail.From=Mail_Address 
			myJmail.Priority=3
			myJmail.Send(Mail_Smtp) 
			myJmail.Close 
			Set myJmail=nothing 
			If Err Then 
				Jmail=Err.Description 
				Err.Clear 
			Else 
				Jmail="发送成功" 
			End If 
		Else
			Jmail="不支持JMAIL组件"
		End If
	End Function 
End Class
%>
