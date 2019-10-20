<!--#Include File="AdminCheck.asp"-->

<%
'==========================================
'文 件 名：Admin/Fiel.asp
'文件用途：文件管理拉取页面
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"0",Request.Cookies("FkAdminLimit1"))

'定义页面变量
Dim NowFile,NowFloder,DirFloder,ObjFiles,ObjFile,ObjFloders,ObjFloder,path
Dim Fk_Template_Name,Fk_Template_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FileList() '上传文件列表
	Case 2
		Call EmptyDelDo() '删除空文件夹

	Case 3
		Call FileListDelDo() '批量删除上传文件执行
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FileList()
'作    用：上传文件列表
'参    数：
'==========================================
Sub FileList()
	Session("NowPage")=FkFun.GetNowUrl()
	On Error Resume Next
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');">刷新</a></li>
		 <li><a href="javascript:void(0);" onclick="DelIt('是否清理空文件夹？','UselessFile.asp?Type=2','MainRight','<%=Session("NowPage")%>');">清理空文件夹</a></li>
    </ul>
</div>
<div id="ListTop">
    冗余文件管理
</div>
<div id="ListContent">
<form action="UselessFile.asp?Type=3" method="post" name="DelList" id="DelList">
    <table width="98%" style="margin:8px auto;" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
	    <td align="center" class="ListTdTop" width=30>选择</td>
		  <td align="center" class="ListTdTop" width=50>序号</td>
            <td align="center" class="ListTdTop">文件/文件夹名</td>
            <td align="center" class="ListTdTop">类型</td>
            <td align="center" class="ListTdTop">标记</td>
        </tr>
<%
	path = FileDir&"Up"
	getAllFileList path,1

%>
 
<tr bgcolor="#FFFFFF" align="center" height="30">
<td colspan="5">
<input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)"> 全选
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('确定要删除选中的文件吗?')){Sends('DelList','UselessFile.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}">&nbsp;&nbsp;</td>
</tr>


        <tr>
            <td height="30" colspan="5">&nbsp;</td>
        </tr>
    </table>
</form>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FileDelDo()
'作    用：删除上传文件执行
'参    数：
'==========================================
Sub FileDelDo()
	Temp=Request.QueryString("File")
	Call FKFso.DelFile(FileDir&"Up/"&Temp)
	Response.Write("文件删除成功！")
End Sub

Sub EmptyDelDo()
    delAllEmpty "/Up"
	Response.Write("空文件夹清理成功！")
End Sub


'==========================================
'函 数 名：FileListDelDo()
'作    用：批量删除上传文件执行
'参    数：
'==========================================
Sub FileListDelDo()
dim ids,idsArray,arrayLen,i
ids=Replace(Trim(Request.Form("Fpath"))," ","")
	If ids="" Then
		Call FKFun.ShowErr("请选择要删除的内容！",2)
	End If
	
idsArray = split(ids,",") 
arrayLen=ubound(idsArray)
	for i=0 to arrayLen
'	Response.Write idsArray(i)
	Call FKFso.DelFile(idsArray(i))	
	next


	Response.Write("批量删除成功！")



End Sub
'==========================================
'函 数 名: getFolderList()
'作    用：获取子文件夹
'参    数：
'==========================================

Function getFolderList(Byval cDir)
	dim objFso,filePath,objFolder,objSubFolder,objSubFolders,i
	i=0
	redim  folderList(0)
	Set objFso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	filePath=server.mapPath(cDir)
	set objFolder=objFso.GetFolder(filePath)
	set objSubFolders=objFolder.Subfolders
	for each objSubFolder in objSubFolders
		ReDim Preserve folderList(i)
		With objSubFolder
			folderList(i)=.name&","&cDir&"/"&.name
		End With
		i=i + 1 
	next 
	set objFolder=nothing
	set objSubFolders=nothing
	Set objFso=nothing
	getFolderList=folderList
End Function


'==========================================
'函 数 名:getFileList()
'作    用：获取文件夹下的文件列表
'参    数：
'==========================================
Function getFileList(Byval cDir)
	dim objFso,filePath,objFolder,objFile,objFiles,i
	i=0
	redim  fileListok(0)
    Set objFso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	filePath=server.mapPath(cDir)
	set objFolder=objFso.GetFolder(filePath)
	set objFiles=objFolder.Files
	for each objFile in objFiles
		ReDim Preserve fileListok(i)
		With objFile
		  
			fileListok(i)=.name&","&cDir&"/"&.name
			
		End With
		
		i=i + 1 
	next 
	set objFolder=nothing
	set objFiles=nothing
	Set objFso=nothing
	getfileList=fileListok
End Function


'==========================================
'函 数 名:getAllFileList()
'作    用：输出所有冗余文件
'参    数：
'==========================================

Function getAllFileList(path,countnum)
	dim fileListArray,i,fileAttr,folderListArray,folderAttr,parentPath
	if not FKFso.IsFolder(path) then 
        FKFso.CreateFolder(path)
	end if	
	if not instr(path,"Up")>0 then	
		echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan=""5"" algin=""center"">你无限访问些目录,<a href=""javascript:history.go(-1)"" class=""txt_C1"">返回</a></td></tr>"
		response.end
	end if

	if not FKFso.IsFolder(path) then	
		echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan=""5"" algin=""center"">没有这个目录,<a href=""javascript:history.go(-1)"" class=""txt_C1"">返回</a></td></tr>"
		response.end
	end if
	folderListArray=getFolderList(path)
	
	if instr(folderListArray(0),",")>0 then
		for i=0 to ubound(folderListArray)
			folderAttr=split(folderListArray(i),",")
				
			if FKFso.IsFolder(folderAttr(1)) then
			
				getAllFileList folderAttr(1),countnum
			end if
		next
	end if	
	
	
	dim objNum0,objNum1,objNum2,objNum3,objNum4,objNum5,objNum6,objNum7,Rongyu
	fileListArray= getFileList(path)
	if instr(fileListArray(0),",")>0 then	
	for  i = 0 to ubound(fileListArray)
	fileAttr=split(fileListArray(i),",")
	Set objNum0=Conn.Execute("select count(*) from Fk_Article where Fk_Article_PicList like '%"&fileAttr(0)&"%' or InStr(1,LCase(Fk_Article_Content),LCase('"&fileAttr(0)&"'),0)<>0 or InStr(1,LCase(Fk_Article_Field),LCase('"&fileAttr(0)&"'),0)<>0")
	Set objNum1=Conn.Execute("select count(*) from Fk_Product where Fk_Product_PicList like '%"&fileAttr(0)&"%' or InStr(1,LCase(Fk_Product_Content),LCase('"&fileAttr(0)&"'),0)<>0 or InStr(1,LCase(Fk_Product_Field),LCase('"&fileAttr(0)&"'),0)<>0")	
	Set objNum2=Conn.Execute("select count(*) from Fk_Down where Fk_Down_PicList like '%"&fileAttr(0)&"%' or Fk_Down_File like '%"&fileAttr(0)&"%' or InStr(1,LCase(Fk_Down_Content),LCase('"&fileAttr(0)&"'),0)<>0 or InStr(1,LCase(Fk_Down_Field),LCase('"&fileAttr(0)&"'),0)<>0")
	Set objNum3=Conn.Execute("select count(*) from Fk_Friends where Fk_Friends_Logo like '%"&fileAttr(0)&"%'")
	Set objNum4=Conn.Execute("select count(*) from Fk_Module where InStr(1,LCase(Fk_Module_Content),LCase('"&fileAttr(0)&"'),0)<>0 or Fk_Module_Pic like '%"&fileAttr(0)&"%' or InStr(1,LCase(Fk_Module_Field),LCase('"&fileAttr(0)&"'),0)<>0")
	Set objNum5=Conn.Execute("select count(*) from Fk_Info where InStr(1,LCase(Fk_Info_Content),LCase('"&fileAttr(0)&"'),0)<>0")
	Set objNum6=Conn.Execute("select count(*) from Fk_Job where InStr(1,LCase(Fk_Job_About),LCase('"&fileAttr(0)&"'),0)<>0 or InStr(1,LCase(Fk_Job_Field),LCase('"&fileAttr(0)&"'),0)<>0")
	Set objNum7=Conn.Execute("select count(*) from Fk_Site where InStr(1,LCase(Fk_Site_Field),LCase('"&fileAttr(0)&"'),0)<>0")
	If   Err.Number <> 0   Then   
      Response.Write   Err.Source&": "&Err.Description 
	End   If
	
	if not objNum0(0)>0 and not objNum1(0)>0 and not objNum2(0)>0 and not objNum3(0)>0 and not objNum4(0)>0 and not objNum5(0)>0 and not objNum6(0)>0 and not objNum7(0)>0 then
	
				Rongyu="<tr>"&vbcrlf
				Rongyu=Rongyu&"<td height=""20"" align=""center"">&nbsp;&nbsp;<input type=""checkbox"" name=""Fpath"" value="""&fileAttr(1)&""" class=""Checks"" /></td></td>"&vbcrlf
				Rongyu=Rongyu&"<td>"&countnum&"</td>"&vbcrlf
				Rongyu=Rongyu&"<td><a target=""_blank"" href="""&fileAttr(1)&""">"&fileAttr(1)&"</a></td>"&vbcrlf
				Rongyu=Rongyu&"<td>"&UCase(Split(fileAttr(0),".")(UBound(Split(fileAttr(0),"."))))&"</td>"&vbcrlf
				Rongyu=Rongyu&"<td align=""center"">[ "&objNum0(0)&"—"&objNum1(0)&"—"&objNum2(0)&"—"&objNum3(0)&"—"&objNum4(0)&"—"&objNum5(0)&"—"&objNum6(0)&" ]</td>"&vbcrlf				
				Rongyu=Rongyu&"</tr>"&vbcrlf				
				Response.Write Rongyu
				countnum=countnum+1
				
	end if
	next		
	end if
End Function



'==========================================
'函 数 名:delAllEmpty()
'作    用：删除所有空文件夹
'参    数：
'==========================================

Function delAllEmpty(path)
	dim fileListArray,i,fileAttr,folderListArray,folderAttr,parentPath

	folderListArray=getFolderList(path)
	
	if instr(folderListArray(0),",")>0 then
		for i=0 to ubound(folderListArray)
			folderAttr=split(folderListArray(i),",")
				
			if FKFso.IsFolder(folderAttr(1)) then
				delAllEmpty folderAttr(1)
			end if
		next
	end if	
	
	

	fileListArray= getFileList(path)
	if not instr(fileListArray(0),",")> 0 and path <> "/Up" then
		FKFso.DelFolder(path)
	end if
End Function



%>
<!--#Include File="../Code.asp"-->