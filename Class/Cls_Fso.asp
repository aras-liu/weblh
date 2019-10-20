<%
'==========================================
'文 件 名：Class/Cls_Fso.asp
'文件用途：常规函数类
'版权所有：
'==========================================

Class Cls_Fso
	'==============================
	'函 数 名：FsoLineWrite
	'作    用：按行写入文件
	'参    数：文件相对路径FilePath，写入行号LineNum，写入内容LineContent
	'==============================
	Public Function FsoLineWrite(FilePath,LineNum,LineContent)
		If LineNum<1 Then Exit Function
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If Not Fso.FileExists(Server.MapPath(FilePath)) Then Exit Function
		Temp=FsoFileRead(FilePath)
		TempArr=Split(Temp,Chr(13)&Chr(10))
		TempArr(LineNum-1)=LineContent
		Temp=Join(TempArr,Chr(13)&Chr(10))
		Call CreateFile(FilePath,Temp)
		Set Fso=Nothing
	End Function
	
	'==============================
	'函 数 名：FsoFileRead
	'作    用：读取文件
	'参    数：文件相对路径FilePath
	'==============================
	Public Function FsoFileRead(FilePath)
		On Error Resume Next
		Set objAdoStream = Server.CreateObject("A"&"dod"&"b.St"&"r"&"eam")
		objAdoStream.Type=2
		objAdoStream.mode=3  
		objAdoStream.charset="utf-8"
		objAdoStream.open 
		objAdoStream.LoadFromFile Server.MapPath(FilePath) 
		FsoFileRead=objAdoStream.ReadText 
		objAdoStream.Close
		Set objAdoStream=Nothing
		If Err Then
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
			Response.Write "<body style='font-size:12px'>"
			Response.Write "文件路径：" & Server.MapPath(FilePath) & "<br />"
			Response.Write "错 误 号：" & Err.Number & "<br />"
			Response.Write "错误描述：" & Err.Description & "<br />"
			Response.Write "错误来源：" & Err.Source & "<br />"
			Response.Write "</body>"
			Err.Clear()
			Response.End()
		End If
	End Function
	
	'==============================
	'函 数 名：CreateFolder
	'作    用：创建文件夹
	'参    数：文件夹相对路径FolderPath
	'==============================
	Public Function CreateFolder(FolderPath)
		If FolderPath<>"" Then
			If Instr(FolderPath,".")>0 Then
				Response.Write("无法生成带点号（.）的目录！")
				Response.End()
			End If
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Set F=Fso.CreateFolder(Server.MapPath(FolderPath))
			CreateFolder=F.Path
			Set F=Nothing
			Set Fso=Nothing
		End If
	End Function
	
	'==============================
	'函 数 名：CreateFile
	'作    用：创建文件
	'参    数：文件相对路径FilePath，文件内容FileContent
	'==============================
	Public Function CreateFile(FilePath,FileContent)
		Dim FsoTemp,FsoTempArr,Fsoi
		FsoTemp=""
		FsoTempArr=Split(FilePath,"/")
		For Fsoi=0 to UBound(FsoTempArr)-1
			If FsoTempArr(Fsoi)<>"" And FsoTempArr(Fsoi)<>".." Then
				FsoTemp=FsoTemp&FsoTempArr(Fsoi)&"/"
				If IsFolder(FsoTemp)=False Then
					Call CreateFolder(FsoTemp)
				End If
			Else
				FsoTemp=FsoTemp&FsoTempArr(Fsoi)&"/"
			End If
		Next
		Set objAdoStream = Server.CreateObject("A"&"dod"&"b.St"&"r"&"eam")
		objAdoStream.Type = 2
		objAdoStream.Charset = "utf-8" 
		objAdoStream.Open
		objAdoStream.WriteText = FileContent
		objAdoStream.SaveToFile Server.MapPath(FilePath),2
		objAdoStream.Close()
		Set objAdoStream = Nothing
	End Function
	
	'==============================
	'函 数 名：DelFolder
	'作    用：删除文件夹
	'参    数：文件夹相对路径FolderPath
	'==============================
	Public Function DelFolder(FolderPath)
		If IsFolder(FolderPath)=True Then
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Fso.DeleteFolder(Server.MapPath(FolderPath))
			Set Fso=Nothing
		End If 
	End Function 
	
	'==============================
	'函 数 名：DelFile
	'作    用：删除文件
	'参    数：文件相对路径FilePath
	'==============================
	Public Function DelFile(FilePath)
		If IsFile(FilePath)=True Then 
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Fso.DeleteFile(Server.MapPath(FilePath))
			Set Fso=Nothing
		End If
	End Function 
	 
	'==============================
	'函 数 名：IsFile
	'作    用：检测文件是否存在
	'参    数：文件相对路径FilePath
	'==============================
	Public Function IsFile(FilePath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If (Fso.FileExists(Server.MapPath(FilePath))) Then
			IsFile=True
		Else
			IsFile=False
		End If
		Set Fso=Nothing
	End Function
	
	'==============================
	'函 数 名：IsFolder
	'作    用：检测文件夹是否存在
	'参    数：文件相对路径FolderPath
	'==============================
	Public Function IsFolder(FolderPath)
		If FolderPath<>"" Then
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			If Fso.FolderExists(Server.MapPath(FolderPath)) Then  
				IsFolder=True
			Else
				IsFolder=False
			End If
			Set Fso=Nothing
		End If
	End Function
	
	'==============================
	'函 数 名：CopyFiles
	'作    用：复制文件
	'参    数：文件来源地址SourcePath，文件复制到地址CopyToPath
	'==============================
	Public Function CopyFiles(SourcePath,CopyToPath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		Fso.CopyFile Server.MapPath(SourcePath),Server.MapPath(CopyToPath)
		Set Fso=nothing
	End Function
	
	'==============================
	'函 数 名：CopyFolder
	'作    用：复制文件夹
	'参    数：源文件夹FolderName，复制到文件夹FolderPath
	'==============================
	Public Function CopyFolder(FolderName,FolderPath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If Fso.Folderexists(Server.MapPath(FolderName)) Then
			If Fso.FolderExists(Server.MapPath(FolderPath)) Then
				Fso.CopyFolder Server.MapPath(FolderName),Server.MapPath(FolderPath)
			Else
				Fso.CreateFolder(Server.MapPath(FolderPath))
				Fso.CopyFolder Server.MapPath(FolderName),Server.MapPath(FolderPath)
			End if 
		End If 
		Set Fso=nothing
	End Function 
	
	'==============================
	'函 数 名：ReNameFolder
	'作    用：重命名文件夹
	'参    数：原文件夹名FolderName，新文件夹名NewFolderName
	'==============================
	Public Function ReNameFolder(FolderName,NewFolderName)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		Fso.MoveFolder Server.MapPath(FolderName),Server.MapPath(NewFolderName)
		Set Fso=nothing
	End Function 

	'==========================================
	'函 数 名：GetSize()
	'作    用：获取空间使用
	'参    数：
	'==========================================
	Public Function GetSize(Path)
		On Error Resume Next
		Dim Size,ShowSize,Paths
		Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
		Paths=Path
		Paths=Server.Mappath(Paths) 		 		
		Set F=Fso.GetFolder(Paths) 		
		Size=F.Size
		ShowSize=Size&"&nbsp;Byte" 
		If Size>1024 Then
			Size=(Size/1024)
			ShowSize=Size&"&nbsp;KB"
		End If
		If Size>1024 Then
			Size=(Size/1024)
			ShowSize=FormatNumber(Size,2)&"&nbsp;MB"		
		End If
		If Size>1024 then
			Size=(Size/1024)
			ShowSize=FormatNumber(Size,2)&"&nbsp;GB"	   
		End If   
		Set Fso=nothing
		GetSize=ShowSize
	End Function

	'==========================================
	'函 数 名：CompactAccess()
	'作    用：压缩数据库
	'参    数：
	'==========================================
	Public Function CompactAccess(BDir,BData)
		On Error Resume Next
		Dim Engine
		Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
		Set Engine=CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(BData),"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(BDir)&"\Temp.mdb"
		If Err Then
			Err.Clear
			Response.Write("数据库被打开了，无法进行压缩！")
			Set Fso=nothing
			Set Engine=nothing
			Response.End()
		End If
		Fso.CopyFile Server.MapPath(BDir)&"\Temp.mdb",Server.MapPath(BData)
		Fso.DeleteFile(Server.MapPath(BDir)&"\Temp.mdb")
		Set Fso=nothing
		Set Engine=nothing
	End Function
End Class
%>
