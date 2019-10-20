<!--#Include File="../../Include.asp"--><%
'==========================================
'文 件 名：Plugin/VoteDo.asp
'文件用途：在线投票提交
'版权所有：
'插件版本：V1.0.0
'==========================================

'定义页面变量
Dim Fk_Vote_Ticket,TempArr2
Dim TempDir
If Fk_Site_Dir<>"" Then
	TempDir=Fk_Site_Dir
Else
	TempDir=SiteDir
End If
Fk_Vote_Ticket=Replace(FKFun.HTMLEncode(Trim(Request.Form("V")))," ","")
Id=Trim(Request.Form("Id"))
If Request.Cookies("V"&Id)="1" Then
	Call FKFun.AlertInfo("您已经投过票了！","1")
End If
Call FKFun.AlertString(Fk_Vote_Ticket,1,255,0,"请至少选择一个选项！","选项不能大于255个字符！")
Call FKFun.AlertNum(Id,"未得到投票编号！")
Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id="&Id&""
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	If IsNull(Rs("Fk_Vote_Ticket")) Or Rs("Fk_Vote_Ticket")="" Then
		TempArr=Split(Rs("Fk_Vote_Content"),"<br />")
		Temp=""
		For i=0 To UBound(TempArr)
			If Instr(Fk_Vote_Ticket,i)>0 Then
				If Temp="" Then
					Temp="1"
				Else
					Temp=Temp&"|1"
				End If
			Else
				If Temp="" Then
					Temp="0"
				Else
					Temp=Temp&"|0"
				End If
			End If
		Next
	Else
		TempArr=Split(Rs("Fk_Vote_Ticket"),"|")
		Temp=""
		For i=0 To UBound(TempArr)
			If Instr(Fk_Vote_Ticket,i)>0 Then
				If Temp="" Then
					Temp=Clng(TempArr(i))+1
				Else
					Temp=Temp&"|"&(Clng(TempArr(i))+1)
				End If
			Else
				If Temp="" Then
					Temp=TempArr(i)
				Else
					Temp=Temp&"|"&TempArr(i)
				End If
			End If
		Next
	End If
	Application.Lock()
	Rs("Fk_Vote_Ticket")=Temp
	Rs("Fk_Vote_Count")=Rs("Fk_Vote_Count")+1
	Rs.Update
	Application.UnLock()
	Response.Cookies("V"&Id)="1"
	Call FKFun.AlertInfo("投票成功！",TempDir)
Else
	Call FKFun.AlertInfo("未找到投票项目！",TempDir)
End If
Rs.Close
%>
<!--#Include File="../../Code.asp"-->
