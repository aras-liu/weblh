<%
'==========================================
'文 件 名：Class/Cls_Jpeg.asp
'文件用途：图片处理函数类
'版权所有：
'==========================================

Class Cls_Jpeg
	Private JpegObj,JpegObj2
	'==============================
	'函 数 名：CreateWaterText
	'作    用：加入水印
	'参    数：
	'==============================
	Public Function CreateWaterText(c_Pic,c_Text,c_Color,c_Font,c_FontSize,c_Weight,c_Position,c_x,c_y,wType,w_Pic,w_Transparence,w_BgColor)
		Dim t_w,t_h,p_w,p_h,cx_x,cx_y
		If Not FKFun.IsObjInstalled("Persits.Jpeg") Then
			Exit Function
		End If
		Set JpegObj=Server.CreateObject("Persits.Jpeg")
		If wType=1 Then
			p_w=(c_FontSize+1)*FKFun.StringLength(c_Text)/2
			p_h=c_FontSize+1
		Else
			If Not FKFso.IsFile(w_Pic) Then
				Exit Function
			End If
			Set JpegObj2=Server.CreateObject("Persits.Jpeg")
			JpegObj2.Open Server.MapPath(w_Pic)
			p_w=Clng(JpegObj2.OriginalWidth)
			p_h=Clng(JpegObj2.OriginalHeight)
		End If
		JpegObj.Open Server.MapPath(c_Pic)
		t_w=Clng(JpegObj.OriginalWidth)
		t_h=Clng(JpegObj.OriginalHeight)
		If t_w>=(p_w+c_x) And t_h>=(p_h+c_y) Then
			cx_x=GetPosition_X(t_w,p_w,c_Position,c_x)
			cx_y=GetPosition_Y(t_h,p_h,c_Position,c_y)
			If wType=1 Then
				JpegObj.Canvas.Font.COLOR="&H"&Replace(Right(c_Color,6),"#","")
				JpegObj.Canvas.Font.Family=c_Font
				JpegObj.Canvas.Font.size=c_FontSize
				JpegObj.Canvas.Font.Bold=c_Weight
				JpegObj.Canvas.Font.Quality=4
				JpegObj.Canvas.PrintText cx_x,cx_y,c_Text
				JpegObj.Quality=90
				JpegObj.Save Server.MapPath(c_Pic)
			Else
				If w_BgColor<>"" Then
					JpegObj.DrawImage cx_x,cx_y,JpegObj2,w_Transparence,"&H"&Replace(Right(w_BgColor,6),"#",""),90
				Else
					JpegObj.DrawImage cx_x,cx_y,JpegObj2,w_Transparence
				End If
				JpegObj.Quality=90
				JpegObj.save Server.MapPath(c_Pic)
				JpegObj2.Close
				Set JpegObj2=Nothing
			End If
		End If
		JpegObj.Close
		Set JpegObj=Nothing
	End Function
	
	'==============================
	'函 数 名：CreateSmall
	'作    用：图片缩略
	'参    数：缩略的图片c_Pic，缩略为图片c_ToPic，限制宽度c_w，限制高度c_h
	'==============================
	Public Function CreateSmall(c_Pic,c_ToPic,c_w,c_h)
		If Not FKFun.IsObjInstalled("Persits.Jpeg") Then
			CreateSmall=c_Pic
			Exit Function
		End If
		Dim t_w,t_h
		c_w=Clng(c_w)
		c_h=Clng(c_h)
		Set JpegObj=Server.CreateObject("Persits.Jpeg")
		JpegObj.Open Server.MapPath(c_Pic)
		t_w=Clng(JpegObj.OriginalWidth)
		t_h=Clng(JpegObj.OriginalHeight)
		If t_w<=c_w And t_h<=c_h Then
			CreateSmall=c_Pic
		Else
			If t_w>c_w Then
				t_h=c_w/t_w*t_h
				t_w=c_w
			End If
			If t_h>c_h Then
				t_w=c_h/t_h*t_w
				t_h=c_h
			End If
			JpegObj.Width=t_w
			JpegObj.Height=t_h
			JpegObj.Quality=95
			JpegObj.Save Server.MapPath(c_ToPic)
			CreateSmall=c_ToPic
		End If
		JpegObj.Close
		Set JpegObj=Nothing
	End Function
	
	Private Function GetPosition_X(Pic_w,Water_w,c_Position,c_x)
		Select Case c_Position
			Case 0 '左上
				GetPosition_X=c_x
			Case 1 '左下
				GetPosition_X=c_x
			Case 2 '居中
				GetPosition_X=(Pic_w-Water_w)/2
			Case 3 '右上
				GetPosition_X=Pic_w-Water_w-c_x
			Case 4 '右下
				GetPosition_X=Pic_w-Water_w-c_x
			Case Else
				GetPosition_X=0
		End Select
	End Function
	
	Private Function GetPosition_Y(Pic_h,Water_h,c_Position,c_y)
		Select Case c_Position
			Case 0 '左上
				GetPosition_Y=c_y
			Case 1 '左下
				GetPosition_Y=Pic_h-Water_h-c_y
			Case 2 '居中
				GetPosition_Y=(Pic_h-Water_h)/2
			Case 3 '右上
				GetPosition_Y=c_y
			Case 4 '右下
				GetPosition_Y=Pic_h-Water_h-c_y
			Case Else
				GetPosition_Y=0
		End Select
	End Function
End Class
%>
