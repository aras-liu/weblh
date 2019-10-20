<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Admin/Get.asp
'文件用途：页面信息拉取页面
'版权所有：方卡在线
'==========================================

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GetTopNav() '读取顶部菜单
	Case 2
		Call GetLeftNav() '读取左侧菜单
	Case 3
		Call GetMain() '读取管理首页信息
	Case 4
		Call GetUserInfo() '读取管理用户信息
	Case 5
		Call GetAbout() '读取系统版权
	Case 7
		Call GetMenuNav() '读取分菜单模块
	Case 8
		Call GetLeftNav2() '读取内容设置菜单
End Select

'==========================================
'函 数 名：GetTopNav()
'作    用：读取顶部菜单
'参    数：
'==========================================
Sub GetTopNav()
%>
    	<ul id="TopNav">
        	<li onclick="ClickNav('Get.asp?Type=2','Get.asp?Type=3','#Nav_Main');" id="Nav_Main" class="NavNow">管理首页</li>
<%
	If Trim(Replace(Request.Cookies("FkAdminLimit2"),",",""))<>"" Or Request.Cookies("FkAdminLimitId")=0 Then
%>
        	<li onclick="ClickNav('Get.asp?Type=8','','#Nav_NS');" id="Nav_NS" class="NavOther">内容设置</li>
<%
	End If
	Sqlstr="Select Fk_Menu_Id,Fk_Menu_Name From [Fk_Menu] Order By Fk_Menu_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Instr(Request.Cookies("FkAdminLimit3"),",Menu"&Rs("Fk_Menu_Id")&",")>0 Or Request.Cookies("FkAdminLimitId")=0 Then
%>
        	<li onclick="ClickNav('Get.asp?Type=7&MenuId=<%=Rs("Fk_Menu_Id")%>','','#Nav_S<%=Rs("Fk_Menu_Id")%>');" id="Nav_S<%=Rs("Fk_Menu_Id")%>" class="NavOther"><%=Rs("Fk_Menu_Name")%></li>
<%
		End If
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </ul>
        <div class="Cal"></div>
<%
End Sub

'==========================================
'函 数 名：GetLeftNav()
'作    用：读取左侧菜单
'参    数：
'==========================================
Sub GetLeftNav()
%>
    	<div id="QuickNav"><a href="javascript:void(0);" onclick="ShowBox('Get.asp?Type=5');" id="QuickNav1"></a><a href="javascript:void(0);" onclick="alert('欢迎通过QQ或者MAIL跟我们联系！');" id="QuickNav2"></a><div class="Cal"></div></div>
        <div onclick="OpenMenu('M1');" class="LeftMenuTop">常规管理</div>
        <ul id="M1">
<%If FkAdmin.AdminCheck(4,"System1",Request.Cookies("FkAdminLimit1")&"|sethidden|6|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="ShowBox('SiteSet.asp?Type=1');"><span>&nbsp;&nbsp;&nbsp;</span>站点设置</a></li><%End If%>
<%If FkAdmin.AdminCheck(4,"Admin",Request.Cookies("FkAdminLimit1")) Then%><%If FkAdmin.AdminCheck(5,"7",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Admin.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>管理员管理</a></li><%End If%>
<%If FkAdmin.AdminCheck(5,"8",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Limit.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>权限管理</a></li><%End If%><%End If%>
<%If FkAdmin.AdminCheck(4,"System2",Request.Cookies("FkAdminLimit1")) Then%><%If FkAdmin.AdminCheck(5,"9",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>模板管理</a></li><%End If%>
<%If FkAdmin.AdminCheck(5,"10",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="ShowBox('TemplateHelp.asp?Type=1');"><span>&nbsp;&nbsp;&nbsp;</span>模板标签生成器</a></li><%End If%><%End If%>
<%If FkAdmin.AdminCheck(4,"System3",Request.Cookies("FkAdminLimit1")&"|sethidden|11|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="ShowBox('Html.asp?Type=1');"><span>&nbsp;&nbsp;&nbsp;</span>静态生成</a></li><%End If%>
<%If FkAdmin.AdminCheck(4,"Admin",Request.Cookies("FkAdminLimit1")) Then%><%If FkAdmin.AdminCheck(5,"12",Fk_Site_SysHidden) Then%>        	<%End If%>
<%If FkAdmin.AdminCheck(5,"13",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','File.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>上传文件管理</a></li><li><a href="javascript:void(0);" onclick="SetRContent('MainRight','UselessFile.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>冗余文件清理</a></li><%End If%><%End If%>
<%If FkAdmin.AdminCheck(4,"System4",Request.Cookies("FkAdminLimit1")&"|sethidden|14|sethidden|"&Fk_Site_SysHidden) Then%>        	<%End If%>

<%If FkAdmin.AdminCheck(4,"System6",Request.Cookies("FkAdminLimit1")&"|sethidden|16|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="ShowBox('Map.asp?Type=1');"><span>&nbsp;&nbsp;&nbsp;</span>搜索引擎地图生成</a></li><%End If%>
        </ul>
<%
End Sub

'==========================================
'函 数 名：GetMain()
'作    用：读取管理首页信息
'参    数：
'==========================================
Sub GetMain()
%>
    	<div id="MainRrightTop">欢迎进入管理后台！</div>
        <div id="NewBox">
        	<div id="NewBoxTop">
            	<ul>
                	<li>官方公告</li>
                </ul>
            </div>
            <div id="NewContent">
            	
            </div>
        </div>
        <div class="Cal"></div>
        <div id="AboutSystem">
        	<p>系统名称：企业网站管理系统&nbsp;&nbsp;&nbsp;&nbsp;系统版本：<%=FkSystemVersion%></p>
        	<p>版权所有：企业网站管理系统</p>
</div>
        <div class="Cal"></div>
<%
End Sub

'==========================================
'函 数 名：GetUserInfo()
'作    用：读取管理用户信息
'参    数：
'==========================================
Sub GetUserInfo()
%>
您的帐号是：&nbsp;&nbsp;<%=Request.Cookies("FkAdminName")%>&nbsp;&nbsp;[&nbsp;<a href="<%=SiteDir%>" target="_blank" title="前台首页">前台首页</a> <a href="javascript:void(0);" onclick="ShowBox('PassWord.asp?Type=1');" title="修改密码">修改密码</a> <a href="Logout.asp" title="退出登录">退出登录</a>&nbsp;]
<%
End Sub

'==========================================
'函 数 名：GetAbout()
'作    用：读取系统版权
'参    数：
'==========================================
Sub GetAbout()
%>
<div id="BoxTop" style="width:500px;">关于本系统[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="42%" height="25" align="right">系统名称：</td>
	        <td width="58%">&nbsp;企业网站管理系统</td>
      </tr>
	    <tr>
	        <td height="25" align="right">系统版本：</td>
	        <td>&nbsp;V1.2.3</td>
      </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
        <input type="button" onclick="CloseBox();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==========================================
'函 数 名：GetMenuNav()
'作    用：读取分菜单模块
'参    数：
'==========================================
Sub GetMenuNav()
	Dim Rs2,Rs3,ModuleSee
	Id=Clng(Request.QueryString("MenuId"))
%>
    	<div id="QuickNav"><a href="javascript:void(0);" onclick="ShowBox('Get.asp?Type=5');" id="QuickNav1"></a><a href="javascript:void(0);" onclick="alert('欢迎通过QQ或者MAIL跟我们联系！');" id="QuickNav2"></a><div class="Cal"></div></div>
        <div id="MenuAll"><select name="ssd" id="ssd" onChange="eval(this.options[this.selectedIndex].value);" class="Input" style="width:160px;">
      <option value="void(0);">快速通道</option>
<%
Call FKAdmin.GetModuleList(0,Id,0,0,"")
%>
</select></div>
<%
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Set Rs3=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level=0 Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		ModuleSee=0
		If FkAdmin.AdminCheck(4,"Module"&Rs("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
			ModuleSee=1
		ElseIf FkAdmin.AdminCheck(4,"See"&Rs("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
			ModuleSee=2
		End If
		If ModuleSee>0 Then
			Response.Write("<div class=""LeftMenuTop""")
			If ModuleSee=1 Then
				Response.Write("onclick="""&FKAdmin.GetNavGo(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))&"""")
			End If
			Response.Write(">"&Rs("Fk_Module_Name")&"</div>"&vbcrlf)
			Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level="&Rs("Fk_Module_Id")&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs2.Open Sqlstr,Conn,1,1
			If Not Rs2.Eof Then
				Response.Write("<ul>"&vbcrlf)
				While Not Rs2.Eof
					ModuleSee=0
					If FkAdmin.AdminCheck(4,"Module"&Rs2("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
						ModuleSee=1
					ElseIf FkAdmin.AdminCheck(4,"See"&Rs2("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
						ModuleSee=2
					End If
					If ModuleSee>0 Then
					Response.Write("<li><a href=""javascript:void(0);""")
					If ModuleSee=1 Then
						Response.Write("onclick="""&FKAdmin.GetNavGo(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"))&"""")
					End If
					Response.Write("><span>&nbsp;&nbsp;&nbsp;</span>"&Rs2("Fk_Module_Name")&"</a></li>"&vbcrlf)
					End If
					Sqlstr="Select Fk_Module_Id,Fk_Module_Name,Fk_Module_Type From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level="&Rs2("Fk_Module_Id")&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
					Rs3.Open Sqlstr,Conn,1,1
					If Not Rs3.Eof Then
						Response.Write("<ul>"&vbcrlf)
						While Not Rs3.Eof
							ModuleSee=0
							If FkAdmin.AdminCheck(4,"Module"&Rs3("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
								ModuleSee=1
							ElseIf FkAdmin.AdminCheck(4,"See"&Rs3("Fk_Module_Id"),Request.Cookies("FkAdminLimit3")) Then
								ModuleSee=2
							End If
							If ModuleSee>0 Then
							Response.Write("<li><a href=""javascript:void(0);""")
							If ModuleSee=1 Then
								Response.Write("onclick="""&FKAdmin.GetNavGo(Rs3("Fk_Module_Type"),Rs3("Fk_Module_Id"))&"""")
							End If
							Response.Write(">&nbsp;├&nbsp;"&Rs3("Fk_Module_Name")&"</a></li>"&vbcrlf)
							End If
							Rs3.MoveNext
						Wend
						Response.Write("</ul>"&vbcrlf)
					End If
					Rs3.Close
					Rs2.MoveNext
				Wend
				Response.Write("</ul>"&vbcrlf)
			End If
			Rs2.Close
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	Set Rs2=Nothing
End Sub

'==========================================
'函 数 名：GetLeftNav2()
'作    用：读取内容设置菜单
'参    数：
'==========================================
Sub GetLeftNav2()
%>
    	<div id="QuickNav"><a href="javascript:void(0);" onclick="ShowBox('Get.asp?Type=5');" id="QuickNav1"></a><a href="javascript:void(0);" onclick="alert('欢迎通过QQ或者MAIL跟我们联系！');" id="QuickNav2"></a><div class="Cal"></div></div>
<%If FkAdmin.AdminCheck(4,"System1",Request.Cookies("FkAdminLimit2")) Then%>        <div class="LeftMenuTop" onclick="OpenMenu('M2');">菜单管理</div>
        <ul id="M2">
<%If FkAdmin.AdminCheck(5,"17",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Menu.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>菜单管理</a></li><%End If%>
<%
	If FkAdmin.AdminCheck(5,"18",Fk_Site_SysHidden) Then
		Sqlstr="Select Fk_Menu_Id,Fk_Menu_Name From [Fk_Menu] Order By Fk_Menu_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		While Not Rs.Eof
%>
        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Module.asp?Type=1&MenuId=<%=Rs("Fk_Menu_Id")%>')"><span>&nbsp;&nbsp;&nbsp;</span><%=Rs("Fk_Menu_Name")%></a></li>
<%
			Rs.MoveNext
		Wend
		Rs.Close
	End If
%>
        </ul><%End If%>
        <div class="LeftMenuTop" onclick="OpenMenu('M3');">其他管理</div>
        <ul id="M3">
<%If FkAdmin.AdminCheck(4,"System2",Request.Cookies("FkAdminLimit2")) Then%><%If FkAdmin.AdminCheck(5,"19",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Friends.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>友情链接管理</a></li><%End If%>
<%If FkAdmin.AdminCheck(5,"20",Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','FriendsType.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>友情链接类型管理</a></li><%End If%><%End If%>

<%If FkAdmin.AdminCheck(4,"System4",Request.Cookies("FkAdminLimit2")&"|sethidden|22|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Word.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>站内关键字管理</a></li><%End If%>
<%If FkAdmin.AdminCheck(4,"System5",Request.Cookies("FkAdminLimit2")&"|sethidden|23|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Recommend.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>推荐类型管理</a></li><%End If%>


<%If FkAdmin.AdminCheck(4,"System8",Request.Cookies("FkAdminLimit2")&"|sethidden|26|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Infos.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>独立信息管理</a></li><%End If%>
<%If FkAdmin.AdminCheck(4,"System9",Request.Cookies("FkAdminLimit2")&"|sethidden|27|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Field.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>自定义字段管理</a></li><%End If%>

<%If FkAdmin.AdminCheck(4,"System11",Request.Cookies("FkAdminLimit2")&"|sethidden|29|sethidden|"&Fk_Site_SysHidden) Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','GModel.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;</span>留言模型管理</a></li><%End If%>
        </ul>
<%
End Sub
%>
<!--#Include File="../Code.asp"-->