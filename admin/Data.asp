<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Admin/Data.asp
'文件用途：数据库操作
'版权所有：方卡在线
'==========================================

Call FKAdmin.AdminCheck(3,"0",Request.Cookies("FkAdminLimit1"))

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call DataList() '操作主页面
	Case 2
		Call CompactDB() '压缩数据库
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：DataList()
'作    用：操作主页面
'参    数：
'==========================================
Sub DataList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Data.asp?Type=1')">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    数据库操作
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">数据库维护</td>
        </tr>
        <tr>
            <td height="30" align="center">数据库压缩：
                <input type="button" onclick="DelIt('是否要压缩数据库？','Data.asp?Type=2','MainRight','Data.asp?Type=1');" class="Button" name="button" id="button" value="压 缩" />&nbsp;&nbsp;<span class="qbox" title="<p>当添加删除过于频繁时，数据库会虚大，可用压缩数据库功能减小数据库硬盘占用。</p>"><img src="Images/help.jpg" /></span></td>
        </tr>
        <tr>
            <td align="center" class="ListTdTop">空间占用</td>
        </tr>
        <tr>
            <td height="30" align="center">占用空间：<%=FKFso.GetSize(SiteDir)%>；附件占用：<%=FKFso.GetSize(SiteDir&"Up")%></td>
        </tr>
        <tr>
            <td height="25" align="center">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：CompactDB()
'作    用：压缩数据库
'参    数：
'==========================================
Sub CompactDB()
	Call FKDB.DB_Close()
	Call FKFso.CompactAccess(SiteDir,SiteData)
	Call FKDB.DB_Open()
	Response.Write("数据库压缩成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->