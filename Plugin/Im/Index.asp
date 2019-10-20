<!--#Include File="../../Include.asp"--><%
'==========================================
'文 件 名：Plugin/Im/Index.asp
'文件用途：客服代码
'版权所有：
'插件版本：V1.0.0
'==========================================
Dim Fk_QQ_Content
Dim TempDir
If Fk_Site_Dir<>"" Then
	TempDir=Fk_Site_Dir
Else
	TempDir=SiteDir
End If
Sqlstr="Select Fk_QQ_Content From [Fk_QQ]"
Rs.Open Sqlstr,Conn,1,1
If Not Rs.Eof Then
	Fk_QQ_Content=Rs("Fk_QQ_Content")
Else
	Rs.Close
	Response.End()
End If
Rs.Close
%>
document.writeln("<style type=\"text/css\">");
document.writeln("<!--");
document.writeln(".QQbox {z-index:99;right:0px;width:178px;position:absolute;top:80px}");
document.writeln(".QQbox .press {right:0px;width:33px;cursor:pointer;border-top-style:none;border-right-style:none;border-left-style:none;position:absolute;height:158px;border-bottom-style:none}");
document.writeln(".QQbox .Qlist {background:url(<%=TempDir%>Plugin/Im/QQBg.gif) repeat-y -155px 0px;left:0px;width:145px;position:absolute}");
document.writeln(".QQbox .Qlist .t {font-size:1px;float:right;width:145px;height:6px}");
document.writeln(".QQbox .Qlist .b {font-size:1px;float:ight;width:145px;height:6px}");
document.writeln(".QQbox .Qlist .t {background:url(<%=TempDir%>Plugin/Im/QQBg.gif) no-repeat left 50%}");
document.writeln(".QQbox .Qlist .b {background:url(<%=TempDir%>Plugin/Im/QQBg.gif) no-repeat right 50%}");
document.writeln(".QQbox .Qlist .con {background:#fff;margin:0px auto;width:90%}");
document.writeln(".QQbox .Qlist .con h2 {border-right:#3a708d 1px solid;border-top:#3a708d 1px solid;background:url(<%=TempDir%>Plugin/Im/QQBg.gif) repeat-y -163px 0px;font:bold 12px/22px \"宋体\";border-left:#3a708d 1px solid;color:#fff;border-bottom:#3a708d 1px solid;height:22px;text-align:center}");
document.writeln(".QQbox .Qlist .con ul {}");
document.writeln(".QQbox .Qlist .con ul li {padding-right:0px;padding-left:8px;background:#e8e8e8;padding-bottom:0px; padding-top:5px; height: 20px}");
document.writeln(".QQbox .Qlist .con ul li.odd {background:#fff}");
document.writeln("-->");
document.writeln("</style>");
document.write("<div class='QQbox' id='divQQbox' >");
document.write("	<div class='Qlist' id='divOnline' onmouseout='hideMsgBox(event);' style='display : none;'>");
document.write("		<div class='t'></div>");
document.write("		<div class='con'>");
<%=FKFun.HtmlToJs(FKFun.HTMLDncode(Fk_QQ_Content))%>
document.write("</div>");
document.write("<div class='b'></div>");
document.write("</div>");
document.write("<div id='divMenu' onmouseover='OnlineOver();'><img src='<%=TempDir%>Plugin/Im/QQ.png' class='press' alt='在线客服'></div>");
document.write("</div>");
//<![CDATA[
var tips;
var theTop = 80
/*这是默认高度,越大越往下*/
;
var old = theTop;
function initFloatTips() {
    tips = document.getElementById('divQQbox');
    moveTips();
};
function moveTips() {
    var tt = 50;
    if (window.innerHeight) {
        pos = window.pageYOffset
    } else if (document.documentElement && document.documentElement.scrollTop) {
        pos = document.documentElement.scrollTop
    } else if (document.body) {
        pos = document.body.scrollTop;
    }
    pos = pos - tips.offsetTop + theTop;
    pos = tips.offsetTop + pos / 10;
    if (pos < theTop) pos = theTop;
    if (pos != old) {
        tips.style.top = pos + "px";
        tt = 10;
        //alert(tips.style.top);
    }
    old = pos;
    setTimeout(moveTips, tt);
}
//!]]>
initFloatTips();

function OnlineOver() {
    document.getElementById("divMenu").style.display = "none";
    document.getElementById("divOnline").style.display = "block";
    document.getElementById("divQQbox").style.width = "145px";
}

function OnlineOut() {
    document.getElementById("divMenu").style.display = "block";
    document.getElementById("divOnline").style.display = "none";
}

function hideMsgBox(theEvent) { //theEvent用来传入事件，Firefox的方式
    if (theEvent) {　
        var browser = navigator.userAgent; //取得浏览器属性
        if (browser.indexOf("Firefox") > 0) { //如果是Firefox
            if (document.getElementById('divOnline').contains(theEvent.relatedTarget)) { //如果是子元素
                return; //结束函式
            }
        }
        if (browser.indexOf("MSIE") > 0) { //如果是IE
            if (document.getElementById('divOnline').contains(event.toElement)) { //如果是子元素
                return; //结束函式
            }
        }
    }
    /*要执行的操作*/
    document.getElementById("divMenu").style.display = "block";
    document.getElementById("divOnline").style.display = "none";
}
<!--#Include File="../../Code.asp"-->