<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/inc.asp"-->
<!--#include file="inc/indexnew.asp"-->
<%call index_ad()%>
<%
sql="select articlecount,softcount,coolsitescount,djcount,newscount,piccount,jscount from allcount" '读取首页信息统计数据
set rs=conn.execute(sql)
articlecount=rs("articlecount")
softcount=rs("softcount")
coolsitescount=rs("coolsitescount")
djcount=rs("djcount")
newscount=rs("newscount")
piccount=rs("piccount")
jscount=rs("jscount")
call head()
call menu()
call index_music()
Response.Write "<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center>" & vbCrLf
Response.Write "<TBODY><TR><TD class=l></TD>" & vbCrLf
response.write"<td width=180 align=center>今天是："&year(now)&"年"&month(now)&"月"&day(now)&"日 <font color='#0000FF'>"&weekdayname(weekday(now))&"</font></td>"& vbcrlf
Response.Write "<TD class=m style=""PADDING-RIGHT: 20px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px""><marquee class=gg behavior='loop' scrollDelay='100' scrollAmount='5'onmouseover='javascript:scrollAmount=1' onmouseout='javascript:scrollAmount=5' width=""100%"">"
call index_diary()
Response.Write "</MARQUEE></TD>" & vbCrLf
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR vAlign=top align=left><TD width=200>" & vbCrLf
Response.Write "<DIV class=member>" & vbCrLf
Response.Write "<DIV class=lframe><SCRIPT type=text/javascript>zonelovetabletop(""用 户 登 录"",""ml"");</SCRIPT>" & vbCrLf
Response.Write "<UL class=nl>" & vbCrLf
Response.Write "<table><tr><td>" & vbCrLf
Response.Write "<iframe width='180' height='160' scrolling='no' border='0' frameborder='0' src='bbs.asp' ></iframe></td></tr></table>" & vbCrLf
Response.Write "</UL>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "</TD><TD>" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""&nbsp;&nbsp;-= 学 校 快 讯 =-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-= 校 内 新 闻 =-"");</SCRIPT>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=ml></TD><TD class=mm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=""1"" bordercolor=""#f0f0f0"" style=""border-collapse: collapse""  height='165'><TBODY><TR>" & vbCrLf
Response.Write "<TD vAlign=top >"
Response.write"<iframe width='218' height='173' scrolling='no' border='0' frameborder='0' src='gundong.asp' ></iframe>"
Response.Write "</TD>" & vbCrLf
Response.Write "<TD vAlign=top width=* style=""padding-top:5px;"">" & vbCrLf
Response.Write "<UL class=nl>"
call index_news()
Response.Write "</UL></MARQUEE></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=mr></TD></TR></TBODY></TABLE></DIV>" & vbCrLf
Response.Write "</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD>" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""我 校 影 像"");</SCRIPT>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR>" & vbCrLf
Response.Write "<TD class=ml></TD><TD class=mm>"
response.write "<td width=760><marquee behavior=""alternate"" height=""150"" onmouseover='javascript:scrollAmount=0' onmouseout='javascript:scrollAmount=1' direction=""left"" scrollamount=""1"" scrolldelay=""30"" width=""100%"">" &vbcrlf
call index_pic()
Response.Write "</marquee></td></TD><TD class=mr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=bl></TD><TD class=bm>&nbsp;</TD><TD class=br></TD></TR></TBODY></TABLE></DIV>" & vbCrLf
Response.Write "</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR vAlign=top align=left><TD width=200>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""邮 箱 登 录"",""ml"");</SCRIPT>"
Response.write"<table>"
Response.Write "<tr><TD vAlign=top>"
Response.write"<iframe width='180' height='79' align='middle' scrolling='no' border='0' frameborder='0' src='email.asp' ></iframe>"
Response.Write "</TD></tr>"
Response.write"</table>"& vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""天 气 预 报"",""ml"");</SCRIPT>"
Response.write"<table>"
Response.Write "<tr><TD vAlign=top>"
Response.write"<iframe width='180' height='50' scrolling='no' border='0' frameborder='0' src='http://weather.265.com/weather.htm' ></iframe>"
Response.Write "</tr></TD>"
Response.write"<tr><TD align=middle><a href=""http://weather.265.com/"" target=""_blank"">欢迎定制各地城市 天气预报</a></td></tr>"
Response.write"</table>"& vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""百 度 搜 索"",""ml"");</SCRIPT>"
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=190 border=0><TR><TD align=middle ><iframe id=""baiduframe"" border=""0"" vspace=""0"" hspace=""0"" marginwidth=""0"" marginheight=""0""framespacing=""0"" frameborder=""0"" scrolling=""no"" width=""140"" height=""47""src=""http://unstat.baidu.com/bdun.bsc?tn=sitesowang&cv=1&cid=316&csid=105&rkcs=0&bgcr=FFFFFF&ftcr=000000&rk=0&bd=0&tbsz=&tbst=&bdas=0""></iframe></TD></TR></TABLE>"
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""网 站 调 查"",""ml"");</SCRIPT>"
call index_vote()
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""友 情 链 接"",""ml"");</SCRIPT>" & vbCrLf
Response.Write "<DIV id=oRollV style=""MARGIN: 0px; OVERFLOW: hidden; HEIGHT: 210px"">"
Response.Write "<DIV id=oRollV1>" & vbCrLf
call index_flink()
Response.Write "</DIV><DIV id=oRollV2></DIV></DIV>" & vbCrLf
Response.Write "<DIV align=center style=""PADDING-TOP: 5px;"">"
call index_flinkwz()
Response.Write "</div><div align=center style=""PADDING-TOP: 7px;""><a href='link.asp'>申请链接</a>&nbsp;<a href='link.asp'>更多链接...</a></div><SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "</TD><TD>" & vbCrLf
Response.Write "<TABLE id=dlAll style=""WIDTH: 100%; BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR height=""220"" valign=""top"">" & vbCrLf
Response.Write "<TD valign=top style=""WIDTH: 50%"">" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD class=tl></TD><TD class=tm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD>" & vbCrLf
Response.Write "<SPAN class=tt> 学 校 党 建 </SPAN></TD>" & vbCrLf
Response.Write "<TD align=right><A href=""art.asp?cat_id=12""><IMG alt=more src=""img/mory.gif"" border=0></A>&nbsp; </TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=ml></TD><TD class=mm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0><TBODY>"
call index_djgz() 
Response.Write "</TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=mr></TD></TR></TBODY></TABLE></DIV>" & vbCrLf
Response.Write "</TD><TD valign=top style=""WIDTH: 50%"">" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=tl></TD><TD class=tm><TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD><SPAN class=tt> 教 师 之 窗</SPAN> </TD>" & vbCrLf
Response.Write "<TD align=right><A href=""art.asp?cat_id=4""><IMG alt=more src=""img/mory.gif"" border=0></A>&nbsp;</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD class=ml></TD><TD class=mm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>"
call index_teacher()
Response.Write "</TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=mr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</DIV></TD>" & vbCrLf


response.write"<table id=footer border='0' cellspacing='0' cellpadding='0' align='center'>" & vbCrLf
Response.Write "<tr>"&vbcrlf
Response.Write "    <td class='zonelovebg' height='3' colspan='2'></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
response.write"<td width='560'><iframe width='560' height='115' scrolling='no' border='0' frameborder='0' src='aoyun.asp' ></iframe></td>" & vbCrLf
response.Write"</tr>"& vbcrlf
response.write"</table>"& vbcrlf
Response.Write "<TABLE id=dlAll style=""WIDTH: 100%; BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR height=""220"" valign=""top""><TD valign=top style=""WIDTH: 50%"">" & vbCrLf
Response.Write "<DIV class=mframe><TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD class=tl></TD><TD class=tm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY><TR><TD><SPAN class=tt>学 生 园 地</SPAN> </TD>" & vbCrLf
Response.Write "<TD align=right><A href=""art.asp?cat_id=5""><IMG alt=more src=""img/mory.gif"" border=0></A>&nbsp; </TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=ml></TD><TD class=mm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>"
call index_student()
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%""><TBODY><TR><TD class=tl></TD><TD class=tm><TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TR><TD><SPAN class=tt>父 母 学 堂</SPAN></TD><TD align=right><A href=""art.asp?cat_id=13""><IMG alt=more src=""img/mory.gif"" border=0>" & vbCrLf
Response.Write "</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>"
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY><TR><TD class=ml></TD><TD class=mm>" & vbCrLf
call index_fmxt()
Response.Write "<TBODY>"
Response.Write "</TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=mr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</DIV>" & vbCrLf
Response.Write "</TD><TD valign=top style=""WIDTH: 50%"">" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD class=tl></TD><TD class=tm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD><SPAN class=tt>轻 松 一 刻</SPAN></TD>" & vbCrLf
Response.Write "<TD align=right><A href=""dj.asp""><IMG alt=more src=""img/mory.gif"" border=0></A>&nbsp; </TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY><TR><TD class=ml></TD><TD class=mm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>"
call index_dj()
Response.Write "<DIV class=mframe>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD class=tl></TD><TD class=tm>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR><TD><SPAN class=tt> 内 务 工 作</SPAN></td><TD align=right><A href=""art.asp?cat_id=3""><IMG alt=more src=""img/mory.gif"" border=0></A>&nbsp;</TD>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=tr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"">" & vbCrLf
Response.Write "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>" & vbCrLf
Response.Write "<TBODY>"
call index_work()
Response.Write "</TBODY></TABLE>" & vbCrLf
Response.Write "</TD><TD class=mr></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</DIV>" & vbCrLf
Response.Write "</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "</DIV>" & vbCrLf
Response.Write "</TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<SCRIPT language=javascript type=text/javascript>StartRollV();StartRollH();</SCRIPT>"
call index_skin()
call footer()
%>


</body>
</html>
