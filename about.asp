<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/inc.asp"-->
<%
start="��վ��Ϣ"
call head()
call menu()
sql="select gy,lx,gg,dt from admin"
set rs=conn.execute(sql)
Response.Write "<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center><TBODY><TR><TD class=l></TD><TD height='23' align='right'><a href='about.asp'>��ѧУ�ſ�</a>&nbsp;<a href='?ly=lx'>����ϵ����</a>&nbsp;<a href='?ly=gg'>���쵼�´�</a>&nbsp;<a href='?ly=dt'>����֯����</a>&nbsp;&nbsp;</TD><TD class=r></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"">" & vbCrLf
Response.Write "<TBODY><TR vAlign=top bgcolor=""#FCFCFC""><TD>" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
ly=request("ly")
if ly="" then
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""ѧУ�ſ�"");</SCRIPT><table width=""100%"" align=""center"" cellspacing=""0"" cellpadding=""0""" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"" align=""center"">" & vbCrLf
Response.Write "<br><font style=""font-size:14px"" class=""book""><b>ѧУ�ſ�</b></font></div><br>" & vbCrLf
Response.Write "<table width=""98%"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"">" & vbCrLf
Response.Write "<div style=""LINE-HEIGHT: 180%"">"&rs("gy")&"</div>" & vbCrLf
Response.Write "</td></tr>" & vbCrLf
Response.Write "</table>"
elseif ly="lx" then
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""��ϵ����"");</SCRIPT><table width=""100%"" align=""center"" cellspacing=""0"" cellpadding=""0""" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"" align=""center"">" & vbCrLf
Response.Write "<br><font style=""font-size:14px"" class=""book""><b>��ϵ����</b></font></div><br>" & vbCrLf
Response.Write "<table width=""98%"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"">" & vbCrLf
Response.Write "<div style=""LINE-HEIGHT: 180%"">"&rs("lx")&"</div>" & vbCrLf
Response.Write "</td></tr>" & vbCrLf
Response.Write "</table>"
elseif ly="gg" then
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""�쵼�´�"");</SCRIPT><table width=""100%"" align=""center"" cellspacing=""0"" cellpadding=""0""" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"" align=""center"">" & vbCrLf
Response.Write "<br><font style=""font-size:14px"" class=""book""><b>�쵼�´�</b></font></div><br>" & vbCrLf
Response.Write "<table width=""98%"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"">" & vbCrLf
Response.Write "<div style=""LINE-HEIGHT: 180%"">"&rs("gg")&"</div>" & vbCrLf
Response.Write "</td></tr>" & vbCrLf
Response.Write "</table>"
elseif ly="dt" then
Response.Write "<SCRIPT type=text/javascript>zonelovetable(""��֯����"");</SCRIPT><table width=""100%"" align=""center"" cellspacing=""0"" cellpadding=""0""" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"" align=""center"">" & vbCrLf
Response.Write "<br><font style=""font-size:14px"" class=""book""><b>��֯����</b></font></div><br>" & vbCrLf
Response.Write "<table width=""98%"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""word-break:break-all;table-layout:fixed;text-align:left"">" & vbCrLf
Response.Write "<tr><td valign=""top"">" & vbCrLf
Response.Write "<div style=""LINE-HEIGHT: 180%"">"&rs("dt")&"</div>" & vbCrLf
Response.Write "</td></tr>" & vbCrLf
Response.Write "</table>"
rs.close
set rs=nothing
end if
call footer()%>


</body>
</html>
