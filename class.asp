<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/inc.asp"-->
<SCRIPT LANGUAGE="JavaScript">
function SetWinHeight(obj)
{
 var win=obj;
 if (document.getElementById)
 {
  if (win && !window.opera)
  {
   if (win.contentDocument && win.contentDocument.body.offsetHeight) 

    win.height = win.contentDocument.body.offsetHeight; 
   else if(win.Document && win.Document.body.scrollHeight)
    win.height = win.Document.body.scrollHeight;
  }
 }
}
</SCRIPT>
<%
start="成绩查询系统"
call head()
call menu()
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"">" & vbCrLf
Response.Write "<TBODY><TR vAlign=top bgcolor=""#FCFCFC""><TD>" & vbCrLf
Response.Write "<iframe width='760' id='win' name='win' onload='Javascript:SetWinHeight(this)' scrolling='no' border='0' frameborder='0' src='class/index_1.asp' ></iframe>" & vbCrLf

Response.Write "</td></tr>" & vbCrLf
Response.Write "</table>"

call footer()%>


</body>
</html>
