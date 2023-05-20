<!--#include file="inc/config.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="mdb.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="modpass" then
dim oldname,adminname,oldpwd,adminpwd,confirm
oldname=md5(trim(replace(request.form("oldname"),"'","")))
adminname=md5(trim(replace(request.form("adminname"),"'","")))
oldpwd=md5(trim(replace(request.form("oldpwd"),"'","")))
adminpwd=md5(trim(replace(request.form("adminpwd"),"'","")))
confirm=md5(trim(replace(request.form("confirm"),"'","")))
if oldname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入旧的管理员名称！');history.back(1);</script>")
end if
if oldpwd="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入旧的管理员密码！');history.back(1);</script>")
end if
if adminpwd="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入新的管理员密码！');history.back(1);</script>")
end if
if adminpwd<>confirm then
  founderr=true
  Response.Write("<script language=javascript>alert('你两次输入的管理员密码不相同！');history.back(1);</script>")
end if
if founderr then
  call diserror()
  response.end
else
sql="select * from admin where admin_name='"&oldname&"' and admin_password='"&oldpwd&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.eof then
 Response.Write("<script language=javascript>alert('管理员名称或密码错误！');history.back(1);</script>")
 call diserror()
 response.end
else
if adminname="" then
  rs("admin_name")=oldname
else
  rs("admin_name")=adminname
end if
rs("admin_password")=adminpwd
rs.update
rs.close
set rs=nothing
session("adminlogin")=""
session("issuper")=""
Response.Write("<script language=javascript>alert('修改成功！请重新登录');this.top.location.href='admin.asp';</script>")
end if
end if
end if
end if%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</head><body>
<%if request.querystring("action")="modpass" then%>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="admin_admin.asp">
<tr> 
    <td class=zonelovess align="center">修改资料</td>
          </tr>
          <tr class=zoneloveds> 
            <td>旧的管理员名称：<input type="text" name="oldname" size="30">
            </td>
          </tr>
<tr class=zoneloveds> 
            <td>新的管理员名称：<input type="text" name="adminname" size="30"> *不修改请留空
            </td>
          </tr>
<tr class=zoneloveds> 
            <td>旧的管理员密码：<input type="password" name="oldpwd" size="30">
            </td>
          </tr>
<tr class=zoneloveds> 
            <td>新的管理员密码：<input type="password" name="adminpwd" size="30">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td>确认管理员密码：<input type="password" name="confirm" size="30">
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="确定修改">
              <input type="reset" name="Reset" value="清空重写">
            </td>
          </tr>
          <input type="hidden" name="action" value="modpass">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
if request.querystring("action")="recount" then
dim newscount,articlecount,softcount,cscount,djcount
  sql="select news_id from news"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  newscount=rs.recordcount
  sql="select art_id from art"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  articlecount=rs.recordcount
  sql="select soft_id from soft"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  softcount=rs.recordcount
  sql="select cs_id from coolsites"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  cscount=rs.recordcount
  sql="select dj_id from dj"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  djcount=rs.recordcount
  sql="select * from allcount"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  rs("newscount")=newscount
  rs("articlecount")=articlecount
  rs("softcount")=softcount
  rs("coolsitescount")=cscount
  rs("djcount")=djcount
  rs.update
  rs.close
  set rs=nothing
  
%>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="admin_admin.asp">
<tr> 
    <td class=zonelovess>首页更新</td>
          </tr>
          <tr class=zoneloveds> 
    <td align="center" height="66">&gt;更新成功&lt;</td>
  </tr>
</table>
<%end if
end if%>