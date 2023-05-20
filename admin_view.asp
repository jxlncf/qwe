<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.QueryString("action")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
end if
dim i
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<%
if request.QueryString("action")="art" then
set rs=server.createobject("adodb.recordset")
sql="select * from art where art_id="&request.querystring("art_id")
rs.open sql,conn,1,1
%>

              <form name="form1" method="post" action="admin_art.asp?action=list">
                <tr class=zonelovess> 
                  <td>管理 <%=rs("art_title")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=art&art_id=<%=request("art_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
if request.QueryString("action")="news" then
set rs=server.createobject("adodb.recordset")
sql="select * from news where news_id="&request.querystring("news_id")
rs.open sql,conn,1,1%>
<tr class=zonelovess> 
                  <td>管理 <%=rs("news_title")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=news&news_id=<%=request("news_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
if request.QueryString("action")="dj" then
set rs=server.createobject("adodb.recordset")
sql="select * from dj where dj_id="&request.querystring("dj_id")
rs.open sql,conn,1,1%>
<tr class=zonelovess> 
                  <td>管理 <%=rs("dj_name")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=dj&dj_id=<%=request("dj_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
if request.QueryString("action")="down" then
set rs=server.createobject("adodb.recordset")
sql="select * from soft where soft_id="&request.querystring("soft_id")
rs.open sql,conn,1,1
%>
<tr class=zonelovess> 
                  <td>管理 <%=rs("soft_name")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=down&soft_id=<%=request("soft_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
if request.QueryString("action")="js" then
set rs=server.createobject("adodb.recordset")
sql="select * from js where js_id="&request.querystring("js_id")
rs.open sql,conn,1,1%>
                <tr class=zonelovess> 
                  <td>管理 <%=rs("js_name")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=js&js_id=<%=request("js_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
if request.QueryString("action")="pic" then
set rs=server.createobject("adodb.recordset")
sql="select * from pic where pic_id="&request.querystring("pic_id")
rs.open sql,conn,1,1
%>
<tr class=zonelovess> 
                  <td>管理 <%=rs("pic_name")%> 的相关评论文章 共：<%=rs("reviewcount")%> 篇</font></td>
                </tr>
<tr class=zoneloveds><td>
<%if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)%>
<%=temp(i)%> 
 <div align="right"><a href="admin_viewok.asp?action=pic&pic_id=<%=request("pic_id")%>&ping_id=<%=i%>"><font color="#CC0000">[删除]</font></a> </div>
<%next
else
Response.Write "<br><p align=center>当前没有评论</p>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
  </tr>
</table>
<%end if
end if%>