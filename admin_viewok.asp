<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request("ping_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
if request("action")="" then
 Response.Write("<script language=javascript>alert('请指定操作的对象!');javascript:history.back();</script>")
   response.end
end if
dim i
dim temp
dim temp_review
dim ping_id
ping_id=request("ping_id")
if request.QueryString("action")="art" then
if request("art_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
art_id=request("art_id")
set rs=server.createobject("adodb.recordset")
sql="select * from art where art_id="&request.querystring("art_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update art set reviewcount=reviewcount-1,review='"&temp_review&"' where art_id="&request.querystring("art_id")
rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=art&art_id="&art_id
end if
if request.QueryString("action")="dj" then
if request("dj_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
dj_id=request("dj_id")
set rs=server.createobject("adodb.recordset")
sql="select * from dj where dj_id="&request.querystring("dj_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update dj set reviewcount=reviewcount-1,review='"&temp_review&"' where dj_id="&request.querystring("dj_id")

rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=dj&dj_ID="&dj_id
end if
if request.QueryString("action")="down" then
if request("soft_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
soft_id=request("soft_id")
set rs=server.createobject("adodb.recordset")
sql="select * from soft where soft_id="&request.querystring("soft_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update soft set reviewcount=reviewcount-1,review='"&temp_review&"' where soft_id="&request.querystring("soft_id")

rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=down&soft_ID="&soft_id
end if
if request.QueryString("action")="js" then
if request("js_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
js_id=request("js_id")
set rs=server.createobject("adodb.recordset")
sql="select * from js where js_id="&request.querystring("js_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update js set reviewcount=reviewcount-1,review='"&temp_review&"' where js_id="&request.querystring("js_id")

rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=js&js_ID="&js_id
end if
if request.QueryString("action")="news" then
if request("news_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
news_id=request("news_id")
set rs=server.createobject("adodb.recordset")
sql="select * from news where news_id="&request.querystring("news_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update news set reviewcount=reviewcount-1,review='"&temp_review&"' where news_id="&request.querystring("news_id")

rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=news&news_ID="&news_id
end if
if request.QueryString("action")="pic" then
if request("pic_id")="" then
 Response.Write("<script language=javascript>alert('参数错误!!');javascript:history.back();</script>")
   response.end
end if
pic_id=request("pic_id")
set rs=server.createobject("adodb.recordset")
sql="select * from pic where pic_id="&request.querystring("pic_id")
rs.open sql,conn,1,1
if rs("review")<>"" then
temp=split(Trim(rs("review")),"|")
for i = 1 to ubound(temp)
if int(i)=int(ping_id) then
temp_review=temp_review
else
temp_review=temp_review&"|"&temp(i)
end if
next
else
 Response.Write("<script language=javascript>alert('目前还没有任何评论文章让您删除呀!');javascript:history.back();</script>")
   response.end
end if
conn.execute"Update pic set reviewcount=reviewcount-1,review='"&temp_review&"' where pic_id="&request.querystring("pic_id")

rs.close
conn.close
set rs=nothing
set conn=nothing
set temp=nothing
set temp_review=nothing
response.redirect "admin_view.asp?action=pic&pic_ID="&pic_id
end if
end if%>