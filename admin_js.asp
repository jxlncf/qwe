<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newjscat" then
sql="select * from jscat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim jscatname
jscatname=trim(replace(request.form("jscat_name"),"'",""))
if jscatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("jscat_name")=jscatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_js.asp?action=jscat"
end if
elseif request.form("action")="newjs" then
sql="select * from js"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim jscatid,jsname,jsurl,jsdesc,jsisbest
jscatid=cint(request.form("jscatid"))
jsname=request.form("name")
jsdesc=request.form("desc")
jsisbest=request.form("isbest")
if jsname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写文章名称！');history.back(1);</script>")
else
  rs("js_name")=jsname
end if
if jscatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择文章所属分类！');history.back(1);</script>")
else
  rs("jscat_id")=jscatid
end if
if jsdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写文章脚本！');history.back(1);</script>")
else
  rs("js_desc")=jsdesc
end if
if cint(jsisbest)=1 then
  rs("isbest")=cint(jsisbest)
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  sql="update allcount set jscount = jscount + 1"
  conn.execute(sql)
  
  response.redirect "admin_js.asp?action=js"
end if
elseif request.form("action")="editjs" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from js where js_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
jscatid=cint(request.form("jscatid"))
jsname=request.form("name")
jsdesc=request.form("desc")
jsisbest=request.form("isbest")
if jsname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写文章名称！');history.back(1);</script>")
else
  rs("js_name")=jsname
end if
if jscatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择文章所属分类！');history.back(1);</script>")
else
  rs("jscat_id")=jscatid
end if
if jsdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写文章脚本！');history.back(1);</script>")
else
  rs("js_desc")=jsdesc
end if
rs("isbest")=cint(jsisbest)

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_js.asp?action=js"
end if
elseif request.form("action")="deljs" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from js where js_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  sql="update allcount set jscount = jscount - 1"
  conn.execute(sql)
  
  response.redirect "admin_js.asp?action=js"
elseif request.form("action")="editjscat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from jscat where jscat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
jscatname=trim(replace(request.form("jscat_name"),"'",""))
if jscatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("jscat_name")=jscatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_js.asp?action=jscat"
end if
elseif request.form("action")="deljscat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from jscat where jscat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_js.asp?action=jscat"
end if
end if
sql="select * from jscat order by jscat_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
<%if request.querystring("action")="jscat" then%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
  <tr class=zonelovess> <td colspan="3">文章分类管理</td></tr>
  <tr align="center" class=zoneloveqs><td width="10%">编号</td><td width="70%">分类名称</td><td width="20%">操作</td></tr>
<%do while not rs.eof%>
  <tr class=zoneloveds><td align="center"><%=rs("jscat_id")%></td><td><a href="#"><%=rs("jscat_name")%></a></td><td align="center"><a href="admin_js.asp?id=<%=rs("jscat_id")%>&action=editjscat">编辑</a> <a href="admin_js.asp?id=<%=rs("jscat_id")%>&action=deljscat">删除</a> <a href="js.asp?jscat_id=<%=rs("jscat_id")%>" target="_blank">查看</a></td></tr>
<%rs.movenext
loop
if rs.bof and rs.eof then%>
   <tr align="center" class=zoneloveds><td colspan="3">当前没有分类！</td></tr>
<%rs.close
set rs=nothing
end if%>
</table>
<%end if
if request.querystring("action")="newjscat" then%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="">
<tr class=zonelovess><td >新增分类</td></tr>
<tr class=zoneloveds><td>分类名称-<input type="text" name="jscat_name" size="40" class="textarea"></td></tr>
<tr class=zoneloveqs><td align="center"><input type="submit" name="Submit" value="确定新增" class="button"><input type="reset" name="Reset" value="清空重填" class="button"></td></tr><input type="hidden" name="action" value="newjscat"><input type="hidden" name="MM_insert" value="true"></form>
</table>
<%end if
if request.QueryString("action")="editjscat" then
if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
response.end
else
if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
response.end
end if
end if
sql="select * from jscat where jscat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="">
<tr class=zonelovess><td>修改分类</td></tr>
<tr class=zoneloveds><td>分类名称-<input name="jscat_name" type="text" class="textarea" id="jscat_name" size="40" value="<%=rs("jscat_name")%>"> </td></tr>
<tr class=zoneloveqs><td align="center"><input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> <input name="Reset" type="reset" class="button" id="Reset" value="清空重填"> </td></tr><input type="hidden" name="id" value="<%=rs("jscat_id")%>"><input type="hidden" name="action" value="editjscat"><input type="hidden" name="MM_insert" value="true"></form>
</table>
<%end if
	  if request.QueryString("action")="deljscat" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from jscat where jscat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="">
<tr class=zonelovess><td>删除分类</td></tr>
<tr class=zoneloveds><td>分类名称- <%=rs("jscat_name")%></td></tr>
<tr class=zoneloveqs><td align="center"><input name="Submit" type="submit" class="button" id="Submit" value="确定删除"></td></tr><input type="hidden" name="id" value="<%=rs("jscat_id")%>"><input type="hidden" name="action" value="deljscat"><input type="hidden" name="MM_insert" value="true"></form>
</table>
<%end if
if request.querystring("action")="js" then
sql="select * from js order by js_id desc"
if request.querystring("jscat_id")<>"" then
sql="select * from js where jscat_id="&cint(request.querystring("jscat_id"))&" order by js_id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
totaljs=rs.recordcount
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form3" method="post" action="">
<tr class=zonelovess><td align="right" colspan="3"><select style="margin:-3px" name="go" onChange='window.location=form.go.options[form.go.selectedIndex].value'><option value="">选择显示方式</option><option value="admin_js.asp?action=js">显示所有文章</option>
<%sql="select * from jscat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_js.asp?action=js&jscat_id=<%=rs2("jscat_id")%>"><%=rs2("jscat_name")%></option>
<%rs2.movenext
loop
if rs2.bof and rs2.eof then%><option value="">当前没有分类</option>
<%end if
rs2.close
set rs2=nothing%></select></td></tr></form><%
if not rs.eof then
rs.movefirst
rs.pagesize=jsperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totaljs=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*jsperpage<totaljs then
	       rs.move(currentpage-1)*jsperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totaljs mod jsperpage)=0 then
      totalpages=totaljs\jsperpage
   else
      totalpages=totaljs\jsperpage+1
   end if
i=0
do while not rs.eof and i<jsperpage
%>
<tr class=zoneloveds><td width="70%">·<%=rs("js_name")%></td><td width="10%">评论：<font color=red><%=rs("reviewcount")%></font></td><td width="20%"> <a href="admin_view.asp?action=js&js_id=<%=rs("js_id")%>">评论</a> <a href="admin_js.asp?id=<%=rs("js_id")%>&action=editjs">编辑</a> <a href="admin_js.asp?id=<%=rs("js_id")%>&action=deljs">删除</a></td>
</tr>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
<tr class=zoneloveds>
  <td colspan="3" height="22" align="center">当前没有文章！</td>
</tr>
<%end if
end if%>
<form name="form1" method="post" action="admin_js.asp?action=js&jscat_id=<%=request.querystring("jscat_id")%>">
<tr class=zoneloveqs><td align="right" colspan="3"> <%=currentpage%> /<%=totalpages%>页,<%=totaljs%>条记录/<%=jsperpage%>篇每页. 
              <%
i=1
showye=totalpages
if showye>10 then
showye=10
end if
for i=1 to showye
if i=currentpage then
%>
<%=i%>
<%else%>
<a href="admin_js.asp?action=js&page=<%=i%>&jscat_id=<%=request.querystring("jscat_id")%>"><%=i%></a> 
<%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_js.asp?action=js&page=<%=page%>&jscat_id=<%=request.querystring("jscat_id")%>" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="newjs" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form2" method="post" action="">
<tr class=zonelovess><td colspan="2">新增文章</td></tr>
<tr class=zoneloveds><td width="17%">文章名称</td><td width="83%"><input type="text" name="name" class="textarea" size="40"></td></tr>
<tr class=zoneloveds><td width="17%">所属分类</td><td width="83%"><select name="jscatid" class="textarea">
<%
sql="select * from jscat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof%><option value="<%=rs("jscat_id")%>"><%=rs("jscat_name")%></option>
<%rs.movenext
loop
if rs.eof and rs.bof then%><option value="0">当前没有分类</option>
<%end if%>
</select></td></tr>
              
<tr class=zoneloveds><td width="17%">内容:</td><td width="83%"><textarea name="desc" class="textarea" cols="65" rows="15"></textarea></td></tr>
<tr align="center" class=zoneloveqs><td colspan="2"><input type="checkbox" name="isbest" value="1" class="textarea">重要&nbsp;&nbsp;&nbsp; <input type="submit" name="Submit" value="确定新增" class="button"><input type="reset" name="Reset" value="清空重填" class="button">[<a href="admin_js.asp?action=js">返回</a>]</td></tr><input type="hidden" name="action" value="newjs"><input type="hidden" name="MM_insert" value="true"></form>
</table>
      <%end if
	  if request.QueryString("action")="editjs" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from js where js_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form2" method="post" action="">
<tr class=zonelovess><td colspan="2" >修改文章</td></tr>
<tr class=zoneloveds><td width="17%">文章名称</td><td width="83%"> <input name="name" type="text" class="textarea" id="name" size="40" value="<%=rs("js_name")%>"></td></tr>
<tr class=zoneloveds><td width="17%">所属分类</td><td width="83%"><select name="jscatid" class="textarea" id="jscatid">
                <%
sql="select * from jscat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
                <option value="<%=rs2("jscat_id")%>" <%if rs2("jscat_id")=rs("jscat_id") then response.write "selected" end if%>><%=rs2("jscat_name")%></option>
                <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
                <option value="0">当前没有分类</option>
                <%end if%>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">内容:</td>
            <td width="83%"> 
              <textarea name="desc" cols="65" rows="15" class="textarea" id="desc"><%=CODEStr(rs("js_desc"))%></textarea>
            </td>
          </tr>
          <tr align="center" class=zoneloveqs> 
            <td colspan="2"> <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              重要
              <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_js.asp?action=js">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("js_id")%>">
          <input type="hidden" name="action" value="editjs">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="deljs" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from js where js_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
<tr class=zonelovess><td colspan="2">删除文章</td></tr>
          <tr class=zoneloveds> 
            <td width="17%">文章名称</td>
            <td width="83%"> <%=rs("js_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">所属分类</td>
            <td width="83%"> 
                <%
sql="select * from jscat where jscat_id="&rs("jscat_id")
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
%><%=rs2("jscat_name")%> 　</td>
          </tr>
          <tr align="center" class=zoneloveqs> 
            <td colspan="2"> 
              <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              重要 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
              [<a href="admin_js.asp?action=js">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("js_id")%>">
          <input type="hidden" name="action" value="deljs">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
end if
%>