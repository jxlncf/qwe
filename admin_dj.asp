<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newdjcat" then
sql="select * from djcat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim djcatname
djcatname=trim(replace(request.form("djcat_name"),"'",""))
if djcatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("djcat_name")=djcatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_dj.asp?action=djcat"
end if
elseif request.form("action")="newdj" then
sql="select * from dj"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim djcatid,djname,djurl,djpic,djdesc,djuser,djisbest
djcatid=cint(request.form("djcatid"))
djname=trim(replace(request.form("name"),"'",""))
djuser=trim(replace(request.form("user"),"'",""))
djurl=trim(replace(request.form("url"),"'",""))
djpic=trim(replace(request.form("pic"),"'",""))
djdesc=trim(replace(request.form("desc"),"'",""))
djisbest=request.form("isbest")
if djname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐名称！');history.back(1);</script>")
else
  rs("dj_name")=djname
end if
if djcatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择所属分类！');history.back(1);</script>")
else
  rs("djcat_id")=djcatid
end if
if djurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐地址！');history.back(1);</script>")
else
  rs("dj_url")=djurl
end if
if djpic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的播放类型！');history.back(1);</script>")
else
  rs("dj_pic")=djpic
end if
if djdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的推荐等级！');history.back(1);</script>")
else
  rs("dj_desc")=djdesc
end if
if djuser="" then
  rs("dj_user")="未知"
else
  rs("dj_user")=djuser
end if
if cint(djisbest)=1 then
  rs("isbest")=cint(djisbest)
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  sql="update allcount set djcount = djcount + 1"
  conn.execute(sql)
  response.redirect "admin_dj.asp?action=dj"
end if
elseif request.form("action")="editdj" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的音乐id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from dj where dj_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
djcatid=cint(request.form("djcatid"))
djname=trim(replace(request.form("name"),"'",""))
djuser=trim(replace(request.form("user"),"'",""))
djurl=trim(replace(request.form("url"),"'",""))
djpic=trim(replace(request.form("pic"),"'",""))
djdesc=trim(replace(request.form("desc"),"'",""))
djisbest=request.form("isbest")
error=request.form("error")
if djname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐名称！');history.back(1);</script>")
else
  rs("dj_name")=djname
end if
if djcatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐所属分类！');history.back(1);</script>")
else
  rs("djcat_id")=djcatid
end if
if djurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐地址！');history.back(1);</script>")
else
  rs("dj_url")=djurl
end if
if djpic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的播放类型！');history.back(1);</script>")
else
  rs("dj_pic")=djpic
end if
if djdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的推荐等级！');history.back(1);</script>")
else
  rs("dj_desc")=djdesc
end if
if djuser="" then
  rs("dj_user")="未知"
else
  rs("dj_user")=djuser
end if
rs("isbest")=cint(djisbest)
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_dj.asp?action=dj"
end if
elseif request.form("action")="editerror" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的音乐id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from dj where dj_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
djcatid=cint(request.form("djcatid"))
djname=trim(replace(request.form("name"),"'",""))
djuser=trim(replace(request.form("user"),"'",""))
djurl=trim(replace(request.form("url"),"'",""))
djpic=trim(replace(request.form("pic"),"'",""))
djdesc=trim(replace(request.form("desc"),"'",""))
djisbest=request.form("isbest")
error=request.form("error")
if djname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐名称！');history.back(1);</script>")
else
  rs("dj_name")=djname
end if
if djcatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐所属分类！');history.back(1);</script>")
else
  rs("djcat_id")=djcatid
end if
if djurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写音乐地址！');history.back(1);</script>")
else
  rs("dj_url")=djurl
end if
if djpic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的播放类型！');history.back(1);</script>")
else
  rs("dj_pic")=djpic
end if
if djdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择音乐的推荐等级！');history.back(1);</script>")
else
  rs("dj_desc")=djdesc
end if
if djuser="" then
  rs("dj_user")="未知"
else
  rs("dj_user")=djuser
end if
rs("isbest")=cint(djisbest)
rs("error")=0
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_dj.asp?action=djerror"
end if
elseif request.form("action")="deldj" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的音乐id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from dj where dj_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  sql="update allcount set djcount = djcount - 1"
  conn.execute(sql)

    response.redirect "admin_dj.asp?action=dj"
 elseif request.form("action")="editdjcat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的音乐分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from djcat where djcat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
djcatname=trim(replace(request.form("djcat_name"),"'",""))
if djcatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("djcat_name")=djcatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_dj.asp?action=djcat"
end if
elseif request.form("action")="deldjcat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的音乐分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from djcat where djcat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_dj.asp?action=djcat"
end if
end if%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
<%if request.querystring("action")="djerror" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zonelovess> 
<td colspan="3">
错误连接管理</td>
</tr>
<tr class=zoneloveqs> 
<td width="60%">音乐名称</td>
<td width="20%" align="center">报错人次</td>
<td width="20%" align="center">操作</td>
</tr>
<%
sql="select * from dj where error > 0 order by dj_id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%><tr class=zoneloveds><td><a href="#"><%=rs("dj_name")%></td>
<td align="center"><%=rs("error")%></td>
            <td  align="center"><a href="admin_dj.asp?id=<%=rs("dj_id")%>&action=editerror">修改</a> <a href="admin_dj.asp?id=<%=rs("dj_id")%>&action=deldj">删除</a></td>
        </tr>
<%rs.movenext
loop
if rs.eof and rs.bof then
Response.Write("<tr class=zoneloveds><td align=""center"" colspan=""3"">当前还没有错误音乐</td></tr>")
end if%>
</table>
<%end if%>
     <%if request.querystring("action")="djcat" then
sql="select * from djcat order by djcat_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr class=zonelovess> 
          <td colspan="3">
          音乐分类管理</td>
        </tr>
        <tr align="center" class=zoneloveqs> 
          <td width="10%">编号</td>
          <td width="70%">分类名称</td>
          <td width="20%">操作</td>
        </tr>
        <%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("djcat_id")%>　</td>
          <td><a href="dj.asp?djcat_id=<%=rs("djcat_id")%>" target="_blank"><%=rs("djcat_name")%></a>　</td>
          <td align="center"><a href="admin_dj.asp?id=<%=rs("djcat_id")%>&action=editdjcat">编辑</a> 
            <a href="admin_dj.asp?id=<%=rs("djcat_id")%>&action=deldjcat">删除</a> <a href="dj.asp?djcat_id=<%=rs("djcat_id")%>" target="_blank">查看</a></td>
        </tr>
        <%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr align="center" class=zoneloveds> 
          <td colspan="3">当前没有音乐分类！</td>
        </tr>
        <%rs.close
set rs=nothing
		end if%>
      </table>
      <br>
     <%end if
	 if request.querystring("action")="newdjcat" then%> 
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action=""><tr class=zonelovess><td>新增音乐分类</td></tr>
        <tr class=zoneloveds> 
            <td>分类名称- 
              <input type="text" name="djcat_name" size="40" class="textarea">
          </td>
        </tr>
        <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=djcat'" class="button">
            </td>
        </tr>
		<input type="hidden" name="action" value="newdjcat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.QueryString("action")="editdjcat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的音乐分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from djcat where djcat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>
            修改音乐分类</td>
          </tr>
          <tr class=zoneloveds> 
            <td>分类名称- 
              <input name="djcat_name" type="text" class="textarea" id="djcat_name" size="40" value="<%=rs("djcat_name")%>"> 
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
               <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=djcat'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("djcat_id")%>">
          <input type="hidden" name="action" value="editdjcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="deldjcat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的音乐分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from djcat where djcat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
    <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>
            删除音乐分类</td>
          </tr>
          <tr class=zoneloveds> 
            <td>分类名称- <%=rs("djcat_name")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
<input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=djcat'" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("djcat_id")%>">
          <input type="hidden" name="action" value="deldjcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="dj" then
sql="select * from dj order by dj_id desc"
if request.querystring("djcat_id")<>"" then
sql="select * from dj where djcat_id="&cint(request.querystring("djcat_id"))&" order by dj_id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
totaldj=rs.recordcount
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form name="form3" method="post" action=""><tr class=zonelovess> 
<td colspan="4">音乐管理</td>
<td align="right" style="padding-top:2px;">
              <select style="margin:-3px" name="go" onChange='window.location=form.go.options[form.go.selectedIndex].value'>
			  <option value="">选择显示方式</option>
			  <option value="admin_dj.asp?action=dj">显示所有音乐</option>
<%sql="select * from djcat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_dj.asp?action=dj&djcat_id=<%=rs2("djcat_id")%>"><%=rs2("djcat_name")%></option>
<%rs2.movenext
loop
if rs2.bof and rs2.eof then%><option value="">当前没有分类</option>
<%end if
rs2.close
set rs2=nothing%>
              </select>
            </td>
		
</tr></form>
<tr class=zoneloveqs> 
<td width="40%">音乐名称</td>
<td width="20%" align="center">加入时间</td>
<td width="10%" align="center">评论</td>
<td width="10%" align="center">点击</td>
<td width="20%" align="center">操作</td>
</tr>
<%
if not rs.eof then
rs.movefirst
rs.pagesize=djperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totaldj=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*djperpage<totaldj then
	       rs.move(currentpage-1)*djperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totaldj mod djperpage)=0 then
      totalpages=totaldj\djperpage
   else
      totalpages=totaldj\djperpage+1
   end if
i=0
do while not rs.eof and i<djperpage
%>
<tr class=zoneloveds> 
<td><a href="#"><%=rs("dj_name")%></td>
<td  align="center"><%=rs("dj_date")%></td>
<td align="center"><font color=red><%=rs("reviewcount")%></font></td>
<td align="center"><%=rs("dj_count")%></td>
<td  align="center"> <a href="admin_view.asp?action=dj&dj_id=<%=rs("dj_id")%>">评论</a> <a href="admin_dj.asp?id=<%=rs("dj_id")%>&action=editdj">编辑</a> <a href="admin_dj.asp?id=<%=rs("dj_id")%>&action=deldj">删除</a></td>
        </tr>     
      <%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr class=zoneloveds> 
          <td colspan="5" height="22" align="center">当前没有音乐！</td>
        </tr>
      <%end if
end if%>
        <form name="form1" method="post" action="admin_dj.asp?action=dj&djcat_id=<%=request.querystring("djcat_id")%>">
          <tr class=zoneloveqs> 
            <td align="right" colspan="5"> <%=currentpage%> /<%=totalpages%>页,<%=totaldj%>条记录/<%=djperpage%>篇每页. 
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
              <a href="admin_dj.asp?action=dj&page=<%=i%>&djcat_id=<%=request.querystring("djcat_id")%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_dj.asp?action=dj&page=<%=page%>&djcat_id=<%=request.querystring("djcat_id")%>" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>

	  <%end if
if request.querystring("action")="newdj" then%>
     <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            新增音乐</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐名称</td>
            <td width="83%"> 
              <input type="text" name="name" size="40">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐作者</td>
            <td width="83%"> 
              <input type="text" name="user" size="40">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐选项</td>
            <td width="83%"> 
              <select name="djcatid">
<%
sql="select * from djcat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof%><option value="<%=rs("djcat_id")%>"><%=rs("djcat_name")%></option>
<%rs.movenext
loop
if rs.eof and rs.bof then%><option value="0">当前没有音乐分类</option>
<%end if%>
</select>
<select name="pic">
<option value="rm">播放类型</option>
<option value="asf">ASF音乐</option>
<option value="asfmtv">ASF影视</option>
<option value="rm">RAM音乐</option>
<option value="rmmtv">RAM影视</option>
<option value="flash">FLASH动画</option>
              </select>
<select name="desc">
<option value="☆☆★★★">推荐等级</option>
<option value="☆☆☆☆★">☆☆☆☆★</option>
<option value="☆☆☆★★">☆☆☆★★</option>
<option value="☆☆★★★">☆☆★★★</option>
<option value="☆★★★★">☆★★★★</option>
<option value="★★★★★">★★★★★</option>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">视听地址</td>
            <td width="83%"> 
              <input type="text" name="url" size="60">(推荐使用)
            </td>
          </tr>
<tr class=zoneloveds> 
              <td> 上传文件： </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=dj"></iframe>(文件500K以内可使用)</td>
            </tr>       
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="checkbox" name="isbest" value="1">
              推荐&nbsp;&nbsp;&nbsp; 
              <input type="submit" name="Submit" value="确定新增" class="button">
             <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=dj'" class="button"></td>
          </tr>
		  <input type="hidden" name="action" value="newdj">
		<input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<br>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zonelovess> 
<td>类型说明</td>
</tr>
<tr  class=zoneloveds>
<td>1."<FONT COLOR="#FF0000">.mp3.mid.wma.asf"纯音乐文件请选择<FONT COLOR="#0000FF">ASF音乐类型<br>2."<FONT COLOR="#FF0000">.wmv.asf.mpg.avi"影视文件选择<FONT COLOR="#0000FF">ASF影视类型.<br>3."<FONT COLOR="#FF0000">.ram.rm"纯音乐文件请选择<FONT COLOR="#0000FF">RAM音乐类型<br>4."<FONT COLOR="#FF0000">.ram.rm"影视文件请选择<FONT COLOR="#0000FF">RAM影视类型<br>5."<FONT COLOR="#FF0000">.swf"动画文件请选择<FONT COLOR="#0000FF">FLASH动画类型
</td></tr></table>
      <%end if
	  if request.QueryString("action")="editdj" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的音乐ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from dj where dj_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            修改音乐</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐名称</td>
            <td width="83%"> 
              <input name="name" type="text" class="textarea" id="name" size="40" value="<%=rs("dj_name")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐作者</td>
            <td width="83%"> 
              <input name="user" type="text" class="textarea" id="user" size="40" value="<%=rs("dj_user")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐分类</td>
            <td width="83%"> 
              <select name="djcatid" id="djcatid">
                <%
sql="select * from djcat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
                <option value="<%=rs2("djcat_id")%>" <%if rs2("djcat_id")=rs("djcat_id") then response.write "selected" end if%>><%=rs2("djcat_name")%></option>
                <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
                <option value="0">当前没有音乐分类</option>
                <%end if%>
              </select>
<select name="pic">
<option value="asf" <%if rs("dj_pic")="asf" then response.write "selected" end if%>>ASF音乐</option>
<option value="asfmtv" <%if rs("dj_pic")="asfmtv" then response.write "selected" end if%>>ASF影视</option>
<option value="rm" <%if rs("dj_pic")="rm" then response.write "selected" end if%>>RAM音乐</option>
<option value="rmmtv" <%if rs("dj_pic")="rmmtv" then response.write "selected" end if%>>RAM影视</option>
<option value="flash" <%if rs("dj_pic")="flash" then response.write "selected" end if%>>FLASH动画</option>
</select>
<select name="desc">
<option value="☆☆☆☆★">☆☆☆☆★</option>
<option value="☆☆☆★★">☆☆☆★★</option>
<option value="☆☆★★★" selected>☆☆★★★</option>
<option value="☆★★★★">☆★★★★</option>
<option value="★★★★★">★★★★★</option>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td  width="17%">视听地址</td>
            <td  width="83%"> 
              <input name="url" type="text" id="url" size="60" value="<%=rs("dj_url")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td> 上传文件： </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=dj"></iframe></td>
            </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> <input name="isbest" type="checkbox"  id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              推荐
              <input name="Submit" type="submit" id="Submit" value="确定修改"> 
 <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=dj'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("dj_id")%>">
          <input type="hidden" name="action" value="editdj">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<br>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zonelovess> 
<td>类型说明</td>
</tr>
<tr  class=zoneloveds>
<td>1."<FONT COLOR="#FF0000">.mp3.mid.wma.asf"纯音乐文件请选择<FONT COLOR="#0000FF">ASF音乐类型<br>2."<FONT COLOR="#FF0000">.wmv.asf.mpg.avi"影视文件选择<FONT COLOR="#0000FF">ASF影视类型.<br>3."<FONT COLOR="#FF0000">.ram.rm"纯音乐文件请选择<FONT COLOR="#0000FF">RAM音乐类型<br>4."<FONT COLOR="#FF0000">.ram.rm"影视文件请选择<FONT COLOR="#0000FF">RAM影视类型<br>5."<FONT COLOR="#FF0000">.swf"动画文件请选择<FONT COLOR="#0000FF">FLASH动画类型
</td></tr></table>
	  <%end if
if request.QueryString("action")="editerror" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的音乐ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from dj where dj_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            修改音乐</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐名称</td>
            <td width="83%"> 
              <input name="name" type="text" class="textarea" id="name" size="40" value="<%=rs("dj_name")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐作者</td>
            <td width="83%"> 
              <input name="usere" type="text" class="textarea" id="user" size="40" value="<%=rs("dj_user")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐分类</td>
            <td width="83%"> 
              <select name="djcatid" id="djcatid">
                <%
sql="select * from djcat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
                <option value="<%=rs2("djcat_id")%>" <%if rs2("djcat_id")=rs("djcat_id") then response.write "selected" end if%>><%=rs2("djcat_name")%></option>
                <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
                <option value="0">当前没有音乐分类</option>
                <%end if%>
              </select>
<select name="pic">
<option value="asf" <%if rs("dj_pic")="asf" then response.write "selected" end if%>>ASF音乐</option>
<option value="asfmtv" <%if rs("dj_pic")="asfmtv" then response.write "selected" end if%>>ASF影视</option>
<option value="rm" <%if rs("dj_pic")="rm" then response.write "selected" end if%>>RAM音乐</option>
<option value="rmmtv" <%if rs("dj_pic")="rmmtv" then response.write "selected" end if%>>RAM影视</option>
<option value="flash" <%if rs("dj_pic")="flash" then response.write "selected" end if%>>FLASH动画</option>
</select>
<select name="desc">
<option value="☆☆☆☆★">☆☆☆☆★</option>
<option value="☆☆☆★★">☆☆☆★★</option>
<option value="☆☆★★★" selected>☆☆★★★</option>
<option value="☆★★★★">☆★★★★</option>
<option value="★★★★★">★★★★★</option>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td  width="17%">视听地址</td>
            <td  width="83%"> 
              <input name="url" type="text" id="url" size="60" value="<%=rs("dj_url")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td> 上传文件： </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=dj"></iframe></td>
            </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> <input name="isbest" type="checkbox"  id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              推荐
              <input name="Submit" type="submit" id="Submit" value="确定修改"> 
 <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=djerror'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("dj_id")%>">
          <input type="hidden" name="action" value="editerror">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<br>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zonelovess> 
<td>类型说明</td>
</tr>
<tr  class=zoneloveds>
<td>1."<FONT COLOR="#FF0000">.mp3.mid.wma.asf"纯音乐文件请选择<FONT COLOR="#0000FF">ASF音乐类型<br>2."<FONT COLOR="#FF0000">.wmv.asf.mpg.avi"影视文件选择<FONT COLOR="#0000FF">ASF影视类型.<br>3."<FONT COLOR="#FF0000">.ram.rm"纯音乐文件请选择<FONT COLOR="#0000FF">RAM音乐类型<br>4."<FONT COLOR="#FF0000">.ram.rm"影视文件请选择<FONT COLOR="#0000FF">RAM影视类型<br>5."<FONT COLOR="#FF0000">.swf"动画文件请选择<FONT COLOR="#0000FF">FLASH动画类型
</td></tr></table>
	  <%end if
	  if request.QueryString("action")="deldj" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的音乐ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from dj where dj_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除音乐</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐名称</td>
            <td width="83%"> <%=rs("dj_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐作者</td>
            <td width="83%"> <%=rs("dj_user")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">音乐分类</td>
            <td width="83%"> 
                <%
sql="select * from djcat where djcat_id="&rs("djcat_id")
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
%><%=rs2("djcat_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">视听地址</td>
            <td width="83%"><%=rs("dj_url")%> 　</td>
          </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input name="isbest" type="checkbox" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              推荐 
              <input name="Submit" type="submit" id="Submit" value="确定删除">
 <input type="button" value="返回"  onClick="location.href='admin_dj.asp?action=dj'" class="button"></td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("dj_id")%>">
          <input type="hidden" name="action" value="deldj">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
end if
%>