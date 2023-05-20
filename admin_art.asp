<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->

<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if Request.form("MM_insert") then
if request.Form("action")="newartcat" then
sql="select * from a_cat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim cat_name
cat_name=trim(replace(request.form("cat_name"),"'",""))
if cat_name="" then
  founderr=true
   Response.Write("<script language=javascript>alert('你必须填写文章分类');history.back(1);</script>")
else
  rs("cat_name")=cat_name
end if

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_art.asp"
end if
end if
if request.Form("action")="editartcat" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('非法的id参数');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from a_cat where cat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cat_name=trim(replace(request.form("cat_name"),"'",""))
if cat_name="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须填写文章分类');history.back(1);</script>")
else
  rs("cat_name")=cat_name
end if

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_art.asp"
end if
end if
if request.Form("action")="delartcat" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('非法的id参数');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from a_cat where cat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_art.asp"
end if
end if
sql="select * from a_cat order by cat_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
<%if request.QueryString("action")="" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr> 
          <td colspan="3" class=zonelovess>文章分类管理</font></td>
        </tr>
        <tr align="center"> 
          <td width="10%" class=zoneloveqs>分类ID号</td>
          <td width="70%" class=zoneloveqs>分类名称</td>
          <td width="20%" class=zoneloveqs>操作</td>
        </tr>
        <%do while not rs.eof%>
        <tr> 
          <td align="center" class=zoneloveds><%=rs("cat_id")%></td>
          <td class=zoneloveds><a href="#"><%=rs("cat_name")%></a></td>
          <td align="center" class=zoneloveds><a href="admin_art.asp?id=<%=rs("cat_id")%>&action=editartcat">编辑</a> 
            <a href="admin_art.asp?id=<%=rs("cat_id")%>&action=delartcat">删除</a> 
            <a href="art.asp?cat_id=<%=rs("cat_id")%>" target="_blank">查看</a></td>
        </tr>
        <%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr align="center"> 
          <td colspan="3" class=zoneloveds>当前没有文章分类！</td>
        </tr>
        <%end if%>
      </table>
      <%end if%>
	  <%if request.QueryString("action")="newartcat" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="admin_art.asp">
		<tr> 
            <td class=zonelovess>新校园内外分类</td>
        </tr>
        <tr> 
          <td class=zoneloveds>分类名称
              <input type="text" name="cat_name" size="40">
          </td>
        </tr>
        <tr> 
            <td class=zoneloveqs align="center" height="30">
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="reset" value="清空重写" class="button">
<input type="button" value="返回"  onClick="location.href='admin_art.asp'" class="button"></td>
        </tr>
		<input type="hidden" name="action" value="newartcat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.QueryString("action")="editartcat" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的id参数');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from a_cat where cat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="admin_art.asp">
          <tr> 
            <td class=zonelovess>修改文章分类</font></td>
          </tr>
          <tr> 
            <td class=zoneloveds>分类名称- 
              <input name="cat_name" type="text" class="textarea" id="cat_name" size="40" value="<%=rs("cat_name")%>"> 
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs align="center" height="30"> <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
<input type="button" value="返回"  onClick="location.href='admin_art.asp'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("cat_id")%>">
		  <input type="hidden" name="action" value="editartcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  	  if request.QueryString("action")="delartcat" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('非法的id参数');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from a_cat where cat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="admin_art.asp">
          <tr> 
            <td class=zonelovess>删除文章分类</font></td>
          </tr>
          <tr> 
            <td class=zoneloveds>分类名称- <%=rs("cat_name")%>
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs align="center" height="30"> <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
<input type="button" value="返回"  onClick="location.href='admin_art.asp'" class="button"></td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("cat_id")%>">
          <input type="hidden" name="action" value="delartcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
end if
if request.QueryString("action")="cat" then
dim totalart,Currentpage,totalpages,i,colname
sql="select * from art order by art_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr> 
          <td colspan="4" class=zonelovess>文章管理</td>
        </tr>
        <tr align="center"> 
          <td class=zoneloveqs width="10%">编号</td>
          <td class=zoneloveqs width="60%">文章标题</td>
          <td class=zoneloveqs width="10%">评论</td>
<td class=zoneloveqs width="20%">操作</td>
        </tr>
<%if not rs.eof then
rs.movefirst
rs.pagesize=articleperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalart=rs.recordcount
   if currentpage<>1 then
      if(currentpage-1)*articleperpage<totalart then
	     rs.move(currentpage-1)*articleperpage
		 dim bookmark
		 bookmark=rs.bookmark
	  end if
   end if
   if (totalart mod articleperpage)=0 then
      totalpages=totalart\articleperpage
   else
      totalpages=totalart\articleperpage+1
   end if
   i=0
do while not rs.eof and i<articleperpage%>
        <tr> 
          <td class=zoneloveds align="center"><%=rs("art_id")%>　</td>
          <td class=zoneloveds><a href="#"><%=rs("art_title")%></a></td>
<td class=zoneloveds align="center"><font color=red><%=rs("reviewcount")%></font></td>
          <td class=zoneloveds align="center"> <a href="admin_view.asp?action=art&art_id=<%=rs("art_id")%>">评论</a> <a href="admin_art.asp?art_id=<%=rs("art_id")%>&cat_id=<%=rs("cat_id")%>&action=editart">编辑</a> 
<a href="admin_art.asp?art_id=<%=rs("art_id")%>&cat_id=<%=rs("cat_id")%>&action=delart">删除</a> 
<a href="showart.asp?art_id=<%=rs("art_id")%>&cat_id=<%=rs("cat_id")%>" target="_blank">查看</a></td>
        </tr>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr align="center"> 
          <td class=zoneloveds colspan="4">当前没有文章！</td>
        </tr>
<%end if
end if%>
          <tr> 
            <td class=zoneloveqs colspan="4" align="right"><%=currentpage%> /<%=totalpages%>页,<%=totalart%>条记录/<%=articleperpage%>篇每页. 
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
              <a href="admin_art.asp?page=<%=i%>&action=cat"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_art.asp?page=<%=page%>&action=cat" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
<%
end if
if request.Form("action")="newart" then
sql="select * from art"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim poster,artfrom,title,keyword,content,artisbest,catid
catid=cint(request.form("cat_id"))
title=trim(replace(request.form("art_title"),"'",""))
keyword=trim(replace(request.form("art_keyword"),"'",""))
artisbest=request.form("isbest")
content=rtrim(replace(request.form("art_content"),"",""))
content=trim(replace(content," ","&nbsp;"))

if catid<1 then
  founderr=true
Response.Write("<script language=javascript>alert('你必须选择文章的分类！');history.back(1);</script>")
else
  rs("cat_id")=catid
end if
if title="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须填写文章的标题！');history.back(1);</script>")
else
  rs("art_title")=title
end if
if keyword="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须选择文章的类型！');history.back(1);</script>")
else
  rs("art_keyword")=keyword
end if
if content="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须填写文章的内容！');history.back(1);</script>")
else
  rs("art_content")=content
end if
if cint(artisbest)=1 then
  rs("isbest")=cint(artisbest)
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  sql="update allcount set articlecount = articlecount + 1"
  conn.execute(sql)
  
  response.redirect "admin_art.asp?action=cat"
end if
end if
if request.Form("action")="editart" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('非法校园内外id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from art where art_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
catid=cint(request.form("cat_id"))
title=trim(replace(request.form("art_title"),"'",""))
keyword=trim(replace(request.form("art_keyword"),"'",""))
artisbest=request.form("isbest")
content=rtrim(replace(request.form("art_content"),"",""))
content=trim(replace(content," ","&nbsp;"))
if catid<1 then
  founderr=true
Response.Write("<script language=javascript>alert('你必须选择文章的分类！');history.back(1);</script>")
else
  rs("cat_id")=catid
end if
if title="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须填写文章的标题！');history.back(1);</script>")
else
  rs("art_title")=title
end if
if keyword="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须选择文章的类型！');history.back(1);</script>")
else
  rs("art_keyword")=keyword
end if
if content="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须填写文章的内容！');history.back(1);</script>")
else
  rs("art_content")=content
end if
rs("isbest")=cint(artisbest)
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_art.asp?action=cat"
end if
end if
if request.Form("action")="delart" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('非法校园内外id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from art where art_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  sql="update allcount set articlecount = articlecount - 1"
  conn.execute(sql)
  
  response.redirect "admin_art.asp?action=cat"
end if
if request.QueryString("action")="newart" then%>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="addart" method="post" action="">
          <tr> 
            <td class=zonelovess>新增文章</td>
          </tr>
          <tr> 
            <td class=zoneloveds>标题
              <input type="text" name="art_title" size="60" class="textarea">
            </td>
          </tr>
          <tr> 
            <td class=zoneloveds>所属分类
              <select name="cat_id">
                <%
sql="select * from a_cat"
set rs=conn.execute(sql)
do while not rs.eof%>
              <option value="<%=rs("cat_id")%>"><%=rs("cat_name")%></option>
<%rs.movenext
loop
rs.close
set rs=nothing%>
              </select>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型 
<select name="art_keyword">
<option value="教师作品">教师作品</option>
<option value="学生作品">学生作品</option>
<option value="上级来文">上级来文</option>
<option value="校内新闻">校内新闻</option>
<option value="校内公告">校内公告</option>
<option value="临时通知">临时通知</option>
<option value="其它">其它</option>
</select>
            </td>
          </tr>
         
          <tr> 
            <td class=zoneloveds> 
              			  <textarea name="art_content" style="display:none"></textarea>
<iframe id="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=art_content&style=blue" frameborder="0" scrolling="No" width="550" height="320"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#E8E8E8" height="30" align="center"><input type="checkbox" name="isbest" value="1">推荐&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重写" class="button">
<input type="button" value="返回"  onClick="location.href='admin_art.asp?action=cat'" class="button"></td>
          </tr>
		  <input type="hidden" name="action" value="newart">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
	  if request.QueryString("action")="editart" then
if request.querystring("art_id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("art_id")) then
Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
if request.querystring("cat_id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("cat_id")) then
Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from art where art_id="&cint(request.querystring("art_id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
      <table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="addart" method="post" action="">
          <tr> 
            <td class=zonelovess>修改文章</td>
          </tr>
          <tr> 
            <td class=zoneloveds>标题
              <input name="art_title" type="text" class="textarea" id="art_title" size="60" value="<%=rs("art_title")%>"> 
            </td>
          </tr>
          <tr> 
            <td class=zoneloveds>所属分类 
              <select name="cat_id" id="cat_id">
                <%
sql="select * from a_cat"
set rs2=conn.execute(sql)
do while not rs2.eof%>

                <option value="<%=rs2("cat_id")%>"<%if cint(rs2("cat_id"))=cint(request("cat_id")) then response.Write " selected"%>><%=rs2("cat_name")%></option>



                <%rs2.movenext
loop
rs2.close
set rs2=nothing%>
              </select> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型 
<select name="art_keyword">
<option value="教师作品"<%if rs("art_keyword")="教师作品" then response.write "selected"%>>教师作品</option>
<option value="学生作品"<%if rs("art_keyword")="学生作品" then response.Write " selected"%>>学生作品</option>
<option value="上级来文"<%if rs("art_keyword")="上级来文" then response.Write " selected"%>>上级来文</option>
<option value="校内新闻"<%if rs("art_keyword")="校内新闻" then response.Write " selected"%>>校内新闻</option>
<option value="校内公告"<%if rs("art_keyword")="校内公告" then response.Write " selected"%>>校友公告</option>
<option value="临时通知"<%if rs("art_keyword")="临时通知" then response.Write " selected"%>>临时通知</option>
<option value="其它"<%if rs("art_keyword")="其它" then response.Write " selected"%>>其它</option>
</select></td>
          </tr>
        
          <tr> 
            <td class=zoneloveds> <textarea name="art_content" style="display:none"><%=Server.HTMLEncode(rs("art_content"))%></textarea>
<iframe id="eWebEditor2" src="eWebEditor/ewebeditor.htm?id=art_content&style=blue" frameborder="0" scrolling="No" width="550" height="320"></iframe>
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs height="30" align="center"><input type="checkbox" name="isbest" value="1" id="isbest" <%if rs("isbest")=1 then response.write "checked" end if%>><input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
<input type="button" value="返回"  onClick="location.href='admin_art.asp?action=cat'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("art_id")%>">
		  <input type="hidden" name="action" value="editart">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  	  if request.QueryString("action")="delart" then
if request.querystring("art_id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("art_id")) then
Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
if request.querystring("cat_id")="" then
Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("cat_id")) then
Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from art where art_id="&cint(request.querystring("art_id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
      <table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="addart" method="post" action="">
          <tr> 
            <td class=zonelovess>删除文章</td>
          </tr>
          <tr> 
            <td class=zoneloveds>标题- <%=rs("art_title")%> 
            </td>
          </tr>
          <tr> 
            <td class=zoneloveds>所属分类 
                <%
sql="select * from a_cat where cat_id="&cint(request.querystring("cat_id"))
set rs2=conn.execute(sql)
%>
<%=rs2("cat_name")%>
<%
rs2.close
set rs2=nothing%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型 <%=rs("art_keyword")%>
            </td>
          </tr>
          <tr> 
            <td class=zoneloveds><%=ubb2html(formatStr(autourl(rs("art_content"))), true, true)%>
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs height="30" align="center"> <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
<input type="button" value="返回"  onClick="location.href='admin_art.asp?action=cat'" class="button"></td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("art_id")%>">
          <input type="hidden" name="action" value="delart">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%rs.close
set rs=nothing

end if%>