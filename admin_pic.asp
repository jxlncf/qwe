<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newpiccat" then
sql="select * from piccat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim piccatname
piccatname=trim(replace(request.form("piccat_name"),"'",""))
if piccatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("piccat_name")=piccatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_pic.asp?action=piccat"
end if
elseif request.form("action")="newpic" then
sql="select * from pic"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim piccatid,picname,picurl,picpic,picspic,picdesc,picisbest
piccatid=cint(request.form("piccatid"))
picname=trim(replace(request.form("name"),"'",""))
picurl=trim(replace(request.form("url"),"'",""))
picpic=trim(replace(request.form("pic"),"'",""))
picspic=trim(replace(request.form("spic"),"'",""))
picdesc=trim(replace(request.form("desc"),"'",""))
picisbest=request.form("isbest")
if picname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片标题！');history.back(1);</script>")
else
  rs("pic_name")=picname
end if
if piccatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择图片所属分类！');history.back(1);</script>")
else
  rs("piccat_id")=piccatid
end if
if picurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写文件大小！');history.back(1);</script>")
else
  rs("pic_url")=picurl
end if
if picspic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写缩图地址！');history.back(1);</script>")
else
  rs("pic_spic")=picspic
end if
if picpic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片地址！');history.back(1);</script>")
else
  rs("pic_pic")=picpic
end if
if picdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片的简要说明！');history.back(1);</script>")
else
  rs("pic_desc")=picdesc
end if
if cint(picisbest)=1 then
  rs("isbest")=cint(picisbest)
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  sql="update allcount set piccount = piccount + 1"
  conn.execute(sql)
  
  response.redirect "admin_pic.asp?action=pic"
end if
elseif request.form("action")="editpic" then
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
sql="select * from pic where pic_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
piccatid=cint(request.form("piccatid"))
picname=trim(replace(request.form("name"),"'",""))
picurl=trim(replace(request.form("url"),"'",""))
picpic=trim(replace(request.form("pic"),"'",""))
picspic=trim(replace(request.form("spic"),"'",""))
picdesc=trim(replace(request.form("desc"),"'",""))
picisbest=request.form("isbest")
if picname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片标题！');history.back(1);</script>")
else
  rs("pic_name")=picname
end if
if piccatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择图片所属分类！');history.back(1);</script>")
else
  rs("piccat_id")=piccatid
end if
if picurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片大小！');history.back(1);</script>")
else
  rs("pic_url")=picurl
end if
if picspic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写缩图地址！');history.back(1);</script>")
else
  rs("pic_spic")=picspic
end if
if picpic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片地址！');history.back(1);</script>")
else
  rs("pic_pic")=picpic
end if
if picdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写图片的简要说明！');history.back(1);</script>")
else
  rs("pic_desc")=picdesc
end if
rs("isbest")=cint(picisbest)

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_pic.asp?action=pic"
end if
elseif request.form("action")="delpic" then
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
sql="select * from pic where pic_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  sql="update allcount set piccount = piccount - 1"
  conn.execute(sql)
  
  response.redirect "admin_pic.asp?action=pic"
elseif request.form("action")="editpiccat" then
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
sql="select * from piccat where piccat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
piccatname=trim(replace(request.form("piccat_name"),"'",""))
if piccatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("piccat_name")=piccatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_pic.asp?action=piccat"
end if
elseif request.form("action")="delpiccat" then
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
sql="select * from piccat where piccat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_pic.asp?action=piccat"
end if

end if%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
     <%if request.querystring("action")="piccat" then
sql="select * from piccat order by piccat_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr class=zonelovess> 
          <td colspan="3">
          图片分类管理</td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="10%">编号</td>
          <td width="70%">分类名称</td>
          <td width="20%">操作</td>
        </tr>
        <%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("piccat_id")%>　</td>
          <td><a href="#"><%=rs("piccat_name")%></a>　</td>

          <td align="center"><a href="admin_pic.asp?id=<%=rs("piccat_id")%>&action=editpiccat">编辑</a> 
            <a href="admin_pic.asp?id=<%=rs("piccat_id")%>&action=delpiccat">删除</a> <a href="pic.asp?piccat_id=<%=rs("piccat_id")%>" target="_blank">查看</a></td>
        </tr>
        <%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr class=zoneloveds align="center"> 
          <td colspan="3">当前没有分类！</td>
        </tr>
        <%rs.close
set rs=nothing
		end if%>
      </table>
     <%end if
	 if request.querystring("action")="newpiccat" then%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
		<tr class=zonelovess> 
            <td>
            新增分类</td>
        </tr>
        <tr class=zoneloveds> 
            <td>分类名称- 
              <input type="text" name="piccat_name" size="40" class="textarea">
          </td>
        </tr>
        <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
        </tr>
		<input type="hidden" name="action" value="newpiccat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.QueryString("action")="editpiccat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from piccat where piccat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>
            修改分类</td>
          </tr>
          <tr class=zoneloveds> 
            <td>分类名称- 
              <input name="piccat_name" type="text" class="textarea" id="piccat_name" size="40" value="<%=rs("piccat_name")%>"> 
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <tdalign="center"> <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
              <input name="Reset" type="reset" class="button" id="Reset" value="清空重填"> </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("piccat_id")%>">
          <input type="hidden" name="action" value="editpiccat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="delpiccat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from piccat where piccat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess><td>删除分类</td></tr>
          <tr class=zoneloveds><td>分类名称- <%=rs("piccat_name")%></td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("piccat_id")%>">
          <input type="hidden" name="action" value="delpiccat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="pic" then
sql="select * from pic order by pic_id desc"
if request.querystring("piccat_id")<>"" then
sql="select * from pic where piccat_id="&cint(request.querystring("piccat_id"))&" order by pic_id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
totalpic=rs.recordcount
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
	    <form name="form3" method="post" action="">
        <tr class=zonelovess><td align="right" width="20%" colspan="4" style="padding-top:2px;">
              <select style="margin:-3px" name="go" onChange='window.location=form.go.options[form.go.selectedIndex].value'>
			  <option value="">选择显示方式</option>
			  <option value="admin_pic.asp?action=pic">显示所有图片</option>
<%sql="select * from piccat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_pic.asp?action=pic&piccat_id=<%=rs2("piccat_id")%>"><%=rs2("piccat_name")%></option>
<%rs2.movenext
loop
if rs2.bof and rs2.eof then%><option value="">当前没有分类</option>
<%end if
rs2.close
set rs2=nothing%>
              </select>
            </td>
        </tr>
		</form>
        <tr class=zoneloveqs align="center"> 
<td  width="80">图片名称</td><td  width="*">图片名称</td><td width="40">评论</td><td  width="100">管理</td>
        </tr>
<%
if not rs.eof then
rs.movefirst
rs.pagesize=picperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalpic=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*picperpage<totalpic then
	       rs.move(currentpage-1)*picperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalpic mod picperpage)=0 then
      totalpages=totalpic\picperpage
   else
      totalpages=totalpic\picperpage+1
   end if
i=0
do while not rs.eof and i<picperpage
%>
        <tr class=zoneloveds> 
         <td><img src='<%=rs("pic_spic")%>'  width='75' border='0'></td><td>·<%=rs("pic_name")%> </td><td align="center"><font color=red><%=rs("reviewcount")%></font> </td><td align="center">
            <a href="admin_view.asp?action=pic&pic_id=<%=rs("pic_id")%>">评论</a>  <a href="admin_pic.asp?id=<%=rs("pic_id")%>&action=editpic">编辑</a> 
            <a href="admin_pic.asp?id=<%=rs("pic_id")%>&action=delpic">删除</a></td>
        </tr>
      <%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr class=zoneloveds> 
          <td colspan="4" height="22" align="center">当前没有图片！</td>
        </tr>
      <%end if
end if%>
        <form name="form1" method="post" action="admin_pic.asp?action=pic&piccat_id=<%=request.querystring("piccat_id")%>">
          <tr class=zoneloveqs> 
            <td align="right" colspan="4"> <%=currentpage%> /<%=totalpages%>页,<%=totalpic%>条记录/<%=picperpage%>篇每页. 
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
              <a href="admin_pic.asp?action=pic&page=<%=i%>&piccat_id=<%=request.querystring("piccat_id")%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_pic.asp?action=pic&page=<%=page%>&piccat_id=<%=request.querystring("piccat_id")%>" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
      <br>
	  <%end if
	  if request.querystring("action")="newpic" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            新增图片</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片名称</td>
            <td width="83%"> 
              <input type="text" name="name" class="textarea" size="40">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">所属分类</td>
            <td width="83%"> 
              <select name="piccatid" class="textarea">
<%
sql="select * from piccat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof%><option value="<%=rs("piccat_id")%>"><%=rs("piccat_name")%></option>
<%rs.movenext
loop
if rs.eof and rs.bof then%><option value="0">当前没有分类</option>
<%end if%>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片大小</td>
            <td width="83%"> 
              <input type="text" name="url" class="textarea" size="60">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">缩图地址</td>
            <td width="83%"> 
              <input type="text" name="spic" class="textarea" size="60">
            </td>
          </tr>
<tr class=zoneloveds> 
<td>上传缩图：</td>
<td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=spic"></iframe></td>
</tr>
          <tr class=zoneloveds> 
            <td width="17%">图片地址</td>
            <td width="83%"> 
              <input type="text" name="pic" class="textarea" size="60">
            </td>
          </tr>
<tr class=zoneloveds> 
<td>上传原图：</td>
<td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=pic"></iframe></td>
</tr>
          <tr class=zoneloveds> 
            <td width="17%">图片介绍</td>
            <td width="83%"> 
              <textarea name="desc" class="textarea" cols="65" rows="6">图片版权归图片本站所有,未经同意不得转载!</textarea>
            </td>
          </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="checkbox" name="isbest" value="1" class="textarea">
              推荐&nbsp;&nbsp;&nbsp; 
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_pic.asp?action=pic">返回</a>] </td>
          </tr>
		  <input type="hidden" name="action" value="newpic">
		<input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
	  if request.QueryString("action")="editpic" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from pic where pic_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            修改图片</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片名称</td>
            <td width="83%"> 
              <input name="name" type="text" class="textarea" id="name" size="40" value="<%=rs("pic_name")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">所属分类</td>
            <td width="83%"> 
              <select name="piccatid" class="textarea" id="piccatid">
                <%
sql="select * from piccat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
                <option value="<%=rs2("piccat_id")%>" <%if rs2("piccat_id")=rs("piccat_id") then response.write "selected" end if%>><%=rs2("piccat_name")%></option>
                <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
                <option value="0">当前没有分类</option>
                <%end if%>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片大小</td>
            <td width="83%"> 
              <input name="url" type="text" class="textarea" id="url" size="60" value="<%=rs("pic_url")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">缩图地址</td>
            <td width="83%"> 
              <input name="spic" type="text" class="textarea" id="spic" size="60" value="<%=rs("pic_spic")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
<td>上传缩图：</td>
<td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=spic"></iframe></td>
</tr>
          <tr class=zoneloveds> 
            <td width="17%">图片地址</td>
            <td width="83%"> 
              <input name="pic" type="text" class="textarea" id="pic" size="60" value="<%=rs("pic_pic")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
<td>上传原图：</td>
<td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=pic"></iframe></td>
</tr>
          <tr class=zoneloveds> 
            <td width="17%">图片介绍</td>
            <td width="83%"> 
              <textarea name="desc" cols="65" rows="6" class="textarea" id="desc"><%=rs("pic_desc")%></textarea>
            </td>
          </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              推荐
              <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_pic.asp?action=pic">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("pic_id")%>">
          <input type="hidden" name="action" value="editpic">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="delpic" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from pic where pic_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除图片</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片名称</td>
            <td width="83%"> <%=rs("pic_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">所属分类</td>
            <td width="83%"> 
                <%
sql="select * from piccat where piccat_id="&rs("piccat_id")
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
%><%=rs2("piccat_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片大小</td>
            <td width="83%"><%=rs("pic_url")%>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">缩图地址</td>
            <td width="83%"><%=rs("pic_spic")%>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片地址</td>
            <td width="83%"><%=rs("pic_pic")%>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">图片介绍</td>
            <td width="83%"><%=rs("pic_desc")%> 　</td>
          </tr>
          <tr  class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              推荐 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除">
              [<a href="admin_pic.asp?action=pic">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("pic_id")%>">
          <input type="hidden" name="action" value="delpic">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
end if
%>