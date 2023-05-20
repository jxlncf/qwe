<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<title><%=webname%>-链接管理</title>
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newflcat" then
sql="select * from flcat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim flcatname
flcatname=trim(replace(request.form("flcat_name"),"'",""))
if flcatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("flcat_name")=flcatname
end if
if request.form("isimage")=1 then
  rs("isimage")=1
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp"
end if
end if
if request.form("action")="editflcat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法校园内外分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from flcat where flcat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
flcatname=trim(replace(request.form("flcat_name"),"'",""))
if flcatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须填写分类名称！');history.back(1);</script>")
else
  rs("flcat_name")=flcatname
end if
if request.form("isimage")<>"" then
  rs("isimage")=1
else
  rs("isimage")=0
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp"
end if
end if
if request.form("action")="delflcat" then
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
sql="select * from flcat where flcat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  sql="update allcount set friendlinkcount = friendlinkcount - 1"
  conn.execute(sql)
  response.redirect "admin_link.asp"
end if
if request.form("action")="editfl" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法校园内外分类id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
flcatid=cint(request.form("flcat_id"))
flname=trim(replace(request.form("fl_name"),"'",""))
flurl=trim(replace(request.form("fl_url"),"'",""))
fllogo=trim(replace(request.form("fl_logo"),"'",""))
if flcatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('非法的分类参数！');history.back(1);</script>")
else
  sql="select isimage from flcat where flcat_id="&cint(request.form("flcatid"))
   set rs2=server.createobject("adodb.recordset")
   rs2.open sql,conn,1,1
   if rs2("isimage")=1 then
     isimage=false
   end if
   rs2.close
   set rs2=nothing
   rs("flcat_id")=flcatid
end if
if flname="" then
   founderr=true
   Response.Write("<script language=javascript>alert('站点名称未填写！');history.back(1);</script>")
else
  if strLength(flname)>50 then
      founderr=true
	  Response.Write("<script language=javascript>alert('站点名称太长，不可以超过50个字符！');history.back(1);</script>")
  else
      rs("fl_name")=flname
  end if
end if
if flurl="" then
   founderr=true
   Response.Write("<script language=javascript>alert('站点地址未填写！');history.back(1);</script>")
else
  if strLength(flurl)>100 then
      founderr=true
	  Response.Write("<script language=javascript>alert('站点地址太长，不可以超过100个字符！');history.back(1);</script>")
  else
      rs("fl_url")=flurl
  end if
end if
rs("fl_logo")=fllogo
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=link"
end if
end if
if request.form("action")="pass" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
flcatid=cint(request.form("flcat_id"))
flname=trim(replace(request.form("fl_name"),"'",""))
flurl=trim(replace(request.form("fl_url"),"'",""))
fllogo=trim(replace(request.form("fl_logo"),"'",""))
if flcatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('非法的分类参数！');history.back(1);</script>")
else
  sql="select isimage from flcat where flcat_id="&cint(request.form("flcatid"))
   set rs2=server.createobject("adodb.recordset")
   rs2.open sql,conn,1,1
   if rs2("isimage")=1 then
     isimage=false
   end if
   rs2.close
   set rs2=nothing
   rs("flcat_id")=flcatid
end if
if flname="" then
   founderr=true
   Response.Write("<script language=javascript>alert('站点名称未填写！');history.back(1);</script>")
else
  if strLength(flname)>50 then
      founderr=true
	  Response.Write("<script language=javascript>alert('站点名称太长，不可以超过50个字符！');history.back(1);</script>")
  else
      rs("fl_name")=flname
  end if
end if
if flurl="" then
   founderr=true
   Response.Write("<script language=javascript>alert('站点地址未填写！');history.back(1);</script>")
else
  if strLength(flurl)>100 then
      founderr=true
	  Response.Write("<script language=javascript>alert('站点地址太长，不可以超过100个字符！');history.back(1);</script>")
  else
      rs("fl_url")=flurl
  end if
end if
rs("fl_logo")=fllogo
if founderr then
  call diserror()
  response.end
else
  rs("passed")=0
  rs.update
  conn.execute(sql)
  sql="update allcount set friendlinkcount = friendlinkcount + 1"
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=check"
end if
end if
if request.form("action")="lkpass" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs("lk")=0
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=lk"
end if
if request.form("action")="lklk" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs("lk")=1
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=link"
end if
if request.form("action")="delfl" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=link"
end if
if request.form("action")="delnopass" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=check"
end if
if request.form("action")="lkdel" then
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
sql="select * from friendlink where fl_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_link.asp?action=lk"
end if
end if
sql="select * from flcat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<%dim passing
		  sql="select fl_id from friendlink where passed=1"
		  set rs2=server.createobject("adodb.recordset")
		  rs2.open sql,conn,1,1
		  passing=rs2.recordcount
		  rs2.close
		  set rs2=nothing%>

<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zoneloveqs> 
		  <td width="33%"><a href="admin_link.asp?action=check">链接审核(待审核数:<span class="newshead"><%=passing%></span>)</a></td>
        </tr>
      </table>
      <table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="6"></td>
        </tr>
      </table>
	  <%if request.querystring("action")="" then%>
     <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr class=zonelovess> 
          <td colspan="3">
          链接分类管理</font></td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="10%">编号</td>
          <td width="70%">分类名称</td>
          <td width="20%">操作</td>
        </tr>
<%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("flcat_id")%>　</td>
          <td><%=rs("flcat_name")%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#ff6600"><%if rs("isimage")=1 then response.write "文本链接" end if%></font></td>
          <td align="center"><a href="admin_link.asp?id=<%=rs("flcat_id")%>&action=editflcat">编辑</a> 
            <a href="admin_link.asp?id=<%=rs("flcat_id")%>&action=delflcat">删除</a></td>
        </tr>
<%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr class=zoneloveds align="center"> 
          <td colspan="4">当前没有链接分类！</td>
        </tr>
<%end if%>
      </table>
      <br>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
		<tr class=zonelovess> 
          <td>
          新增链接分类</font></td>
        </tr>
        <tr class=zoneloveds> 
            <td>分类名称-
              <input type="text" name="flcat_name" size="40" class="textarea">
              &nbsp;&nbsp;&nbsp;
              <input type="checkbox" name="isimage" value="1" class="textarea">
              文本链接</td>
        </tr>
        <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
        </tr>
		<input type="hidden" name="action" value="newflcat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.querystring("action")="editflcat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的链接分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from flcat where flcat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
	  <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>
            修改链接分类</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>分类名称- 
              <input type="text" name="flcat_name" size="40" class="textarea" value="<%=rs("flcat_name")%>">
			  &nbsp;&nbsp;&nbsp;
              <input type="checkbox" name="isimage" value="1" class="textarea" <%if rs("isimage")=1 then response.write "checked" end if%>>
              文本链接 </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("flcat_id")%>">
          <input type="hidden" name="action" value="editflcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="delflcat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的链接分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from flcat where flcat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>
            删除链接分类</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>分类名称- <%=rs("flcat_name")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("flcat_id")%>">
          <input type="hidden" name="action" value="delflcat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="link" then
sql="select * from friendlink where passed=0 and lk=0 order by fl_id deSC"
if request.querystring("flcat_id")<>"" then
sql="select * from friendlink where flcat_id="&cint(request.querystring("flcat_id"))&" and passed=0 and lk=0 order by fl_id deSC"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
     <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
	    <form name="form3" method="post" action="">
        <tr class=zonelovess><td colspan="4">链接管理</td>
          <td align="right" style="padding-top:2px;">
            <select name="go" style="margin:-3px" onChange='window.location=form.go.options[form.go.selectedIndex].value'>
			  <option value="">选择显示方式</option>
			  <option value="admin_link.asp?action=link">显示所有链接</option>
<%sql="select * from flcat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_link.asp?action=link&flcat_id=<%=rs2("flcat_id")%>"><%=rs2("flcat_name")%></option>
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
          <td width="50">编号</td>
          <td width="100">LOGO</td>
          <td width="*">链接名称</td>
          <td width="100">分类</td>
          <td width="120">操作</td>
        </tr>
<%if not rs.eof then
rs.movefirst
rs.pagesize=adflperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalfl=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*adflperpage<totalfl then
	       rs.move(currentpage-1)*adflperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalfl mod adflperpage)=0 then
      totalpages=totalfl\adflperpage
   else
      totalpages=totalfl\adflperpage+1
   end if
i=0
do while not rs.eof and i<adflperpage
sql="select flcat_name from flcat where flcat_id="&rs("flcat_id")
set rs2=conn.execute(sql)%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("fl_id")%></td><td align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"' width='88' height='31' border='0'>" else response.write "<font color='#cccccc'>无</font>" end if%></td>
          <td><a href="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td align="center"><%=rs2("flcat_name")%></td>
          <td align="center"><a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lklk">黑名单</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=editfl">编辑</a> 
            <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=delfl">删除</a></td>
        </tr>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr class=zoneloveds> 
          <td colspan="5" align="center">当前还没有链接！</td>
        </tr>
<%end if
end if%>
        <form name="form1" method="post" action="admin_link.asp?action=link&flcat_id=<%=request.querystring("flcat_id")%>">
          <tr class=zoneloveqs> 
            <td align="right" colspan="5"> <%=currentpage%> /<%=totalpages%>页,<%=totalfl%>条记录/<%=adflperpage%>篇每页. 
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
              <a href="admin_link.asp?action=link&page=<%=i%>&flcat_id=<%=request.querystring("flcat_id")%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_link.asp?action=link&page=<%=page%>&flcat_id=<%=request.querystring("flcat_id")%>" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="editfl" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
	  <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            修改链接</font></td>
          </tr>
          <tr class=zoneloveds><td align="right" width="30%">选择类别 </td><td width="70%">
              <select name="flcat_id" id="flcat_id">
<%sql="select * from flcat"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%><option value="<%=rs("flcat_id")%>" <%if rs("flcat_id")=rs2("flcat_id") then response.write "selected" end if%>><%=rs("flcat_name")%></option><%rs.movenext
loop%></select></tr>
<tr class=zoneloveds><td align="right" width="30%">站点名称 </td><td width="70%"><input type="text" name="fl_name" size="40" class="textarea" value="<%=rs2("fl_name")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址 </td><td width="70%">
              <input type="text" name="fl_url" size="40" class="textarea" value="<%=rs2("fl_url")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址 </td><td width="70%">
              <input type="text" name="fl_logo" size="40" class="textarea" value="<%=rs2("fl_logo")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td align="right">上传LOGO： </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=link"></iframe></td>
            </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
		  <input type="hidden" name="flcatid" value="<%=rs2("flcat_id")%>">
		  <input type="hidden" name="action" value="editfl">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.querystring("action")="pass" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
	  <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            链接通过</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">选择类别 </td><td width="70%">
              <select name="flcat_id" id="flcat_id">
                <%sql="select * from flcat"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%>
                <option value="<%=rs("flcat_id")%>" <%if rs("flcat_id")=rs2("flcat_id") then response.write "selected" end if%>><%=rs("flcat_name")%></option>
                <%rs.movenext
loop%>
              </select>
<tr class=zoneloveds><td align="right" width="30%">站点名称</td><td width="70%"><input type="text" name="fl_name" size="40" class="textarea" value="<%=rs2("fl_name")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址</td><td width="70%"><input type="text" name="fl_url" size="40" class="textarea" value="<%=rs2("fl_url")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址</td><td width="70%"><input type="text" name="fl_logo" size="40" class="textarea" value="<%=rs2("fl_logo")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td align="right">上传LOGO： </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=link"></iframe></td>
            </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定通过" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
		  <input type="hidden" name="flcatid" value="<%=rs2("flcat_id")%>">
		  <input type="hidden" name="action" value="pass">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.querystring("action")="lkpass" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
	  <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            取消黑名单</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">所属类别 </td><td width="70%">
<%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">站点名称</td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址</td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址</td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定取消" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
		  <input type="hidden" name="flcatid" value="<%=rs2("flcat_id")%>">
		  <input type="hidden" name="action" value="lkpass">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.querystring("action")="lklk" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
	  <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            拉入黑名单</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">所属类别 </td><td width="70%">
<%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">站点名称</td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址</td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址</td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定拉入" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
		  <input type="hidden" name="flcatid" value="<%=rs2("flcat_id")%>">
		  <input type="hidden" name="action" value="lklk">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="delfl" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除链接</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">所属类别 </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">站点名称: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址: </td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
          <input type="hidden" name="action" value="delfl">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%
end if
	  if request.querystring("action")="lkdel" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除黑名单链接</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">所属类别 </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">站点名称: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址: </td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
              <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
          <input type="hidden" name="action" value="lkdel">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%
end if
if request.querystring("action")="delnopass" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的分类ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from friendlink where fl_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除链接</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">所属类别 </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">站点名称: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">站点地址: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo地址: </td><td width="70%"><%=rs2("fl_logo")%></td></tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
 <input type="button" value="返回"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs2("fl_id")%>">
          <input type="hidden" name="action" value="delnopass">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="check" then
sql="select * from friendlink where passed=1"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="checklink.asp">
          <tr class=zonelovess> 
            <td colspan="5">
            链接审核</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
          <td width="50">编号</td>
          <td width="100">LOGO</td>
          <td width="*">链接名称</td>
          <td width="100">分类</td>
          <td width="100">操作</td>
          </tr>
          <%for i=1 to rs.recordcount
sql="select flcat_name from flcat where flcat_id="&rs("flcat_id")
set rs2=conn.execute(sql)%>
          <tr class=zoneloveds> 
            <td align="center"><%=rs("fl_id")%></td><td width="100" align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"'>" else response.write "<font color='#cccccc'>无</font>" end if%></td>
            <td><a href="<%=rs("fl_url")%>" title="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td width="100" align="center"><%=rs2("flcat_name")%></td>
            <td align="center"> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=pass">通过</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=delnopass">删除</a></td>
          </tr>
          <%rs.movenext
next
if rs.bof and rs.eof then%>
          <tr class=zoneloveds align="center"> 
            <td colspan="5">当前没有待审核的链接！</td>
          </tr>
          <%end if%>
          <input type="hidden" name="action" value="check">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if

if request.querystring("action")="lk" then
sql="select * from friendlink where lk=1"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
          <tr class=zonelovess> 
            <td colspan="5">
            黑名单管理</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
          <td width="50">编号</td>
          <td width="100">LOGO</td>
          <td width="*">链接名称</td>
          <td width="100">分类</td>
          <td width="100">操作</td>
          </tr>
          <%for i=1 to rs.recordcount
sql="select flcat_name from flcat where flcat_id="&rs("flcat_id")
set rs2=conn.execute(sql)%>
          <tr class=zoneloveds> 
            <td align="center"><%=rs("fl_id")%></td><td width="100" align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"'>" else response.write "<font color='#cccccc'>无</font>" end if%></td>
            <td><a href="<%=rs("fl_url")%>" title="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td width="100" align="center"><%=rs2("flcat_name")%></td>
            <td align="center"> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lkpass">取消</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lkdel">删除</a></td>
          </tr>
          <%rs.movenext
next
if rs.bof and rs.eof then%>
          <tr class=zoneloveds align="center"> 
            <td colspan="5">当前没有黑名单链接！</td>
          </tr>
          <%end if%>
          <input type="hidden" name="action" value="check">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
rs.close
set rs=nothing
end if
%>