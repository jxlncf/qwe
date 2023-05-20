<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
 if request.form("action")="newvote" then
  dim vtname,startdate,expiredate
  vtname=trim(replace(request.form("vt_name"),"'",""))
  startdate=request.form("vt_startdate")
  expiredate=request.form("vt_expiredate")
  if vtname="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须填写投票主题！');history.back(1);</script>")
  end if
  if startdate="" then
    startdate=formatdatetime(now(),2)
  else
    startdate=formatdatetime(startdate,2)
  end if
  if expiredate="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须填写该投票主题的过期日期！');history.back(1);</script>")
  else
    expiredate=formatdatetime(expiredate,2)
  end if
  
  if founderr then
    call diserror()
	response.end
  else
	sql="select * from votetopic"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	rs("vt_name")=vtname
	rs("vt_startdate")=startdate
	rs("vt_expiredate")=expiredate
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_vote.asp?action=vote"
  end if
 end if
 if request.form("action")="editvote" then
  dim vtid
  vtname=trim(replace(request.form("vt_name"),"'",""))
  startdate=request.form("vt_startdate")
  expiredate=request.form("vt_expiredate")
  if request.form("id")="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  else
    vtid=cint(request.form("id"))
  end if
  if vtname="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须填写投票主题！');history.back(1);</script>")
  end if
  if startdate="" then
    startdate=formatdatetime(now(),2)
  else
    startdate=formatdatetime(startdate,2)
  end if
  if expiredate="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须填写该投票主题的过期日期！');history.back(1);</script>")
  else
    expiredate=formatdatetime(expiredate,2)
  end if
  
  if founderr then
    call diserror()
	response.end
  else
	sql="select * from votetopic where vt_id="&vtid
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs("vt_name")=vtname
	rs("vt_startdate")=startdate
	rs("vt_expiredate")=expiredate
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_vote.asp?action=vote"
  end if
 end if
 if request.form("action")="delvote" then
  if request.form("id")="" then
    founderr=true
	Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  else
    vtid=cint(request.form("id"))
  end if
  
  if founderr then
    call diserror()
	response.end
  else
	sql="select * from votetopic where vt_id="&vtid
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.delete
	rs.close
	set rs=nothing
	response.redirect "admin_vote.asp?action=vote"
  end if
 end if
 if request.form("action")="newitem" then
   dim itemname
   vtid=cint(request.form("id"))
   itemname=trim(replace(request.form("item_name"),"'",""))
   if vtid<1 then
     founderr=true
	 Response.Write("<script language=javascript>alert('非法的对象参数！');history.back(1);</script>")
   end if
   if itemname="" then
     founderr=true
	 Response.Write("<script language=javascript>alert('你必须填写选项名称！');history.back(1);</script>")
   end if
   
   if founderr then
     call diserror()
	 response.end
   else
	 sql="select * from voteitem"
	 set rs=server.createobject("adodb.recordset")
	 rs.open sql,conn,1,3
	 rs.addnew
	 rs("item_name")=itemname
	 rs("vt_id")=vtid
	 rs.update
	 rs.close
	 set rs=nothing
	 response.redirect "admin_vote.asp?action=vote&caction=edit&id="&vtid
   end if
 end if
 if request.form("action")="edititem" then
   dim itemid
   itemid=cint(request.form("id"))
   itemname=trim(replace(request.form("item_name"),"'",""))
   if itemid<1 then
     founderr=true
	 Response.Write("<script language=javascript>alert('非法的对象参数！');history.back(1);</script>")
   end if
   if itemname="" then
     founderr=true
	 Response.Write("<script language=javascript>alert('你必须填写选项名称！');history.back(1);</script>")
   end if
   
   if founderr then
     call diserror()
	 response.end
   else
	 sql="select * from voteitem where item_id="&itemid
	 set rs=server.createobject("adodb.recordset")
	 rs.open sql,conn,1,3
	 rs("item_name")=itemname
	 rs.update
	 rs.close
	 set rs=nothing
	 response.redirect "admin_vote.asp?action=vote&caction=edit&id="&request.form("vtid")
   end if
 end if
 if request.form("action")="delitem" then
   itemid=cint(request.form("id"))
   if itemid<1 then
     founderr=true
	 Response.Write("<script language=javascript>alert('非法的对象参数！');history.back(1);</script>")
   end if

   if founderr then
     call diserror()
	 response.end
   else
	 sql="select * from voteitem where item_id="&itemid
	 set rs=server.createobject("adodb.recordset")
	 rs.open sql,conn,1,3
	 rs("item_name")=itemname
	 rs.delete
	 rs.close
	 set rs=nothing
	 response.redirect "admin_vote.asp?action=vote&caction=edit&id="&request.form("vtid")
   end if
 end if
end if
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<script language="javascript" src="inc/PopupCalendar.js" ></script>
<script >
var oCalendarEn=new PopupCalendar("oCalendarEn");
oCalendarEn.Init();
var oCalendarChs=new PopupCalendar("oCalendarChs");
oCalendarChs.weekDaySting=new Array("日","一","二","三","四","五","六");
oCalendarChs.monthSting=new Array("一月","二月","三月","四月","五月","六月","七月","八月","九月","十月","十一月","十二月");
oCalendarChs.oBtnTodayTitle="今天";
oCalendarChs.oBtnCancelTitle="取消";
oCalendarChs.Init();
</script>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<%if request.querystring("action")="newvote" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            新增投票主题</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="16%">主题名称</td>
            <td width="84%"> 
              <input type="text" name="vt_name" size="50" maxlength="50" class="textarea">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="16%">开始日期</td>
            <td width="84%"> 
              <input readonly type="text" name="vt_startdate" id="aa" onclick="getDateString(this,oCalendarChs)" value=""></td>
          </tr>
          <tr class=zoneloveds> 
            <td width="16%">过期日期</td>
            <td width="84%"> 
              <input readonly type="text" name="vt_expiredate" id="aa" onclick="getDateString(this,oCalendarChs)" value=""></td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2" height="30" bgcolor="#F5F5F5"> 
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
          <input type="hidden" name="action" value="newvote">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="vote" then
sql="select * from votetopic"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr class=zonelovess> 
          <td colspan="3">
          投票主题管理</td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="10%">编号</td>
          <td width="60%">投票主题</td>
          <td width="30%">操作</td>
        </tr>
<%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("vt_id")%>　</td>
          <td><%=rs("vt_name")%>　</td>
          <td align="center"><a href="admin_vote.asp?action=vote&caction=edit&id=<%=rs("vt_id")%>">编辑</a> 
            <a href="admin_vote.asp?action=vote&caction=del&id=<%=rs("vt_id")%>">删除</a></td>
        </tr>
<%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr class=zoneloveds> 
          <td colspan="3" align="center">当前没有投票主题！</td>
        </tr>
<%end if%>
      </table><br>
<%if request.querystring("action")="vote" and request.querystring("caction")="edit" then
if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  response.end
end if
sql="select * from votetopic where vt_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="3">
            修改投票主题</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">主题名称</td>
            <td colspan="2"> 
              <input type="text" name="vt_name" size="50" maxlength="50" class="textarea" value="<%=rs("vt_name")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">开始日期</td>
            <td colspan="2"> 
              <input readonly type="text" name="vt_startdate" id="aa" onclick="getDateString(this,oCalendarChs)" value="<%=rs("vt_startdate")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">过期日期</td>
            <td colspan="2"> 
              <input readonly type="text" name="vt_expiredate" id="aa" onclick="getDateString(this,oCalendarChs)" value="<%=rs("vt_expiredate")%>">
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="3"><a href="admin_vote.asp?action=newitem&id=<%=rs("vt_id")%>">新增投票选项</a></td>
          </tr>
<%sql="select * from voteitem where vt_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof
%>
          <tr class=zoneloveds> 
            <td width="18%">投票选项</td>
            <td width="57%"><%=rs2("item_name")%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="newshead"><%=rs2("item_count")%></span>票</td>
            <td width="25%" align="center"><a href="admin_vote.asp?action=edititem&id=<%=rs2("item_id")%>&vt_id=<%=request.querystring("id")%>">编辑</a> 
              <a href="admin_vote.asp?action=delitem&id=<%=rs2("item_id")%>&vt_id=<%=request.querystring("id")%>">删除</a></td>
          </tr>
<%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
          <tr class=zoneloveds> 
            <td colspan="3">当前没有投票选项！</td>
          </tr>
<%end if
rs2.close
set rs2=nothing%>
          <tr class=zoneloveqs align="center"> 
            <td colspan="3">
              <input type="submit" name="Submit" value="确定修改" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("vt_id")%>">
          <input type="hidden" name="action" value="editvote">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%
	  rs.close
set rs=nothing
	  end if
	  end if
	  if request.querystring("action")="newitem" then
if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  response.end
end if
sql="select vt_name,vt_id from votetopic where vt_id="&cint(request.querystring("id"))
set rs=conn.execute(sql)%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form3" method="post" action="">
		<tr class=zonelovess> 
          <td colspan="2">
          新增投票选项</td>
        </tr>
        <tr class=zoneloveds> 
          <td width="13%">所属主题</td>
          <td width="87%"><%=rs("vt_name")%>　</td>
        </tr>
        <tr class=zoneloveds> 
          <td width="13%">选项名称</td>
          <td width="87%">
              <input type="text" name="item_name" size="40" maxlength="50" class="textarea">
            </td>
        </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2">
              <input type="submit" name="Submit" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
        </tr>
		<input type="hidden" name="id" value="<%=rs("vt_id")%>">
		<input type="hidden" name="action" value="newitem">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
      <%rs.close
	  set rs=nothing
	  end if
	  if request.querystring("action")="edititem" then
if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  response.end
end if
sql="select vt_name,vt_id from votetopic where vt_id="&cint(request.querystring("vt_id"))
set rs=conn.execute(sql)
sql="select item_id,item_name from voteitem where item_id="&cint(request.querystring("id"))
set rs2=conn.execute(sql)%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form3" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            修改投票选项</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="13%">所属主题</td>
            <td width="87%"><%=rs("vt_name")%>　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="13%">选项名称</td>
            <td width="87%"> 
              <input type="text" name="item_name" size="40" maxlength="50" class="textarea" value="<%=rs2("item_name")%>">
            </td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="submit" name="Submit" value="确定修改" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_vote.asp?action=vote&caction=edit&id=<%=request.querystring("vt_id")%>">返回</a>] 
            </td>
          </tr>
          <input type="hidden" name="vtid" value="<%=rs("vt_id")%>">
          <input type="hidden" name="id" value="<%=rs2("item_id")%>">
          <input type="hidden" name="action" value="edititem">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%rs2.close
	  set rs2=nothing
	  rs.close
	  set rs=nothing
	  end if
	  if request.querystring("action")="delitem" then
if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  response.end
end if
sql="select vt_name,vt_id from votetopic where vt_id="&cint(request.querystring("vt_id"))
set rs=conn.execute(sql)
sql="select item_id,item_name from voteitem where item_id="&cint(request.querystring("id"))
set rs2=conn.execute(sql)%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form3" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除投票选项</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="13%">所属主题</td>
            <td width="87%"><%=rs("vt_name")%>　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="13%">选项名称</td>
            <td width="87%"><%=rs2("item_name")%> 　</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
              [<a href="admin_vote.asp?action=vote&caction=edit&id=<%=request.querystring("vt_id")%>">返回</a>] </td>
          </tr>
          <input type="hidden" name="vtid" value="<%=rs("vt_id")%>">
          <input type="hidden" name="id" value="<%=rs2("item_id")%>">
          <input type="hidden" name="action" value="delitem">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%rs2.close
	  set rs2=nothing
	  rs.close
	  set rs=nothing
	  end if
	  if request.querystring("action")="vote" and request.querystring("caction")="del" then
	  if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
  response.end
end if
sql="select * from votetopic where vt_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">
            删除投票主题</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">主题名称</td>
            <td width="82%"> <%=rs("vt_name")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">开始日期</td>
            <td width="82%"> <%=rs("vt_startdate")%> 　</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="18%">过期日期</td>
            <td width="82%"> <%=rs("vt_expiredate")%> 　</td>
          </tr>
          <%sql="select * from voteitem where vt_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof
%>
          <tr class=zoneloveds> 
            <td width="18%">投票选项</td>
            <td width="82%"><%=rs2("item_name")%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="newshead"><%=rs2("item_count")%></span>票</td>
          </tr>
          <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
          <tr class=zoneloveds> 
            <td colspan="2">当前没有投票选项！</td>
          </tr>
          <%end if
rs2.close
set rs2=nothing%>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="submit" name="Submit" value="确定删除" class="button">
              [<a href="admin_vote.asp?action=vote">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("vt_id")%>">
          <input type="hidden" name="action" value="delvote">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%
	  rs.close
set rs=nothing
	  end if
end if
%>