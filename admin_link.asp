<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<title><%=webname%>-���ӹ���</title>
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('����δ��¼�����߳�ʱ�ˣ������µ�¼');this.top.location.href='admin.asp';</script>")
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
  Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ�У԰�������id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ�У԰�������id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
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
   Response.Write("<script language=javascript>alert('վ������δ��д��');history.back(1);</script>")
else
  if strLength(flname)>50 then
      founderr=true
	  Response.Write("<script language=javascript>alert('վ������̫���������Գ���50���ַ���');history.back(1);</script>")
  else
      rs("fl_name")=flname
  end if
end if
if flurl="" then
   founderr=true
   Response.Write("<script language=javascript>alert('վ���ַδ��д��');history.back(1);</script>")
else
  if strLength(flurl)>100 then
      founderr=true
	  Response.Write("<script language=javascript>alert('վ���ַ̫���������Գ���100���ַ���');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
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
   Response.Write("<script language=javascript>alert('վ������δ��д��');history.back(1);</script>")
else
  if strLength(flname)>50 then
      founderr=true
	  Response.Write("<script language=javascript>alert('վ������̫���������Գ���50���ַ���');history.back(1);</script>")
  else
      rs("fl_name")=flname
  end if
end if
if flurl="" then
   founderr=true
   Response.Write("<script language=javascript>alert('վ���ַδ��д��');history.back(1);</script>")
else
  if strLength(flurl)>100 then
      founderr=true
	  Response.Write("<script language=javascript>alert('վ���ַ̫���������Գ���100���ַ���');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���id������');history.back(1);</script>")
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
<HTML><HEAD><TITLE>��������</TITLE>
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
		  <td width="33%"><a href="admin_link.asp?action=check">�������(�������:<span class="newshead"><%=passing%></span>)</a></td>
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
          ���ӷ������</font></td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="10%">���</td>
          <td width="70%">��������</td>
          <td width="20%">����</td>
        </tr>
<%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("flcat_id")%>��</td>
          <td><%=rs("flcat_name")%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#ff6600"><%if rs("isimage")=1 then response.write "�ı�����" end if%></font></td>
          <td align="center"><a href="admin_link.asp?id=<%=rs("flcat_id")%>&action=editflcat">�༭</a> 
            <a href="admin_link.asp?id=<%=rs("flcat_id")%>&action=delflcat">ɾ��</a></td>
        </tr>
<%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr class=zoneloveds align="center"> 
          <td colspan="4">��ǰû�����ӷ��࣡</td>
        </tr>
<%end if%>
      </table>
      <br>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
		<tr class=zonelovess> 
          <td>
          �������ӷ���</font></td>
        </tr>
        <tr class=zoneloveds> 
            <td>��������-
              <input type="text" name="flcat_name" size="40" class="textarea">
              &nbsp;&nbsp;&nbsp;
              <input type="checkbox" name="isimage" value="1" class="textarea">
              �ı�����</td>
        </tr>
        <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
        </tr>
		<input type="hidden" name="action" value="newflcat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.querystring("action")="editflcat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ������ӷ���ID������');history.back(1);</script>")
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
            �޸����ӷ���</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>��������- 
              <input type="text" name="flcat_name" size="40" class="textarea" value="<%=rs("flcat_name")%>">
			  &nbsp;&nbsp;&nbsp;
              <input type="checkbox" name="isimage" value="1" class="textarea" <%if rs("isimage")=1 then response.write "checked" end if%>>
              �ı����� </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ������ӷ���ID������');history.back(1);</script>")
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
            ɾ�����ӷ���</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>��������- <%=rs("flcat_name")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="ȷ��ɾ��" class="button">
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
        <tr class=zonelovess><td colspan="4">���ӹ���</td>
          <td align="right" style="padding-top:2px;">
            <select name="go" style="margin:-3px" onChange='window.location=form.go.options[form.go.selectedIndex].value'>
			  <option value="">ѡ����ʾ��ʽ</option>
			  <option value="admin_link.asp?action=link">��ʾ��������</option>
<%sql="select * from flcat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_link.asp?action=link&flcat_id=<%=rs2("flcat_id")%>"><%=rs2("flcat_name")%></option>
<%rs2.movenext
loop
if rs2.bof and rs2.eof then%><option value="">��ǰû�з���</option>
<%end if
rs2.close
set rs2=nothing%>
            </select>
          </td>
        </tr>
		</form>
        <tr class=zoneloveqs align="center"> 
          <td width="50">���</td>
          <td width="100">LOGO</td>
          <td width="*">��������</td>
          <td width="100">����</td>
          <td width="120">����</td>
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
          <td align="center"><%=rs("fl_id")%></td><td align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"' width='88' height='31' border='0'>" else response.write "<font color='#cccccc'>��</font>" end if%></td>
          <td><a href="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td align="center"><%=rs2("flcat_name")%></td>
          <td align="center"><a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lklk">������</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=editfl">�༭</a> 
            <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=delfl">ɾ��</a></td>
        </tr>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr class=zoneloveds> 
          <td colspan="5" align="center">��ǰ��û�����ӣ�</td>
        </tr>
<%end if
end if%>
        <form name="form1" method="post" action="admin_link.asp?action=link&flcat_id=<%=request.querystring("flcat_id")%>">
          <tr class=zoneloveqs> 
            <td align="right" colspan="5"> <%=currentpage%> /<%=totalpages%>ҳ,<%=totalfl%>����¼/<%=adflperpage%>ƪÿҳ. 
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
              <a href="admin_link.asp?action=link&page=<%=page%>&flcat_id=<%=request.querystring("flcat_id")%>" title="��һҳ">>></a> 
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            �޸�����</font></td>
          </tr>
          <tr class=zoneloveds><td align="right" width="30%">ѡ����� </td><td width="70%">
              <select name="flcat_id" id="flcat_id">
<%sql="select * from flcat"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%><option value="<%=rs("flcat_id")%>" <%if rs("flcat_id")=rs2("flcat_id") then response.write "selected" end if%>><%=rs("flcat_name")%></option><%rs.movenext
loop%></select></tr>
<tr class=zoneloveds><td align="right" width="30%">վ������ </td><td width="70%"><input type="text" name="fl_name" size="40" class="textarea" value="<%=rs2("fl_name")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ </td><td width="70%">
              <input type="text" name="fl_url" size="40" class="textarea" value="<%=rs2("fl_url")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ </td><td width="70%">
              <input type="text" name="fl_logo" size="40" class="textarea" value="<%=rs2("fl_logo")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td align="right">�ϴ�LOGO�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=link"></iframe></td>
            </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ����ͨ��</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">ѡ����� </td><td width="70%">
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
<tr class=zoneloveds><td align="right" width="30%">վ������</td><td width="70%"><input type="text" name="fl_name" size="40" class="textarea" value="<%=rs2("fl_name")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ</td><td width="70%"><input type="text" name="fl_url" size="40" class="textarea" value="<%=rs2("fl_url")%>"></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ</td><td width="70%"><input type="text" name="fl_logo" size="40" class="textarea" value="<%=rs2("fl_logo")%>">
            </td>
          </tr>
<tr class=zoneloveds> 
              <td align="right">�ϴ�LOGO�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=link"></iframe></td>
            </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ��ͨ��" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ȡ��������</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">������� </td><td width="70%">
<%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">վ������</td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ</td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ</td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ��ȡ��" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ���������</td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">������� </td><td width="70%">
<%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">վ������</td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ</td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ</td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ������" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ɾ������</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">������� </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">վ������: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ: </td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ��ɾ��" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ɾ������������</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">������� </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">վ������: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ: </td><td width="70%"><%=rs2("fl_logo")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ��ɾ��" class="button">
              <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��ķ���ID������');history.back(1);</script>")
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
            ɾ������</font></td>
          </tr>
<tr class=zoneloveds><td align="right" width="30%">������� </td><td width="70%">
                <%sql="select * from flcat where flcat_id="&rs2("flcat_id")
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
%><%=rs("flcat_name")%>
<tr class=zoneloveds><td align="right" width="30%">վ������: </td><td width="70%"><%=rs2("fl_name")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">վ���ַ: </td><td width="70%"><%=rs2("fl_url")%></td></tr>
<tr class=zoneloveds><td align="right" width="30%">Logo��ַ: </td><td width="70%"><%=rs2("fl_logo")%></td></tr>
          <tr class=zoneloveqs> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ��ɾ��" class="button">
 <input type="button" value="����"  onClick="location.href='<%=request.serverVariables("Http_REFERER")%>'" class="button">
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
            �������</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
          <td width="50">���</td>
          <td width="100">LOGO</td>
          <td width="*">��������</td>
          <td width="100">����</td>
          <td width="100">����</td>
          </tr>
          <%for i=1 to rs.recordcount
sql="select flcat_name from flcat where flcat_id="&rs("flcat_id")
set rs2=conn.execute(sql)%>
          <tr class=zoneloveds> 
            <td align="center"><%=rs("fl_id")%></td><td width="100" align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"'>" else response.write "<font color='#cccccc'>��</font>" end if%></td>
            <td><a href="<%=rs("fl_url")%>" title="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td width="100" align="center"><%=rs2("flcat_name")%></td>
            <td align="center"> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=pass">ͨ��</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=delnopass">ɾ��</a></td>
          </tr>
          <%rs.movenext
next
if rs.bof and rs.eof then%>
          <tr class=zoneloveds align="center"> 
            <td colspan="5">��ǰû�д���˵����ӣ�</td>
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
            ����������</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
          <td width="50">���</td>
          <td width="100">LOGO</td>
          <td width="*">��������</td>
          <td width="100">����</td>
          <td width="100">����</td>
          </tr>
          <%for i=1 to rs.recordcount
sql="select flcat_name from flcat where flcat_id="&rs("flcat_id")
set rs2=conn.execute(sql)%>
          <tr class=zoneloveds> 
            <td align="center"><%=rs("fl_id")%></td><td width="100" align="center"><%if rs("fl_logo")<>"" then response.write "<img src='"&rs("fl_logo")&"'>" else response.write "<font color='#cccccc'>��</font>" end if%></td>
            <td><a href="<%=rs("fl_url")%>" title="<%=rs("fl_url")%>" target="_blank"><%=rs("fl_name")%></a></td><td width="100" align="center"><%=rs2("flcat_name")%></td>
            <td align="center"> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lkpass">ȡ��</a> <a href="admin_link.asp?id=<%=rs("fl_id")%>&action=lkdel">ɾ��</a></td>
          </tr>
          <%rs.movenext
next
if rs.bof and rs.eof then%>
          <tr class=zoneloveds align="center"> 
            <td colspan="5">��ǰû�к��������ӣ�</td>
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