<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<!--#include file="inc/UBBshow.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('����δ��¼�����߳�ʱ�ˣ������µ�¼');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newcat" then
sql="select * from d_cat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim cat_name
cat_name=trim(replace(request.form("cat_name"),"'",""))

if cat_name="" then
   founderr=true
Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
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
  response.redirect "admin_down.asp"
end if
elseif request.form("action")="newclass" then
sql="select * from d_class"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim class_name,cat_id
cat_id=cint(request.form("cat_id"))
class_name=trim(replace(request.form("class_name"),"'",""))

if cat_id<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�����ѡ�������������ƣ�');history.back(1);</script>")
else
   rs("cat_id")=cat_id
end if
if class_name="" then
   founderr=true
Response.Write("<script language=javascript>alert('�������д�ӷ������ƣ�');history.back(1);</script>")
else
   rs("class_name")=class_name
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_down.asp"
end if
elseif request.form("action")="editcat" then
if cint(request.form("id"))<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
   response.end
end if
sql="select * from d_cat where cat_id="&request.form("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cat_name=trim(replace(request.form("cat_name"),"'",""))

if cat_name="" then
   founderr=true
Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
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
  response.redirect "admin_down.asp"
end if
elseif request.form("action")="editclass" then
if cint(request.form("id"))<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
   response.end
end if
sql="select * from d_class where class_id="&request.form("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cat_id=cint(request.form("cat_id"))
class_name=trim(replace(request.form("class_name"),"'",""))

if cat_id<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�����ѡ�������������ƣ�');history.back(1);</script>")
else
   rs("cat_id")=cat_id
end if
if class_name="" then
   founderr=true
Response.Write("<script language=javascript>alert('�������д�ӷ������ƣ�');history.back(1);</script>")
else
   rs("class_name")=class_name
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_down.asp"
end if
elseif request.form("action")="delcat" then
if cint(request.form("id"))<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
   response.end
end if
sql="select * from d_cat where cat_id="&request.form("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_down.asp"
elseif request.form("action")="delclass" then
if cint(request.form("id"))<1 then
   founderr=true
Response.Write("<script language=javascript>alert('�Ƿ��ķ��������');history.back(1);</script>")
   response.end
end if
sql="select * from d_class where class_id="&request.form("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_down.asp"
end if
end if
sql="select * from d_cat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>��������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<%if request.querystring("action")="" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr> 
          <td colspan="2" class=zonelovess>���ط������</font></td>
        </tr>
<%do while not rs.eof%><tr class=zoneloveqs><td width="70%"><%=rs("cat_name")%>&nbsp;[<a href="admin_down.asp?action=newclass">�½��ӷ���</a>]</td><td width="30%">[<a href="admin_down.asp?action=editcat&id=<%=rs("cat_id")%>&name=<%=rs("cat_name")%>">�޸�</a>]&nbsp;[<a href="admin_down.asp?action=delcat&id=<%=rs("cat_id")%>&name=<%=rs("cat_name")%>">ɾ��</a>]</td></tr>
<%sql="select * from d_class where cat_id="&rs("cat_id")
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<tr class=zoneloveds><td width="70%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs2("class_name")%></td><td width="30%">[<a href="admin_down.asp?action=editclass&id=<%=rs2("class_id")%>&cid=<%=rs("cat_id")%>&name=<%=rs2("class_name")%>">�޸�</a>]&nbsp;[<a href="admin_down.asp?action=delclass&id=<%=rs2("class_id")%>&cid=<%=rs("cat_id")%>&name=<%=rs2("class_name")%>">ɾ��</a>]</td></tr>
<%
rs2.movenext
loop
response.write "<tr class=zoneloveqs><td>"
if rs2.bof and rs2.eof then
response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ǰû���ӷ��࣡"
end if
rs2.close
set rs2=nothing
rs.movenext
loop
if rs.bof and rs.eof then
response.write "<tr class=zoneloveqs><td>��ǰû�з��࣡"
end if
%>
</td></tr></table>  
<%
end if
if request.querystring("action")="newcat" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="admin_down.asp">
		<tr> 
          <td class=zonelovess>�½�����</td>
        </tr>
        <tr> 
          <td class=zoneloveds>��������-
            <input type="text" name="cat_name" size="40" class="textarea">
          </td>
        </tr>
        <tr> 
          <td class=zoneloveqs align="center">
              <input type="submit" name="Submit" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="ȡ������" class="button">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
        </tr>
		<input type="hidden" name="action" value="newcat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
<%end if
if request.querystring("action")="newclass" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
	    <form name="form2" method="post" action="admin_down.asp">
        <tr> 
          <td class=zonelovess>�½��ӷ���</td>
        </tr>
        <tr> 
          <td class=zoneloveds>�ӷ�������-
              <input type="text" name="class_name" size="50" class="textarea">
              <br>
              ��������-
              <select name="cat_id">
                <%sql="select * from d_cat"
set rs=conn.execute(sql)
do while not rs.eof
%><option value="<%=rs("cat_id")%>"><%=rs("cat_name")%></option>
<%rs.movenext
loop
if rs.eof and rs.bof then%><option value="0">���޷���</option>
<%end if%>
              </select>
            </td>
        </tr>
        <tr> 
            <td class=zoneloveqs align="center"> 
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset2" value="ȡ������" class="button">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
        </tr>
		<input type="hidden" name="action" value="newclass">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
<%end if
if request.querystring("action")="editcat" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
	    <form name="form1" method="post" action="admin_down.asp">
        <tr> 
          <td class=zonelovess>�޸ķ���</td>
        </tr>
        <tr> 
            <td class=zoneloveds>��������- 
              <input type="text" name="cat_name" size="40" class="textarea" value="<%=request.querystring("name")%>">
          </td>
        </tr>
        <tr> 
            <td class=zoneloveqs align="center">
              <input type="submit" name="Submit3" value="ȷ���޸�" class="button">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
        </tr>
		<input type="hidden" name="action" value="editcat">
		<input type="hidden" name="id" value="<%=request.querystring("id")%>">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
<%end if
if request.querystring("action")="editclass" then%>
      <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="admin_down.asp">
          <tr> 
            <td class=zonelovess>�޸��ӷ���</td>
          </tr>
          <tr> 
            <td class=zoneloveds>�ӷ�������- 
              <input type="text" name="class_name" size="50" class="textarea" value="<%=request.querystring("name")%>">
              <br>
              ��������- 
              <select name="cat_id">
                <%sql="select * from d_cat"
set rs=conn.execute(sql)
do while not rs.eof
%>
                <option value="<%=rs("cat_id")%>" <%if rs("cat_id")=cint(request.querystring("cid")) then response.write "selected" end if%>><%=rs("cat_name")%></option>
                <%rs.movenext
loop
if rs.eof and rs.bof then%>
                <option value="0">���޷���</option>
                <%end if%>
              </select>
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
          </tr>
          <input type="hidden" name="action" value="editclass">
		  <input type="hidden" name="id" value="<%=request.querystring("id")%>">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="delcat" then%>
     <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="admin_down.asp">
          <tr> 
            <td class=zonelovess>ɾ������</td>
          </tr>
          <tr> 
            <td class=zoneloveds>��������- 
              <input type="text" name="cat_name2" size="40" class="textarea" value="<%=request.querystring("name")%>"> 
            </td>
          </tr>
          <tr> 
            <td class=zoneloveqs align="center"> <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
          </tr>
          <input type="hidden" name="action" value="delcat">
          <input type="hidden" name="id" value="<%=request.querystring("id")%>">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="delclass" then%>
     <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="admin_down.asp">
          <tr> 
            <td class=zonelovess>ɾ���ӷ���</td>
          </tr>
          <tr> 
            <td class=zoneloveds>�ӷ�������-<%=request.querystring("name")%>
              <br>
              ��������- 
                <%sql="select * from d_cat where cat_id="&cint(request.querystring("cid"))
set rs=conn.execute(sql)
%>
<%=rs("cat_name")%>
</td>
          </tr>
          <tr> 
            <td class=zoneloveqs align="center"> <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��">
<input type="button" value="����"  onClick="location.href='admin_down.asp'" class="button"></td>
          </tr>
          <input type="hidden" name="action" value="delclass">
          <input type="hidden" name="id" value="<%=request.querystring("id")%>">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.querystring("action")="list" then
dim keyword,class_id,colname,totalsoft,Currentpage,totalpages,i
keyword=trim(replace(request.form("keyword"),"'",""))
cat_id=request("cat_id")
class_id=request("class_id")
colname=request("colname")

sql="select soft_id,soft_name,soft_joindate,soft_catname,soft_classname,soft_rcount,soft_dcount,soft_desc,review,reviewcount from soft order by soft_joindate desc"
if cat_id<>"" and class_id="" then
    sql="SELECT soft_id,soft_name,soft_joindate,soft_catname,soft_classname,soft_rcount,soft_dcount,soft_desc,review,reviewcount FROM soft where soft_catid="&cat_id&" order by soft_joindate desc"
elseif class_id<>"" and keyword="" then
	sql="SELECT soft_id,soft_name,soft_joindate,soft_catname,soft_classname,soft_rcount,soft_dcount,soft_desc,review,reviewcount FROM soft where soft_classid="&class_id&" order by soft_joindate desc"
elseif keyword<>"" and colname<>"0" then
	sql="select soft_id,soft_name,soft_joindate,soft_catname,soft_classname,soft_rcount,soft_dcount,soft_desc,review,reviewcount from soft where "&colname&" like '%"&keyword&"%' order by soft_joindate desc"
else
	sql="SELECT soft_id,soft_name,soft_joindate,soft_catname,soft_classname,soft_rcount,soft_dcount,soft_desc,review,reviewcount FROM soft order by soft_joindate desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
 <table align="center" width="98%" align="center" border="0" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zoneloveqs><form name="form1" method="post" action="admin_down.asp?action=list"><td> 
<select name="colname"><option value=0 selected>--==������Χ==--</option><option value="soft_name">�������</option><option value="soft_desc">�������</option></select>&nbsp; <input type="text" name="keyword" class="textarea" size="20">&nbsp;&nbsp; <input type="submit" name="Submit" value="����" class="button"></td></form><form><td >��ת��<select name="go" onChange='window.location=form.go.options[form.go.selectedIndex].value'><option value="admin_down.asp?action=list">---===ѡ�����===---</option><option value="admin_down.asp?action=list">---===���������ҳ===---</option><%
Set rscat = Server.CreateObject("ADODB.Recordset")
sqlcat="SELECT * FROM d_cat"
rscat.OPEN sqlcat,Conn,1,1
do while not rscat.eof
colparam1=rscat("cat_id")
%><option value="admin_down.asp?action=list&cat_id=<%=rscat("cat_id")%>">----<%=rscat("cat_name")%>----</option>
                      <%
Set rsclass = Server.CreateObject("ADODB.Recordset")
sqlclass="SELECT * FROM d_class where cat_id=" &colparam1& ""
rsclass.OPEN sqlclass,Conn,1,1
	do while not rsclass.eof
%><option value='admin_down.asp?action=list&cat_id=<%=rsclass("cat_id")%>&class_id=<%=rsclass("class_id")%>'><%=rsclass("class_name")%></option>
<%
        rsclass.movenext
	loop
	rsclass.close
set rsclass=nothing
rscat.movenext
loop
rscat.close
set rscat=nothing
set rstotal=nothing
%></select></td></form><td align="right"> 
<%dim flname
			flname=""
			if request.querystring("class_id")<>"" then
			sql="select class_name from d_class where class_id="&request.querystring("class_id")
			set rsclass=conn.execute(sql)
			flname=rsclass("class_name")
			rsclass.close
			set rsclass=nothing
			elseif request.querystring("class_id")="" and request.querystring("cat_id")<>"" then
			sql="select cat_name from d_cat where cat_id="&request.querystring("cat_id")
			set rscat=conn.execute(sql)
			flname=rscat("cat_name")
			rscat.close
			set rscat=nothing
			end if%>
                  <font color="#ff6600"><%=flname%></font>���๲��<%=rs.recordcount%>������ </td>
              </tr>
            </table>
            <br>
<%if not rs.eof then
rs.movefirst
rs.pagesize=softperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalsoft=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*softperpage<totalsoft then
	       rs.move(currentpage-1)*softperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalsoft mod softperpage)=0 then
      totalpages=totalsoft\softperpage
   else
      totalpages=totalsoft\softperpage+1
   end if
i=0
do while not rs.eof and i<softperpage%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
              <tr class="zoneloveqs"> 
                <td width="60%"><a href="#"><b><%=rs("soft_name")%></b></a> <span class="date"><%=rs("soft_joindate")%></span> <span class="newshead"><%=rs("soft_dcount")%></span>/<span class="newshead"><%=rs("soft_rcount")%></span></td><td width="20%" align="center"><font color=red><%=rs("reviewcount")%></font></span>ƪ����</td>
                <td width="20%" align="center"><a href="admin_view.asp?action=down&soft_id=<%=rs("soft_id")%>">����</a> <a href="admin_down.asp?id=<%=rs("soft_id")%>&action=editsoft">�༭</a> 
                  <a href="admin_down.asp?id=<%=rs("soft_id")%>&action=delsoft">ɾ��</a> </td>
              </tr>
              <tr class="zoneloveds"> 
                <td colspan="3"><%=cutstr(rs("soft_desc"),40,false,none)%></td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="8"></td>
              </tr>
            </table>
<%i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
              <tr class="zonelovess"> 
                <td align="center">��ǰû�г���</td>
              </tr>
            </table>
<%end if
end if%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form3" method="post" action="admin_down.asp?action=list&cat_id=<%=request("cat_id")%>&keyword=<%=request("keyword")%>&colname=<%=request("colname")%>">
          <tr class="zoneloveds"> 
            <td align="right"><%=currentpage%> /<%=totalpages%>ҳ,<%=totalsoft%>����¼/<%=softperpage%>��ÿҳ. 
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
              <a href="admin_down.asp?action=list&page=<%=i%>&cat_id=<%=request("cat_id")%>&keyword=<%=request("keyword")%>&colname=<%=request("colname")%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_down.asp?action=list&page=<%=i%>&cat_id=<%=request("cat_id")%>&class_id=<%=request("class_id")%>" title="��һҳ">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
<%end if
if request.form("MM_insert") then

if request.Form("action")="newsoft" then
sql="select * from soft"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim soft_name,soft_size,soft_mode,roof,commend,demo,home,showpic,desc,downisbest,url1,url2,url3,url4
soft_name=trim(replace(request.form("name"),"'",""))
soft_size=trim(replace(request.form("size"),"'",""))
soft_mode=trim(replace(request.form("mode"),"'",""))
roof=trim(replace(request.form("roof"),"'",""))
commend=cint(request.form("commend"))
demo=trim(replace(request.form("demo"),"'",""))
home=trim(replace(request.form("home"),"'",""))
showpic=trim(replace(request.form("showpic"),"'",""))
desc=trim(replace(request.form("desc"),"'",""))
downisbest=request.form("isbest")
url1=trim(replace(request.form("url1"),"'",""))
url2=trim(replace(request.form("url2"),"'",""))
url3=trim(replace(request.form("url3"),"'",""))
url4=trim(replace(request.form("url4"),"'",""))

if soft_name="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д�ļ����ƣ�');history.back(1);</script>")
else
  rs("soft_name")=soft_name
end if
if soft_size="" then
  rs("soft_size")="δ֪"
else
  rs("soft_size")=soft_size
end if
if soft_mode="" then
  rs("soft_mode")=""
else
  rs("soft_mode")=soft_mode
end if
if roof="" then
  rs("soft_roof")=""
else
  rs("soft_roof")=roof
end if
if commend<1 then
  founderr=true
Response.Write("<script language=javascript>alert('�����ѡ���ļ��Ƽ��̶ȣ�');history.back(1);</script>")
else
  rs("soft_commend")=commend
end if
if cint(downisbest)=1 then
  rs("isbest")=cint(downisbest)
end if
if request("of")<>"" then
		Set rsclass = Server.CreateObject("ADODB.Recordset")
		sqlclass="SELECT * FROM d_class where class_id="&request("of")
		rsclass.OPEN sqlclass,Conn,1,1
		rs("soft_classid")=rsclass("class_id")
		rs("soft_classname")=rsclass("class_name")
		rs("soft_catid")=rsclass("cat_id")
		set rscat = Server.CreateObject("ADODB.Recordset")
		sqlcat="select cat_name from d_cat where cat_id="&rsclass("cat_id")
		rscat.open sqlcat,conn,1,1
		rs("soft_catname")=rscat("cat_name")
		rscat.close
		rsclass.close
		set rscat=nothing
		set rsclass=nothing
else
		founderr=true
Response.Write("<script language=javascript>alert('�����ѡ���������࣬���߷Ƿ�������');history.back(1);</script>")
end if
if demo="" then
  rs("soft_demo")=""
else
  rs("soft_demo")=demo
end if
if home="" then
  rs("soft_home")=""
else
  rs("soft_home")=home
end if
if showpic="" then
  rs("soft_showpic")=""
else
  rs("soft_showpic")=showpic
end if
if desc="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д�ļ���飡');history.back(1);</script>")
else
  rs("soft_desc")=desc
end if
if url1="" and url2="" and url3="" and url4="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д���ص�ַ��');history.back(1);</script>")
end if
if url1="" then
  rs("soft_url1")=""
else
  rs("soft_url1")=url1
end if
if url2="" then
  rs("soft_url2")=""
else
  rs("soft_url2")=url2
end if
if url3="" then
  rs("soft_url3")=""
else
  rs("soft_url3")=url3
end if
if url4="" then
  rs("soft_url4")=""
else
  rs("soft_url4")=url4
end if

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  sql="update allcount set softcount = softcount + 1"
  conn.execute(sql)
  
  response.redirect "admin_down.asp?action=list"
end if
end if

'edit
if request.Form("action")="editsoft" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('�Ƿ������id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from soft where soft_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
soft_name=trim(replace(request.form("name"),"'",""))
soft_size=trim(replace(request.form("size"),"'",""))
soft_mode=trim(replace(request.form("mode"),"'",""))
roof=trim(replace(request.form("roof"),"'",""))
commend=cint(request.form("commend"))
demo=trim(replace(request.form("demo"),"'",""))
home=trim(replace(request.form("home"),"'",""))
showpic=trim(replace(request.form("showpic"),"'",""))
desc=trim(replace(request.form("desc"),"'",""))
downisbest=request.form("isbest")
url1=trim(replace(request.form("url1"),"'",""))
url2=trim(replace(request.form("url2"),"'",""))
url3=trim(replace(request.form("url3"),"'",""))
url4=trim(replace(request.form("url4"),"'",""))

if soft_name="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д�ļ����ƣ�');history.back(1);</script>")
else
  rs("soft_name")=soft_name
end if
if soft_size="" then
  rs("soft_size")="δ֪"
else
  rs("soft_size")=soft_size
end if
if soft_mode="" then
  rs("soft_mode")=""
else
  rs("soft_mode")=soft_mode
end if
if roof="" then
  rs("soft_roof")=""
else
  rs("soft_roof")=roof
end if
if commend<1 then
  founderr=true
Response.Write("<script language=javascript>alert('�����ѡ���ļ��Ƽ��̶ȣ�');history.back(1);</script>")
else
  rs("soft_commend")=commend
end if
rs("isbest")=cint(downisbest)
if request("of")<>"" then
		Set rsclass = Server.CreateObject("ADODB.Recordset")
		sqlclass="SELECT * FROM d_class where class_id="&request("of")
		rsclass.OPEN sqlclass,Conn,1,1
		rs("soft_classid")=rsclass("class_id")
		rs("soft_classname")=rsclass("class_name")
		rs("soft_catid")=rsclass("cat_id")
		set rscat = Server.CreateObject("ADODB.Recordset")
		sqlcat="select cat_name from d_cat where cat_id="&rsclass("cat_id")
		rscat.open sqlcat,conn,1,1
		rs("soft_catname")=rscat("cat_name")
		rscat.close
		rsclass.close
		set rscat=nothing
		set rsclass=nothing
else
		founderr=true
Response.Write("<script language=javascript>alert('�����ѡ���������࣬���߷Ƿ�������');history.back(1);</script>")
end if
if demo="" then
  rs("soft_demo")=""
else
  rs("soft_demo")=demo
end if
if home="" then
  rs("soft_home")=""
else
  rs("soft_home")=home
end if
if showpic="" then
  rs("soft_showpic")=""
else
  rs("soft_showpic")=showpic
end if
if desc="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д�ļ���飡');history.back(1);</script>")
else
  rs("soft_desc")=desc
end if
if url1="" and url2="" and url3="" and url4="" then
  founderr=true
Response.Write("<script language=javascript>alert('�������д���ص�ַ��');history.back(1);</script>")
end if
if url1="" then
   rs("soft_url1")=""
else
  rs("soft_url1")=url1
end if
if url2="" then
  rs("soft_url2")=""
else
  rs("soft_url2")=url2
end if
if url3="" then
  rs("soft_url3")=""
else
  rs("soft_url3")=url3
end if
if url4="" then
  rs("soft_url4")=""
else
  rs("soft_url4")=url4
end if

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_down.asp?action=list"
end if
end if

'delete
if request.Form("action")="delsoft" then
if request.Form("id")="" then
  founderr=true
Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
Response.Write("<script language=javascript>alert('�Ƿ����ļ�id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from soft where soft_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  sql="update allcount set softcount = softcount - 1"
  conn.execute(sql)
  
  response.redirect "admin_down.asp?action=list"
end if
end if
sql="select * from d_cat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if request.QueryString("action")="newsoft" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
              <form name="form1" method="post" action="admin_down.asp?action=list">
			  <tr class=zonelovess> 
                <td colspan="2">�������</td>
              </tr>
              <tr  class=zoneloveds> 
                <td width="16%">�������</td>
                <td width="84%">
                    <input type="text" name="name" size="40" class="textarea">
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>�����С</td>
                <td><input type="text" name="size" size="20" class="textarea"></td>
              </tr>
              <tr class=zoneloveds> 
                <td>��Ȩ��ʽ</td>
                <td>
                    <select name="mode">
		     <option value="������" selected>������</option>
                      <option value="�������">�������</option>
                      <option value="�������">�������</option>
                      <option value="�ƽ����">�ƽ����</option>
                    </select>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>Ӧ��ƽ̨</td>
                <td>
                    <select name="roof">
		     <option value="Win9x/Me">Win9x/Me</option>
		     <option value="IIS����" selected>IIS����</option>
                      <option value="WinNT/2000/XP">WinNT/2000/XP</option>
                      <option value="Win9x/Me/NT/2000/XP">Win9x/Me/NT/2000/XP</option>
                    </select>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>�Ƽ��̶�</td>
                <td>
                    <select name="commend">
		     <option value="5">5stars</option>
                      <option value="4">4stars</option>
                      <option value="3" selected>3stars</option>
                      <option value="2">2stars</option>
                      <option value="1">1stars</option>
                      <option value="0">0stars</option>
                    </select>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>��������</td>
                <td>
                    <select name="of">
                      <%
dim catname,classname,catid,classid
Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM d_cat"
rs.OPEN sql,Conn,1,1
if not (rs.eof and rs.bof) then
rs.movefirst
do while not rs.eof
sql="select * from d_class where cat_id="&rs("cat_id")
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if not (rs2.eof and rs2.bof) then
do while not rs2.eof
catid=rs("cat_id")
classid=rs2("class_id")
catname=rs("cat_name")
classname=rs2("class_name")
%>
                      <option value='<%=rs2("class_id")%>'><%=rs2("class_name")%>��<%=rs("cat_name")%>��</option>
<%
rs2.movenext
loop
end if
rs.movenext
rs2.close
set rs2=nothing
loop
end if
%>
                    </select>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>��ʾ��ַ</td>
                <td>
                    <input type="text" name="demo" size="60" class="textarea">
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>��ҳ��ַ</td>
                <td>
                    <input type="text" name="home" size="60" class="textarea" value="">
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>ͼƬ��ַ</td>
                <td>
                    <input type="text" name="showpic" size="60" class="textarea" value="">
                  </td>
              </tr>
             <tr class=zoneloveds> 
              <td> �ϴ�ͼƬ�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=downpic"></iframe></td>
            </tr>
              <tr class=zoneloveds> 
                <td>������</td>
                <td>
                    <textarea name="desc" cols="65" rows="10" class="textarea">û��</textarea>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>��������</td>
                <td>
                    <input type="text" name="url1" size="60" class="textarea">(�Ƽ�ʹ��)
                  </td>
              </tr>
             <tr class=zoneloveds> 
              <td> �ϴ������ </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=down"></iframe>(�����С����500K�������)</td>
            </tr>
              <tr class=zoneloveds> 
                <td>��������</td>
                <td>
                    <input type="text" name="url2" size="60" class="textarea">
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>Զ������1</td>
                <td>
                    <input type="text" name="url3" size="60" class="textarea">
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>Զ������2</td>
                <td>
                    <input type="text" name="url4" size="60" class="textarea">
                  </td>
              </tr>
              <tr class=zoneloveqs> 
                <td colspan="2" align="center" height="30">
<input type="checkbox" name="isbest" value="1">�Ƽ�&nbsp;&nbsp;&nbsp; <input type="submit" name="Submit" value="ȷ������" class="button">
                    <input type="reset" name="Reset" value="�������" class="button">
<input type="button" value="����"  onClick="location.href='admin_down.asp?action=list'" class="button"></td>
              </tr>
			  <input type="hidden" name="action" value="newsoft">
			  <input type="hidden" name="MM_insert" value="true">
			  </form>
            </table>
			<%end if
			if request.QueryString("action")="editsoft" then
			if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('�Ƿ������ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from soft where soft_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
              <form name="form1" method="post" action="admin_down.asp?action=list">
                <tr class=zonelovess> 
                  <td colspan="2">�༭���</td>
                </tr>
                <tr class=zoneloveds> 
                  <td width="16%">�������</td>
                  <td width="84%"> <input type="text" name="name" size="40" class="textarea" value="<%=rs2("soft_name")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>�����С</td>
                  <td> <input type="text" name="size" size="20" class="textarea" value="<%=rs2("soft_size")%>"> 
                  </td>
                </tr>
              <tr class=zoneloveds> 
                <td>��Ȩ��ʽ</td>
                <td>
                    <select name="mode">
		     <option value="������" <%if rs2("soft_mode")="������" then response.write "selected" end if%>>������</option>
                      <option value="�������" <%if rs2("soft_mode")="�������" then response.write "selected" end if%>>�������</option>
                      <option value="�������" <%if rs2("soft_mode")="�������" then response.write "selected" end if%>>�������</option>
                      <option value="�ƽ����" <%if rs2("soft_mode")="�ƽ����" then response.write "selected" end if%>>�ƽ����</option>
                    </select>
                  </td>
              </tr>
              <tr class=zoneloveds> 
                <td>Ӧ��ƽ̨</td>
                <td>
                    <select name="roof">
		     <option value="IIS����" <%if rs2("soft_roof")="IIS����" then response.write "selected" end if%>>IIS����</option>
		     <option value="Win9x/Me" <%if rs2("soft_roof")="Win9x/Me" then response.write "selected" end if%>>Win9x/Me</option>
                      <option value="WinNT/2000/XP" <%if rs2("soft_roof")="WinNT/2000/XP" then response.write "selected" end if%>>WinNT/2000/XP</option>
                      <option value="Win9x/Me/NT/2000/XP" <%if rs2("soft_roof")="Win9x/Me/NT/2000/XP" then response.write "selected" end if%>>Win9x/Me/NT/2000/XP</option>
                    </select>
                  </td>
              </tr>
                <tr class=zoneloveds> 
                  <td>�Ƽ��̶�</td>
                  <td> <select name="commend" id="commend">
                      <option value="5" <%if rs2("soft_commend")=5 then response.write "selected" end if%>>5stars</option>
                      <option value="4" <%if rs2("soft_commend")=4 then response.write "selected" end if%>>4stars</option>
                      <option value="3" <%if rs2("soft_commend")=3 then response.write "selected" end if%>>3stars</option>
                      <option value="2" <%if rs2("soft_commend")=2 then response.write "selected" end if%>>2stars</option>
                      <option value="1" <%if rs2("soft_commend")=1 then response.write "selected" end if%>>1stars</option>
                      <option value="0" <%if rs2("soft_commend")=0 then response.write "selected" end if%>>0stars</option>
                    </select> </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td> <select name="of" id="of">
                      <%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM d_cat"
rs.OPEN sql,Conn,1,1
if not (rs.eof and rs.bof) then
rs.movefirst
do while not rs.eof
sql="select * from d_class where cat_id="&rs("cat_id")
Set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.OPEN sql,Conn,1,1
if not (rs3.eof and rs3.bof) then
do while not rs3.eof
catid=rs("cat_id")
classid=rs3("class_id")
catname=rs("cat_name")
classname=rs3("class_name")
%>
                      <option value='<%=rs3("class_id")%>' <%if rs3("class_id")=rs2("soft_classid") then response.Write "selected" end if%>><%=rs3("class_name")%>��<%=rs("cat_name")%>��</option>
                      <%
rs3.movenext
loop
end if
rs.movenext
rs3.close
set rs3=nothing
loop
end if
%>
                    </select> </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��ʾ��ַ</td>
                  <td> <input type="text" name="demo" size="60" class="textarea" value="<%=rs2("soft_demo")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��ҳ��ַ</td>
                  <td> <input type="text" name="home" size="60" class="textarea" value="<%=rs2("soft_home")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>ͼƬ��ַ</td>
                  <td> <input type="text" name="showpic" size="60" class="textarea" value="<%=rs2("soft_showpic")%>"> 
                  </td>
                </tr>
             <tr class=zoneloveds> 
              <td> �ϴ�ͼƬ�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=downpic"></iframe></td>
            </tr>
                <tr class=zoneloveds> 
                  <td>������</td>
                  <td> <textarea name="desc" cols="65" rows="10" class="textarea" id="desc"><%=rs2("soft_desc")%></textarea> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td> <input type="text" name="url1" size="60" class="textarea" value="<%=rs2("soft_url1")%>"> 
                  </td>
                </tr>
             <tr class=zoneloveds> 
              <td> �ϴ������ </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=down"></iframe></td>
            </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td> <input type="text" name="url2" size="60" class="textarea" value="<%=rs2("soft_url2")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>Զ������1</td>
                  <td> <input type="text" name="url3" size="60" class="textarea" value="<%=rs2("soft_url3")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>Զ������2</td>
                  <td> <input type="text" name="url4" size="60" class="textarea" value="<%=rs2("soft_url4")%>"> 
                  </td>
                </tr>
                <tr class=zoneloveqs> 
                  <td colspan="2" align="center" height="30"> 
<input type="checkbox" name="isbest" value="1" id="isbest" <%if rs2("isbest")=1 then response.write "checked" end if%>>
              �Ƽ�&nbsp;&nbsp;&nbsp; 
                    <input name="Submit" type="submit" class="button" id="Submit" value="ȷ���޸�"> 
<input type="button" value="����"  onClick="location.href='admin_down.asp?action=list'" class="button"></td>
                </tr>
				<input type="hidden" name="id" value="<%=request.QueryString("id")%>">
                <input type="hidden" name="action" value="editsoft">
                <input type="hidden" name="MM_insert" value="true">
              </form>
            </table>
			<%end if
			if request.QueryString("action")="delsoft" then
			if request.querystring("id")="" then
Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
Response.Write("<script language=javascript>alert('�Ƿ������ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from soft where soft_id="&cint(request.querystring("id"))
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
			%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
              <form name="form1" method="post" action="admin_down.asp?action=list">
                <tr> 
                  <td colspan="2" class="diaryhead">ɾ�����</td>
                </tr>
                <tr class=zoneloveds> 
                  <td width="16%">�������</td>
                  <td width="84%"><%=rs2("soft_name")%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>�����С</td>
                  <td><%=rs2("soft_size")%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��Ȩ��ʽ</td>
                  <td><%=rs2("soft_mode")%> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>Ӧ��ƽ̨</td>
                  <td><%=rs2("soft_roof")%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>�Ƽ��̶�</td>
                  <td><img src="img/star<%=rs2("soft_commend")%>.gif"></td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td><%=rs2("soft_catname")%> | <%=rs2("soft_classname")%></td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��ʾ��ַ</td>
                  <td><%=rs2("soft_demo")%> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��ҳ��ַ</td>
                  <td><%=rs2("soft_home")%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>ͼƬ��ַ</td>
                  <td><%=rs2("soft_showpic")%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>������</td>
                  <td><%=ubb2html(formatStr(autourl(rs2("soft_desc"))), true, true)%>
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td><%=rs2("soft_url1")%> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>��������</td>
                  <td><%=rs2("soft_url2")%> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>Զ������1</td>
                  <td><%=rs2("soft_url3")%> 
                  </td>
                </tr>
                <tr class=zoneloveds> 
                  <td>Զ������2</td>
                  <td><%=rs2("soft_url4")%> 
                  </td>
                </tr>
                <tr class=zoneloveqs> 
                  <td colspan="2" align="center" height="30"> 
                    <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��">
<input type="button" value="����"  onClick="location.href='admin_down.asp?action=list'" class="button"> </td>
                </tr>
                <input type="hidden" name="id" value="<%=request.QueryString("id")%>">
                <input type="hidden" name="action" value="delsoft">
                <input type="hidden" name="MM_insert" value="true">
              </form>
            </table>
<%end if
end if
rs.close
set rs=nothing

%>