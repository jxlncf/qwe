<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('����δ��¼�����߳�ʱ�ˣ������µ�¼');this.top.location.href='admin.asp';</script>")
  response.end
else
if Request.form("MM_insert") then
if request.form("action")="newdiary" then
sql="select * from diary"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim d_content
d_content=rtrim(request.form("textfield"))
d_content=trim(replace(d_content," ","&nbsp;"))
if d_content="" then
  founderr=true
  Response.Write("<script language=javascript>alert('��������빫�����ݣ�');history.back(1);</script>")
else
  rs("d_content")=d_content
end if
if not founderr then
rs.update
rs.close
set rs=nothing
conn.execute(sql)
sql="update allcount set diarycount = diarycount + 1"
response.redirect "admin_diary.asp"
else
call diserror()
response.end
end if
end if
if request.form("action")="editdiary" then
 if request.form("id")="" then
   founderr=true
   Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
 else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ���id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from diary where d_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
d_content=rtrim(request.form("textfield"))
d_content=trim(replace(d_content," ","&nbsp;"))
if d_content="" then
  founderr=true
  Response.Write("<script language=javascript>alert('��������빫�����ݣ�');history.back(1);</script>")
else
  rs("d_content")=d_content
end if
if not founderr then
rs.update
rs.close
set rs=nothing
response.redirect "admin_diary.asp"
else
call diserror()
response.end
end if
end if
if request.form("action")="deldiary" then
 if request.form("id")="" then
   founderr=true
   Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
 else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ���id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from diary where d_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not founderr then
rs.delete
rs.close
set rs=nothing
conn.execute(sql)
sql="update allcount set diarycount = diarycount - 1" 
response.redirect "admin_diary.asp"
else
call diserror()
response.end
end if
end if
end if
dim totaldiary,adperpage,Currentpage,totalpages,i
adperpage=15
sql="select * from diary order by d_id DESC"
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
        <tr class=zonelovess> 
          <td colspan="4">�������</td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="50">���</td>
          <td width="*">����</td>
          <td width="150">ʱ��</td>
          <td width="100">����</td>
        </tr>
        <%if not rs.eof then
rs.Movefirst
rs.pagesize=adperpage
if trim(request("page"))<>"" then
    currentpage=clng(request("page"))
	if currentpage>rs.pagecount then
	   currentpage=rs.pagecount
	end if
else
    currentpage=1
end if
    totaldiary=rs.recordcount
	if currentpage<>1 then
	   if(currentpage-1)*adperpage<totaldiary then
	      rs.move(currentpage-1)*adperpage
		  dim bookmark
		  bookmark=rs.bookmark
	   end if
	end if
	if (totaldiary mod adperpage)=0 then
	   totalpages=totaldiary\adperpage
	else
	   totalpages=totaldiary\adperpage+1
	end if
i=0
do while not rs.eof and i<adperpage
%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("d_id")%>��</td>
          <td><%=cutstr(formatstr(rs("d_content")),20,false,none)%></td>
<td align="center"><%=rs("d_date")%></td>
          <td align="center"><a href="admin_diary.asp?id=<%=rs("d_id")%>&action=editdiary#editdiary">�༭</a> 
            <a href="admin_diary.asp?id=<%=rs("d_id")%>&action=deldiary#deldiary">ɾ��</a></td>
        </tr>
        <%i=i+1
rs.movenext
loop
else
If rs.EOF And rs.BOF Then%>
        <tr class=zoneloveds> 
          <td colspan=4 align="center">��ǰ��û�й��棡</td>
        </tr>
        <%end if
end if%>
        <form name="form2" method="post" action="admin_diary.asp">
          <tr class=zoneloveqs> 
            <td colspan=4 align="right"><%=currentpage%>/<%=totalpages%>ҳ,<%=totaldiary%>����¼/<%=adperpage%>ƪÿҳ. 
              <%
i=1
dy10=false
showye=totalpages
if showye>10 then
dy10=true
showye=10
end if
for i=1 to showye
if i=currentpage then
%>
              <%=i%> 
              <%else%>
              <a href="admin_diary.asp?page=<%=i%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_diary.asp?page=<%=page%>" title="��һҳ">>></a>
              <%end if%>
            </td>
          </tr>
        </form>
      </table>
      <%end if
if request.querystring("action")="newdiary" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td>�µĹ���</td>
          </tr>
          <tr class=zoneloveds> 
            <td> <input type="text" name="textfield" size="80" class="textarea">
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center" height="30"> 
              <input type="submit" name="Submit" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�����д" class="button">
            </td>
          </tr>
          <input type="hidden" name="action" value="newdiary">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="editdiary" then
if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ���ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from diary where d_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td>�޸Ĺ���</td>
          </tr>
          <tr class=zoneloveds> 
            <td> <input type="text" name="textfield" size="80" class="textarea" value="<%=rs("d_content")%>">
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center" height="30"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button">
              <input type="reset" name="Reset" value="�����д" class="button">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("d_id")%>">
          <input type="hidden" name="action" value="editdiary">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="deldiary" then
if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
   Response.Write("<script language=javascript>alert('�Ƿ���ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from diary where d_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td>ɾ������</td>
          </tr>
          <tr class=zoneloveds> 
            <td><%=ubb2html(formatStr(autourl(rs("d_content"))), true, true)%>
			<div align="right">������-<span class="newshead"><%=rs("d_date")%></span></div></td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="ȷ��ɾ��" class="button">
              [<a href="admin_diary.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("d_id")%>">
          <input type="hidden" name="action" value="deldiary">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
rs.close
set rs=nothing
end if
%>