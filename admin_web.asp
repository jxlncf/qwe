<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('����δ��¼�����߳�ʱ�ˣ������µ�¼');this.top.location.href='admin.asp';</script>")
  response.end
else
if request.form("MM_insert") then
if request.form("action")="newcscat" then
sql="select * from cscat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim cscatname
cscatname=trim(replace(request.form("cscat_name"),"'",""))
if cscatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
else
  rs("cscat_name")=cscatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_web.asp?action=cscat"
end if
elseif request.form("action")="newcs" then
sql="select * from coolsites"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim cscatid,csname,csurl,cspic,csdesc,csisbest
cscatid=cint(request.form("cscatid"))
csname=trim(replace(request.form("name"),"'",""))
csurl=trim(replace(request.form("url"),"'",""))
cspic=trim(replace(request.form("pic"),"'",""))
csdesc=trim(replace(request.form("desc"),"'",""))
csisbest=request.form("isbest")
if csname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д���ƣ�');history.back(1);</script>")
else
  rs("cs_name")=csname
end if
if cscatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ѡ�������������࣡');history.back(1);</script>")
else
  rs("cscat_id")=cscatid
end if
if csurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д���ʱ�䣡');history.back(1);</script>")
else
  rs("cs_url")=csurl
end if
if cspic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ��ͼƬ��ַ��');history.back(1);</script>")
else
  rs("cs_pic")=cspic
end if
if csdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ�ļ�Ҫ˵��(���)��');history.back(1);</script>")
else
  rs("cs_desc")=csdesc
end if
if cint(csisbest)=1 then
  rs("isbest")=cint(csisbest)
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  sql="update allcount set coolsitescount = coolsitescount + 1"
  conn.execute(sql)
  
  response.redirect "admin_web.asp?action=sites"
end if
elseif request.form("action")="editcs" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦid������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from coolsites where cs_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cscatid=cint(request.form("cscatid"))
csname=trim(replace(request.form("name"),"'",""))
csurl=trim(replace(request.form("url"),"'",""))
cspic=trim(replace(request.form("pic"),"'",""))
csdesc=trim(replace(request.form("desc"),"'",""))
csisbest=request.form("isbest")
if csname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ������');history.back(1);</script>")
else
  rs("cs_name")=csname
end if
if cscatid<1 then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ѡ���ʦ�������࣡');history.back(1);</script>")
else
  rs("cscat_id")=cscatid
end if
if csurl="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ���ʱ�䣡');history.back(1);</script>")
else
  rs("cs_url")=csurl
end if
if cspic="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ��ͼƬ��ַ��');history.back(1);</script>")
else
  rs("cs_pic")=cspic
end if
if csdesc="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д��ʦ�ļ�Ҫ˵��(���)��');history.back(1);</script>")
else
  rs("cs_desc")=csdesc
end if
rs("isbest")=cint(csisbest)

if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_web.asp?action=sites"
end if
elseif request.form("action")="delcs" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦid������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from coolsites where cs_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  sql="update allcount set coolsitescount = coolsitescount - 1"
  conn.execute(sql)
  
  response.redirect "admin_web.asp?action=sites"
elseif request.form("action")="editcscat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦ����id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from cscat where cscat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cscatname=trim(replace(request.form("cscat_name"),"'",""))
if cscatname="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�������д�������ƣ�');history.back(1);</script>")
else
  rs("cscat_name")=cscatname
end if
if founderr then
  call diserror()
  response.end
else
  rs.update
  rs.close
  set rs=nothing
  response.redirect "admin_web.asp?action=cscat"
end if
elseif request.form("action")="delcscat" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('�����ָ�������Ķ���');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦ����id������');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from cscat where cscat_id="&cint(request.form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
  rs.delete
  rs.close
  set rs=nothing
  response.redirect "admin_web.asp?action=cscat"
end if
end if%>
<HTML><HEAD><TITLE>��������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
<%
if request.querystring("action")="cscat" then
sql="select * from cscat order by cscat_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <tr class=zonelovess> 
          <td colspan="3">��ʦ��ɷ������</font></td>
        </tr>
        <tr class=zoneloveqs align="center"> 
          <td width="10%">���</td>
          <td width="70%">��������</td>
          <td width="20%">����</td>
        </tr>
        <%do while not rs.eof%>
        <tr class=zoneloveds> 
          <td align="center"><%=rs("cscat_id")%>��</td>
          <td><a href="#"><%=rs("cscat_name")%></a>��</td>
          <td align="center"><a href="admin_web.asp?id=<%=rs("cscat_id")%>&action=editcscat">�༭</a> 
            <a href="admin_web.asp?id=<%=rs("cscat_id")%>&action=delcscat">ɾ��</a> <a href="showcs.asp?cscat_id=<%=rs("cscat_id")%>" target="_blank">�鿴</a></td>
        </tr>
        <%rs.movenext
loop
if rs.bof and rs.eof then%>
        <tr class=zoneloveds align="center"> 
          <td colspan="3">��ǰû�з��࣡</td>
        </tr>
        <%rs.close
set rs=nothing
		end if%>
      </table>
     <%end if
	 if request.querystring("action")="newcscat" then%> 
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
		<tr class=zonelovess> 
            <td>������ʦ����</font></td>
        </tr>
        <tr class=zoneloveds> 
            <td>��������- 
              <input type="text" name="cscat_name" size="40" class="textarea">
          </td>
        </tr>
        <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
        </tr>
		<input type="hidden" name="action" value="newcscat">
		<input type="hidden" name="MM_insert" value="true">
		</form>
      </table>
	  <%end if
	  if request.QueryString("action")="editcscat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦ����ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from cscat where cscat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>�޸Ľ�ʦ����</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>��������- 
              <input name="cscat_name" type="text" class="textarea" id="cscat_name" size="40" value="<%=rs("cscat_name")%>"> 
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> <input name="Submit" type="submit" class="button" id="Submit" value="ȷ���޸�"> 
              <input name="Reset" type="reset" class="button" id="Reset" value="�������"> </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("cscat_id")%>">
          <input type="hidden" name="action" value="editcscat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="delcscat" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ��Ľ�ʦ����ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from cscat where cscat_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	  %>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr class=zonelovess> 
            <td>ɾ����ʦ����</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>��������- <%=rs("cscat_name")%>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��">
            </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("cscat_id")%>">
          <input type="hidden" name="action" value="delcscat">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
if request.querystring("action")="sites" then
sql="select * from coolsites order by cs_id desc"
if request.querystring("cscat_id")<>"" then
sql="select * from coolsites where cscat_id="&cint(request.querystring("cscat_id"))&" order by cs_id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
totalcs=rs.recordcount
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
	    <form name="form3" method="post" action="">
        <tr class=zonelovess>
          <td align="right" colspan="2" style="padding-top:2px;">
              <select style="margin:-3px" name="go" onChange='window.location=form.go.options[form.go.selectedIndex].value'>
			  <option value="">ѡ����ʾ��ʽ</option>
			  <option value="admin_web.asp?action=sites">��ʾ���н�ʦ</option>
<%sql="select * from cscat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
<option value="admin_web.asp?action=sites&cscat_id=<%=rs2("cscat_id")%>"><%=rs2("cscat_name")%></option>
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
<%
if not rs.eof then
rs.movefirst
rs.pagesize=coolsitesperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalcs=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*coolsitesperpage<totalcs then
	       rs.move(currentpage-1)*coolsitesperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalcs mod coolsitesperpage)=0 then
      totalpages=totalcs\coolsitesperpage
   else
      totalpages=totalcs\coolsitesperpage+1
   end if
i=0
do while not rs.eof and i<coolsitesperpage
%>
        <tr class=zoneloveds> 
          <td width="37%" align="center"><img src="<%=rs("cs_pic")%>" width='150' border='0'>
          </td>
          <td width="63%">��ʦ������<%=rs("cs_name")%><br>
            ���ʱ�䣺<%=rs("cs_date")%><br>
            <a href="admin_web.asp?id=<%=rs("cs_id")%>&action=editcs">�༭</a> 
            <a href="admin_web.asp?id=<%=rs("cs_id")%>&action=delcs">ɾ��</a></td>
        </tr>
      <%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
        <tr class=zoneloveds> 
          <td colspan="2" height="22" align="center">��ǰû���κν�ʦ��Ϣ��</td>
        </tr>
      <%end if
end if%>
        <form name="form1" method="post" action="admin_web.asp?action=sites&cscat_id=<%=request.querystring("cscat_id")%>">
          <tr class=zoneloveqs> 
            <td align="right" colspan="2"> <%=currentpage%> /<%=totalpages%>ҳ,<%=totalcs%>����¼/<%=coolsitesperpage%>ƪÿҳ. 
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
              <a href="admin_web.asp?action=sites&page=<%=i%>&cscat_id=<%=request.querystring("cscat_id")%>"><%=i%></a> 
              <%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_web.asp?action=sites&page=<%=page%>&cscat_id=<%=request.querystring("cscat_id")%>" title="��һҳ">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button">
            </td>
          </tr>
        </form>
      </table>
	  <%end if
	  if request.querystring("action")="newcs" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">������ʦ</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ����</td>
            <td width="83%"> 
              <input type="text" name="name" class="textarea" size="40">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��������</td>
            <td width="83%"> 
              <select name="cscatid" class="textarea">
<%
sql="select * from cscat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof%><option value="<%=rs("cscat_id")%>"><%=rs("cscat_name")%></option>
<%rs.movenext
loop
if rs.eof and rs.bof then%><option value="0">��ǰû�н�ʦ����</option>
<%end if%>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">�������:</td>
            <td width="83%"> 
              <input type="text" name="url" class="textarea" size="60">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">ͼƬ��ַ:</td>
            <td width="83%"> 
              <input type="text" name="pic" class="textarea" size="60">
            </td>
          </tr>
             <tr class=zoneloveds> 
              <td> �ϴ�ͼƬ�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=web"></iframe></td>
            </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ����</td>
            <td width="83%"> 
              <textarea name="desc" class="textarea" cols="65" rows="6"></textarea>
            </td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input type="checkbox" name="isbest" value="1" class="textarea">
              �Ƽ�&nbsp;&nbsp;&nbsp; 
              <input type="submit" name="Submit" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_web.asp?action=sites">����</a>] </td>
          </tr>
		  <input type="hidden" name="action" value="newcs">
		<input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
      <%end if
	  if request.QueryString("action")="editcs" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ���ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from coolsites where cs_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">�޸Ľ�ʦ��Ϣ</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ����</td>
            <td width="83%"> 
              <input name="name" type="text" class="textarea" id="name" size="40" value="<%=rs("cs_name")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��������</td>
            <td width="83%"> 
              <select name="cscatid" class="textarea" id="cscatid">
                <%
sql="select * from cscat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof%>
                <option value="<%=rs2("cscat_id")%>" <%if rs2("cscat_id")=rs("cscat_id") then response.write "selected" end if%>><%=rs2("cscat_name")%></option>
                <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
                <option value="0">��ǰû�з���</option>
                <%end if%>
              </select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��������</td>
            <td width="83%"> 
              <input name="url" type="text" class="textarea" id="url" size="60" value="<%=rs("cs_url")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">ͼƬ��ַ</td>
            <td width="83%"> 
              <input name="pic" type="text" class="textarea" id="pic" size="60" value="<%=rs("cs_pic")%>">
            </td>
          </tr>
             <tr class=zoneloveds> 
              <td> �ϴ�ͼƬ�� </td>
              <td align="left" height="25"><iframe frameborder=0 width=290 height=25 scrolling=no src="upload.asp?action=web"></iframe></td>
            </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ���</td>
            <td width="83%"> 
              <textarea name="desc" cols="65" rows="6" class="textarea" id="desc"><%=rs("cs_desc")%></textarea>
            </td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2" height="30" bgcolor="#F5F5F5"> <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              �Ƽ�
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ���޸�"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_web.asp?action=sites">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("cs_id")%>">
          <input type="hidden" name="action" value="editcs">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
	  if request.QueryString("action")="delcs" then
	  if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('��ָ�������Ķ���');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('�Ƿ���У������ID������');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from coolsites where cs_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
 <table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form2" method="post" action="">
          <tr class=zonelovess> 
            <td colspan="2">ɾ����ʦ��Ϣ</font></td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ����</td>
            <td width="83%"> <%=rs("cs_name")%> ��</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��������</td>
            <td width="83%"> 
                <%
sql="select * from cscat where cscat_id="&rs("cscat_id")
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
%><%=rs2("cscat_name")%> ��</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">�������</td>
            <td width="83%"><%=rs("cs_url")%> ��</td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">ͼƬ��ַ</td>
            <td width="83%"><img src="<%=rs("cs_pic")%>">
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td width="17%">��ʦ���</td>
            <td width="83%"><%=rs("cs_desc")%> ��</td>
          </tr>
          <tr class=zoneloveqs align="center"> 
            <td colspan="2"> 
              <input name="isbest" type="checkbox" class="textarea" id="isbest" value="1" <%if rs("isbest")=1 then response.write "checked" end if%>>
              �Ƽ� 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��">
              [<a href="admin_web.asp?action=sites">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("cs_id")%>">
          <input type="hidden" name="action" value="delcs">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
end if
%>