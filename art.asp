<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/inc.asp"-->
<!--#include file="inc/format.asp"-->
<%
start="����"
dim founderr
founderr=false
if request.querystring("cat_id")<>"" then
  if not isInteger(request.querystring("cat_id")) then
    founderr=true
    Response.Write "<script language=javascript>alert('�����Ƿ�');javascript:history.back();</script>"
  end if
end if
if request("page")<>"" then
  if not isInteger(request("page")) then
    founderr=true
    Response.Write "<script language=javascript>alert('�����Ƿ�');javascript:history.back();</script>"
  end if
end if
if request("keyword")<>"" then
  if instr(request("keyword"),"'")>0 then
    founderr=true
    Response.Write "<script language=javascript>alert('���������Ƿ�');javascript:history.back();</script>"
  end if
end if
call head()
call menu()
dim totalart,Currentpage,totalpages,i,j,colname
sql="select art_id,cat_id,art_title,art_date,art_count,art_content,isbest,review,reviewcount from art order by art_id DESC"
if request("cat_id")<>"" then
sql="select art_id,cat_id,art_title,art_date,art_count,art_content,isbest,review,reviewcount from art where cat_id="&request("cat_id")&" order by art_id DESC"
elseif request("keyword")<>"" then
sql="select art_id,cat_id,art_title,art_date,art_count,art_content,isbest,review,reviewcount from art where "&request("select")&" like '%"&request("keyword")&"%'order by art_id DESC"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center><TBODY><TR><TD class=l></TD><TD class=m height="25">&nbsp; 
<%
sql="select cat_id,cat_name from a_cat"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
do while not rs2.eof
if request("cat_id")=cstr(rs2("cat_id")) then
response.write " �����ڵ�λ�� �� <font class=""title""> "&rs2("cat_name")&"</font>&nbsp;"
end if
rs2.movenext
loop
if rs2.eof and rs2.bof then
response.write "��ǰû�з��࣡"
end if
%>
</TD><TD class=r></TD></TR></TBODY></TABLE>
<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder="0">
  <TBODY>
  <TR vAlign=top align=left>
    <TD>
<DIV class=mframe>
<SCRIPT type=text/javascript>zonelovetable("<%if request("cat_id")<> "" then%>�����๲��<%elseif request("select")<>"" then%>��������<%else%>��ǰ����<%end if%><span><%=rs.recordcount%></span>ƪ����");</SCRIPT>
<table width="100%" border="1" align="center" cellspacing="0" cellpadding="3" bgcolor="#FFFFFF" bordercolor="#f0f0f0" style="border-collapse: collapse" frame=rhs><%
if not rs.eof then
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
do while not rs.eof and i<articleperpage
%>
<tr><td width="*"><UL class=nl><LI class=zonelove><a href='showart.asp?cat_id=<%=rs("cat_id")%>&art_id=<%=rs("art_id")%>' target='_blank' Title='���±��⣺<%=rs("art_title")%>&#13&#10����ʱ�䣺<%=rs("art_date")%>&#13&#10�Ķ�������<%=rs("art_count")%>��'><%=rs("art_title")%></a></li></ul></td><TD width="10%" align=Center><%=rs("art_count")%></TD><td width="25%" align="center"><%=rs("art_date")%></td></tr>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then
%>
<tr align="center"><td align=middle height="60" colSpan=4><%if request("cat_id")<> "" then%>�÷�����ʱû������<%elseif request("keyword")<>"" then%>û���ҵ�����[<b><font color=red><%=request("keyword")%></font></b>]У԰���⣡<%else%>û���κ����£������Ա����̨��ӣ�<%end if%></td></tr>
<%end if
end if%>
<form name="form1" method="post" action="art.asp?select=<%=request("select")%>&keyword=<%=request("keyword")%>&cat_id=<%=request("cat_id")%>"><tr><td colspan="3"><TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0><TBODY>
<TR><TD align=middle width="35%" height=25><IMG height=14 src="img/so.gif" width=14 align=absMiddle> ��[<font color="#FF6666"><%=totalart%></font>]����Ʒ����[<font color="#FF6666"><%=totalpages%></font>]ҳ</TD><TD width="40%" align=middle><IMG height=11 src="img/lt.gif" width=11 align=absMiddle>
<%
if CurrentPage<2 then
response.write "<font color='999966'>��ҳ ��һҳ</font> "
else
response.write "<a href=art.asp?select="&request("select")&"&keyword="&request("keyword")&"&page=1&cat_id="&request.querystring("cat_id")&">��ҳ</a> "
response.write "<a href=art.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&CurrentPage-1&"&cat_id="&request.querystring("cat_id")&">��һҳ</a> "
end if
if totalpages-currentpage<1 then
response.write "<font color='999966'>��һҳ βҳ</font>"
else
response.write "<a href=art.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&CurrentPage+1&"&cat_id="&request.querystring("cat_id")
response.write ">��һҳ</a> <a href=art.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&totalpages&"&cat_id="&request.querystring("cat_id")&">βҳ</a>"
end if
%>
<IMG height=11 src="img/rt.gif" width=11 align=absMiddle></TD><TD align=middle width="25%"><select name="page" class="zonelove">
<%
i=1
for i=1 to totalpages
if i=currentpage then
%>
<option value=<%=i%> selected>��<%=i%>ҳ</option><%else%><option value=<%=i%>>��<%=i%>ҳ</option>
<%end if
next%>
</select><input type="submit" name="Submit2" value="ת��" onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff"> </TD></TR></TABLE></td></tr></form>
<tr align="center"><td height="1" colspan="3"></td></tr></table><br></TD><TD width=5></TD><TD width=190>
<DIV class=member>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT><UL class=nl>
<%
rs.close
set rs=nothing
set rs3=server.createobject("adodb.recordset")
if request.querystring("cat_id")<>"" then
sql="select top "&toparticlenum&" art_id,art_title,art_count,cat_id,art_date from art where cat_id="&request.querystring("cat_id")&" order by art_count DESC"
else
sql="select top "&toparticlenum&" art_id,art_title,art_count,cat_id,art_date from art order by art_count DESC"
end if
rs3.open sql,conn,1,1
do while not rs3.eof
%>
<LI><a href='showart.asp?cat_id=<%=rs3("cat_id")%>&art_id=<%=rs3("art_id")%>' target='_blank' Title='���±��⣺<%=rs3("art_title")%>&#13&#10����ʱ�䣺<%=rs3("art_date")%>&#13&#10�Ķ�������<%=rs3("art_count")%>��'><%=left(rs3("art_title"),14)%></a></li><%
   rs3.movenext
   loop
   if rs3.eof and rs3.bof then
   response.write "<div align=center><br>û���������<br><br></div>"
   end if
   rs3.close
   set rs3=nothing
   %>
</UL>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT><UL class=nl>
<%
set rs4=server.createobject("adodb.recordset")
sql="select top "&toparticlenum&" art_id,art_title,art_count,cat_id,art_date from art where isbest=1 order by art_id DESC"
rs4.open sql,conn,1,1
do while not rs4.eof
%>
<LI><a href='showart.asp?cat_id=<%=rs4("cat_id")%>&art_id=<%=rs4("art_id")%>' target='_blank' Title='���±��⣺<%=rs4("art_title")%>&#13&#10����ʱ�䣺<%=rs4("art_date")%>&#13&#10�Ķ�������<%=rs4("art_count")%>��'><%=left(rs4("art_title"),14)%></a></li><%
   rs4.movenext
   loop
   if rs4.eof and rs4.bof then
   response.write "<div align=center><br>û���������<br><br></div>"
   end if
   rs4.close
   set rs4=nothing
   %>
</UL>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("վ �� �� ��","mr");</SCRIPT>
<form name="form2" method="post" action="art.asp"><div align=center><input type="radio" name="select" value="art_title" checked>����&nbsp;<input type="radio" name="select" value="art_content">����&nbsp;<input type='radio' name='select' value='review'>����<br><input type='text' name='keyword'  size='15' value='�����ؼ���' maxlength='50' onFocus='this.select();' class="zonelove">&nbsp;<input type='submit' name='search'  value='����' onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff">
</div></form>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("�� վ �� ��","mr");</SCRIPT>
<div style="LINE-HEIGHT: 180%">&nbsp;&nbsp;��վ�������ݰ�Ȩ������У����,��δ����Уͬ�������,�κε�λ������Ͻ���Ϯ������,��������ܵ���У���̶ȵ�����!</div>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
</TD>
</TR></TBODY></TABLE>
<%
rs2.close
set rs2=nothing
call footer()%>
<body>
</body>
</html>