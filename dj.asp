<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/inc.asp"-->
<%
start="��������"
dim founderr
founderr=false
if request.querystring("djcat_id")<>"" then
  if not isInteger(request.querystring("djcat_id")) then
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
sql="select djcat_id,djcat_name from djcat"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
response.write "<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center><TBODY><TR><TD class=l></TD>" & vbCrLf
response.write "<TD height=""23"" align=""right"">"
do while not rs.eof
if request("djcat_id")=cstr(rs("djcat_id")) then
response.write "<font class=""title"">��"&rs("djcat_name")&"</font>&nbsp;"
else
response.write "<a href='?djcat_id="&rs("djcat_id")&"'>��"&rs("djcat_name")&"</a>&nbsp;"
end if
rs.movenext
loop
if rs.bof and rs.eof then
response.write "��ǰû�з���&nbsp;"
rs.close
end if%>
</TD><TD class=r></TD></TR></TBODY></TABLE>
<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder="0" height=300>
  <TBODY>
  <TR vAlign=top align=left>
    <TD width=200>
<DIV class=member>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","ml");</SCRIPT>
<UL class=nl>
<%
sql="SELECT top "&topdjnum&" dj_id,dj_name,dj_count,dj_date FROM dj ORDER by dj_count DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%>
<LI><a href='showdj.asp?dj_id=<%=rs("dj_id")%>' onclick="LyWindow(this.href,'dj','400','330','no','no','center');return false" Title='�������ƣ�<%=rs("dj_name")%>&#13&#10���������<%=rs("dj_count")%>&#13&#10�ϴ�ʱ�䣺<%=rs("dj_date")%>'><%=left(rs("dj_name"),14)%></a></LI>
<%rs.movenext
loop
if rs.eof and rs.bof then%><div align=center><br>��ǰ��û������<br><br></div>
<%
rs.close
end if%></UL>
<SCRIPT type=text/javascript>zonelovetablebottom("mr");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("FLASH �� ��","ml");</SCRIPT>
<UL class=nl>
<%
sql="SELECT top "&bestdj&" dj_id,dj_name,dj_count,dj_date FROM dj where isbest = 1 order by dj_id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%><LI><a href='showdj.asp?dj_id=<%=rs("dj_id")%>' onclick="LyWindow(this.href,'dj','400','330','no','no','center');return false" Title='�������ƣ�<%=rs("dj_name")%>&#13&#10����������<%=rs("dj_count")%>&#13&#10�ϴ�ʱ�䣺<%=rs("dj_date")%>'><%=left(rs("dj_name"),14)%></a></LI>
<%rs.movenext
loop
if rs.eof and rs.bof then%><div align=center><br>��ǰ��û������<br><br></div>
<%rs.close
end if%>
</UL>
<SCRIPT type=text/javascript>zonelovetablebottom("mr");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("FLASH���� �� �� ��","mr");</SCRIPT>
<form name="form2" method="post" action="dj.asp"><div align=center><input type='radio' name='select' value='dj_name' checked>����&nbsp;<input type='radio' name='select' value='dj_user'>����&nbsp;<input type='radio' name='select' value='review'>����<input type='text' name='keyword'  size='15' value='�����ؼ���' maxlength='50' onFocus='this.select();' class="zonelove">&nbsp;<input type='submit' name='search'  value='����' onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff">
</div></form>
<SCRIPT type=text/javascript>zonelovetablebottom("mr");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("�� վ �� ��","ml");</SCRIPT>
<div style="LINE-HEIGHT: 180%">&nbsp;1.��վFLASH������������������,��Ȩ��ԭ��������!���а�Ȩ���⾴��<a href='mailto:<%=webmail%>' Title='��վ��д��'><font color=green>��ϵ����</font></a>,лл!<br>&nbsp;2.��������ԭ���е��������ӽ����������ĵȴ���<br>&nbsp;3.�緢�ָ����������,����FLASH���Ŵ��ڰ�[<font color=red><b>B</b></font>]��&nbsp;���ʹ��󱨸�!</UL>
<SCRIPT type=text/javascript>zonelovetablebottom("mr");</SCRIPT>
<%
dim totalcs,Currentpage,totalpages,i
sql="select * from dj order by dj_id DESC"
if request.querystring("djcat_id")<>"" then
sql="select dj_id,dj_name,dj_count,dj_date,dj_user,review,reviewcount from dj where djcat_id="&request.querystring("djcat_id")&" order by dj_id DESC"
elseif request("keyword")<>"" then
sql="select dj_id,dj_name,dj_count,dj_date,dj_user,review,reviewcount from dj where "&request("select")&" like '%"&request("keyword")&"%'order by dj_id DESC"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
totalcs=rs.recordcount
%>
</DIV></TD>
<TD>
<DIV class=mframe>
<SCRIPT type=text/javascript>zonelovetable("<%if request("djcat_id")<> "" then%>�����๲��<%elseif request("keyword")<>"" then%>��������<%else%>��ǰ����<%end if%><span><%=totalcs%></span>������");</SCRIPT></SPAN>
<table width="100%" border="1" align="center" cellspacing="0" cellpadding="3" bgcolor="#FFFFFF" bordercolor="#f0f0f0" style="border-collapse: collapse" frame=lhs>
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
   if currentpage<>1 then
       if (currentpage-1)*djperpage<totalcs then
	       rs.move(currentpage-1)*djperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalcs mod djperpage)=0 then
      totalpages=totalcs\djperpage
   else
      totalpages=totalcs\djperpage+1
   end if
i=0
do while not rs.eof and i<djperpage
%>
<TR bgColor=#FFFFFF><TD nowrap width="60%"><UL class=nl><LI class=zonelove><a href='showdj.asp?dj_id=<%=rs("dj_id")%>' onclick="LyWindow(this.href,'dj','400','330','no','no','center');return false" Title='�������ƣ�<%=rs("dj_name")%>&#13&#10����������<%=rs("dj_count")%>&#13&#10�ϴ�ʱ�䣺<%=rs("dj_date")%>'><%=rs("dj_name")%></a></LI></UL></TD><TD width="15%" align=Center><a href="dj.asp?select=dj_user&keyword=<%=rs("dj_user")%>" title='��������[<%=rs("dj_user")%>]�����и���'><%=rs("dj_user")%></a></TD><TD width="10%" align=Center><%=rs("dj_count")%></TD><TD width="15%" align=Center><%=rs("dj_date")%></TD></TR>
<%
i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then%><tr><td align=middle height="60" colSpan=4><%if request("djcat_id")<> "" then%>�÷�����ʱû������<%elseif request("keyword")<>"" then%>û���ҵ�����[<b><font color=red><%=request("keyword")%></font></b>]�����֣�<%else%>û���κ����֣������Ա����̨��ӣ�<%end if%></td></td></tr>
<%end if
end if
%>
<form name="form1" method="post" action="dj.asp?select=<%=request("select")%>&keyword=<%=request("keyword")%>&djcat_id=<%=request.querystring("djcat_id")%>">
<tr> 
<td align="right" colspan="4">
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
<TBODY>
<TR>
<TD align=middle width="200"><IMG height=14 src="img/so.gif" width=14 align=absMiddle> ��[<font color="#FF6666"><%=totalcs%></font>]����Ʒ����[<font color="#FF6666"><%=totalpages%></font>]ҳ</TD>
<TD width="200" align=middle><IMG height=11 src="img/lt.gif" width=11 align=absMiddle>
<%
if CurrentPage<2 then
response.write "<font color='999966'>��ҳ ��һҳ</font> "
else
response.write "<a href=dj.asp?select="&request("select")&"&keyword="&request("keyword")&"&page=1&djcat_id="&request.querystring("djcat_id")&">��ҳ</a> "
response.write "<a href=dj.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&CurrentPage-1&"&djcat_id="&request.querystring("djcat_id")&">��һҳ</a> "
end if
if totalpages-currentpage<1 then
response.write "<font color='999966'>��һҳ βҳ</font>"
else
response.write "<a href=dj.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&CurrentPage+1&"&djcat_id="&request.querystring("djcat_id")
response.write ">��һҳ</a> <a href=dj.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&totalpages&"&djcat_id="&request.querystring("djcat_id")&">βҳ</a>"
end if
%>&nbsp;<IMG height=11 src="img/rt.gif" width=11 align=absMiddle></TD>
<TD align=middle width="120">
<select name="page" class="zonelove">
<%
i=1
for i=1 to totalpages
if i=currentpage then
%>
<option value=<%=i%> selected>��<%=i%>ҳ</option> 
<%else%>
<option value=<%=i%>>��<%=i%>ҳ</option> 
<%end if
next%>
</select><input type="submit" name="Submit2" value="ת��" onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff"> </TD>
</TR>
</TABLE></td></tr></FORM>
<tr align="center"><td height="1" colspan="3"></td></tr></table>
</DIV><br></TD></TR>
</TBODY></TABLE>
<%
rs.close
set rs=nothing
call footer()
%>
</body>
</html>