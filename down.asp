<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/inc.asp"-->
<%
start="��������"
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
dim keyword,class_id,colname,totalsoft,Currentpage,totalpages,i
keyword=request("keyword")
cat_id=request("cat_id")
class_id=request("class_id")
colname=request("colname")
sql="SELECT soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount FROM soft order by soft_id desc"
if cat_id<>"" and class_id="" then
sql="SELECT soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount FROM soft where soft_catid="&cat_id&" order by soft_id desc"
elseif class_id<>"" and keyword="" then
sql="SELECT soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount FROM soft where soft_classid="&class_id&" order by soft_id desc"
elseif keyword<>"" and colname<>"0" then
sql="select soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount from soft where "&colname&" like '%"&keyword&"%' order by soft_id desc"
elseif keyword<>"" and colname="0" then
sql="select soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount from soft where soft_name like '%"&keyword&"%' or soft_desc like '%"&keyword&"%' order by soft_id desc"
else
sql="SELECT soft_id,soft_name,soft_dcount,soft_rcount,soft_demo,soft_mode,soft_desc,isbest,soft_size,soft_commend,review,reviewcount FROM soft order by soft_id desc"
end if
set rssoft=server.createobject("adodb.recordset")
rssoft.open sql,conn,1,1
response.write "<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center><TBODY><TR><TD class=m></TD>" & vbCrLf
response.write "<TD height=""23"">"
Set rscat = Server.CreateObject("ADODB.Recordset")
sqlcat="SELECT cat_id,cat_name FROM d_cat"
rscat.OPEN sqlcat,Conn,1,1
do while not rscat.eof
if request("cat_id")=cstr(rscat("cat_id")) then
response.write "&nbsp;<font class=""title"">��"&rscat("cat_name")&"</font>"
else
response.write "&nbsp;<a href='?cat_id="&rscat("cat_id")&"'>��"&rscat("cat_name")&"</a>&nbsp;"
end if
rscat.movenext
loop
rscat.close
set rscat=nothing

%>
</TD><TD class=r></TD></TR></TBODY></TABLE>
<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder="0">
  <TBODY>
  <TR vAlign=top align=left>
    <TD>
<DIV class=mframe>
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
<SCRIPT type=text/javascript>zonelovetable("<%=flname%>����<%if keyword<>"" then%>��������<%else%>����<%end if%><%=rssoft.recordcount%>���ļ�");</SCRIPT>
<table width="100%" border="1" align="center" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#f0f0f0" style="border-collapse: collapse" frame=rhs><%
if not rssoft.eof then
rssoft.movefirst
rssoft.pagesize=softperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rssoft.pagecount then
      currentpage=rssoft.pagecount
   end if
else
   currentpage=1
end if
   totalsoft=rssoft.recordcount
   if currentpage<>1 then
       if (currentpage-1)*softperpage<totalsoft then
	       rssoft.move(currentpage-1)*softperpage
		   dim bookmark
		   bookmark=rssoft.bookmark
	   end if
   end if
   if (totalsoft mod softperpage)=0 then
      totalpages=totalsoft\softperpage
   else
      totalpages=totalsoft\softperpage+1
   end if
%>
<%
i=0
do while not rssoft.eof and i<softperpage
%>
<tr bgcolor="#ffffff"><td colspan="5" height="10"></td></tr><tr bgcolor="#ffffff"><td colspan="5" height="10">
<table width="100%" border="0"><tr><td width="*"><UL class=nl><LI class=zonelove><a href='showdown.asp?soft_id=<%=rssoft("soft_id")%>' target='_blank' Title='�鿴��ϸ����'><font class='view'><%=rssoft("soft_name")%></font></a></li></ul></td><td width="50"><%if rssoft("soft_demo")<>"" then%><a href="<%=rssoft("soft_demo")%>" target="_blank" title="�鿴��ʾ">
<font color="#009900">��ʾ</font><%end if%></td>
</tr></table></td></tr>
<tr bgcolor="#ffffff"><td colspan="5" height="30">��<%=cutstr(rssoft("soft_desc"),40,false,none)%></td></tr>
 <tr bgcolor="#ffffff"> 
<td width="20%" height="25">&nbsp;�����С��<font color="#000000"><%=rssoft("soft_size")%></font></td>
<td width="25%">&nbsp;��Ȩ��ʽ��<%=rssoft("soft_mode")%></td>
<td width="25%">&nbsp;���أ�<%=rssoft("soft_dcount")%> �����<%=rssoft("soft_rcount")%></td>
<td width="30%">&nbsp;�Ƽ��̶ȣ�<img src="img/star<%=rssoft("soft_commend")%>.gif" align="absmiddle"></td>
        </tr>
      <%
i=i+1
rssoft.movenext
loop
else
if rssoft.eof and rssoft.bof then
%>
        <tr> 
          <td height="100" colspan="5" bgcolor="#Ffffff" align="center"> 
            <%if keyword<>"" then%>û���ҵ�����[<b><font color=red><%=request("keyword")%></font></b>]���ļ���
            <%else%>û���κ��ļ��ṩ���أ������Ա����̨��ӣ�
            <%end if%> 
          </td>
        </tr>
<%end if
end if
%>
<tr bgcolor="#ffffff"><td colspan="5" height="10"></td></tr><tr bgcolor="#ffffff"><td colspan="5" height="30">
<TABLE width="100%" align=center border=0><TBODY>
<form name="form1" method="post" action="down.asp?colname=<%=request("colname")%>&keyword=<%=request("keyword")%>&cat_id=<%=request("cat_id")%>&class_id=<%=request("class_id")%>">
<TR> 
<TD align=middle width="35%" height=25><IMG height=14 src="img/so.gif" width=14 align=absMiddle> ��[<font color="#FF6666"><%=totalsoft%></font>]����Ʒ����[<font color="#FF6666"><%=totalpages%></font>]ҳ</TD>
<TD width="40%" align=middle><IMG height=11 src="img/lt.gif" width=11 align=absMiddle>
<%
if CurrentPage<2 then
response.write "<font color='999966'>��ҳ ��һҳ</font> "
else
response.write "<a href=down.asp?colname="&request("colname")&"&keyword="&request("keyword")&"&page=1&cat_id="&request.querystring("djcat_id")&"&class_id="&request("class_id")&">��ҳ</a> "
response.write "<a href=down.asp?colname="&request("colname")&"&keyword="&request("keyword")&"&page="&CurrentPage-1&"&cat_id="&request.querystring("djcat_id")&"&class_id="&request("class_id")&">��һҳ</a> "
end if
if totalpages-currentpage<1 then
response.write "<font color='999966'>��һҳ βҳ</font>"
else
response.write "<a href=down.asp?colname="&request("colname")&"&keyword="&request("keyword")&"&page="&CurrentPage+1&"&cat_id="&request.querystring("cat_id")&"&class_id="&request("class_id")
response.write ">��һҳ</a> <a href=down.asp?colname="&request("colname")&"&keyword="&request("keyword")&"&page="&totalpages&"&cat_id="&request.querystring("cat_id")&"&class_id="&request("class_id")&">βҳ</a>"
end if
%>
 <IMG height=11 src="img/rt.gif" width=11 align=absMiddle></TD>
<TD align=middle width="25%">
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
</select><input type="submit" name="Submit2" value="ת��" onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff"></TD>
</TR>
</FORM>
</TABLE>          
</TD>
</TR></FORM>
<tr align="center"><td bgcolor="#F0f0f0" height="1" colspan="3"></td></tr>
</TABLE><br>  
</TD><TD width=5></TD><TD width=190>
<DIV class=member>
<%if not cat_id = "" then%>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT>
<UL class=nl><%Set rsclass = Server.CreateObject("ADODB.Recordset")
sqlclass="SELECT cat_id,class_id,class_name FROM d_class where cat_id=" &cat_id& ""
rsclass.OPEN sqlclass,Conn,1,1
	do while not rsclass.eof
if request("class_id")=cstr(rsclass("class_id")) then
response.write "<LI><font class=""title"">"&rsclass("class_name")&"</font></LI>"
else
response.write "<LI><a href='?cat_id="&rsclass("cat_id")&"&class_id="&rsclass("class_id")&"'>"&rsclass("class_name")&"</a></LI>"
end if
		rsclass.movenext
	loop
	rsclass.close
set rsclass=nothing
%></UL>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<%end if
sql1="SELECT top " &topsoftnum& " soft_dcount,soft_rcount,soft_id,soft_name,soft_joindate FROM soft where soft_id ORDER BY soft_dcount DESC , soft_id DESC"
sql2="SELECT top " &newsoft& " soft_id,soft_name,soft_joindate,soft_dcount,isbest,soft_rcount FROM soft where isbest=1 order by soft_id DESC"
set rs1=server.createobject("adodb.recordset")
rs1.open sql1,conn,1,1
set rs2=server.createobject("adodb.recordset")
rs2.open sql2,conn,1,1
if request.querystring("class_id")<>"" or request.querystring("cat_id")<>"" then
if request.querystring("class_id")="" and request.querystring("cat_id")<>"" then
sql3="SELECT top 10 soft_id,soft_name,soft_joindate,soft_dcount,soft_rcount FROM soft WHERE soft_catid= " &request.querystring("cat_id")& " and soft_id order by soft_dcount DESC , soft_id DESC"
elseif request.querystring("class_id")<>"" then
sql3="SELECT top 10 soft_id,soft_name,soft_joindate,soft_dcount,soft_rcount FROM soft WHERE soft_classid= " &request.querystring("class_id")& " and soft_id order by soft_dcount DESC , soft_joindate DESC"
end if
set rs3=server.createobject("adodb.recordset")
rs3.open sql3,conn,1,1
%>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT><UL class=nl>
<%do while not rs3.eof%>
<LI><a href='showdown.asp?soft_id=<%=rs3("soft_id")%>' target='_blank' Title='�ļ����ƣ�<%=rs3("soft_name")%>&#13&#10�ϴ�ʱ�䣺<%=rs3("soft_joindate")%>&#13&#10���ش�����<%=rs3("soft_dcount")%>'><%=left(rs3("soft_name"),14)%></a></li>
              <%rs3.movenext
loop
if rs3.eof and rs3.bof then%>
<div align=center><br>Ŀǰ��û���ļ���<br><br></div>
              <%end if%>
</UL><SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
      <%
rs3.close
set rs3=nothing
end if
%>
<SCRIPT type=text/javascript>zonelovetabletop("ȫ �� �� ��","mr");</SCRIPT><UL class=nl>
<%do while not rs1.eof%>
<LI><a href='showdown.asp?soft_id=<%=rs1("soft_id")%>' target='_blank' Title='�ļ����ƣ�<%=rs1("soft_name")%>&#13&#10�ϴ�ʱ�䣺<%=rs1("soft_joindate")%>&#13&#10���ش�����<%=rs1("soft_dcount")%>'><%=left(rs1("soft_name"),14)%></a>
              <%rs1.movenext
loop
if rs1.eof and rs1.bof then%>
<div align=center><br>Ŀǰ��û���ļ���<br><br></div>
              <%end if%>
</UL><SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT><UL class=nl>
              <%do while not rs2.eof%>
<LI><a href='showdown.asp?soft_id=<%=rs2("soft_id")%>' target='_blank' Title='�ļ����ƣ�<%=rs2("soft_name")%>&#13&#10�ϴ�ʱ�䣺<%=rs2("soft_joindate")%>&#13&#10���ش�����<%=rs2("soft_dcount")%>'><%=left(rs2("soft_name"),14)%></a></li>
              <%rs2.movenext
loop
if rs2.eof and rs2.bof then%>
<div align=center><br>Ŀǰ��û���ļ���<br><br></div>
             <%end if%>
          </td>
        </tr>
      <%
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
%>

</UL><SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
<SCRIPT type=text/javascript>zonelovetabletop("�� �� �� ��","mr");</SCRIPT><center>
<form name="form1" method="post" action="down.asp"> 
<input type="radio" name="colname" value="soft_name" checked>����&nbsp;<input type="radio" name="colname" value="soft_desc">���&nbsp;<input type='radio' name='colname' value='review'>����
<input type='text' name='keyword'  size='15' value='�����ؼ���' maxlength='50' onFocus='this.select();' class="zonelove">&nbsp;<input type='submit' name='Submit'  value='����' onmouseover="this.className='boton'" onmouseout="this.className='botoff'" class="botoff"></form>
<SCRIPT type=text/javascript>zonelovetablebottom("ml");</SCRIPT>
</TD>
</TR></TBODY></TABLE>
<%
rssoft.close
set rssoft=nothing
call footer()
%>
<body>
</body>
</html>