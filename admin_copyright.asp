<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('����δ��¼�����߳�ʱ�ˣ������µ�¼');this.location.href='admin.asp';</script>")
  response.end
else
sql="select gy,lx,gg,dt from admin"
set rs=conn.execute(sql)
gy=rs("gy")
lx=rs("lx")
gg=rs("gg")
dt=rs("dt")
if request.form("MM_insert") then
gy=request.form("gy")
lx=request.form("lx")
gg=request.form("gg")
dt=request.form("dt")
if gy="" then
  rs("gy")=""
else
  rs("gy")=gy
end if
if lx="" then
  rs("lx")=""
else
  rs("lx")=lx
end if
if gg="" then
  rs("gg")=""
else
  rs("gg")=gg
end if
if dt="" then
  rs("dt")=""
else
  rs("dt")=dt
end if
conn.execute("update [admin] set gy='"&(Request("gy"))&"'")
conn.execute("update [admin] set lx='"&(Request("lx"))&"'")
conn.execute("update [admin] set gg='"&(Request("gg"))&"'")
conn.execute("update [admin] set dt='"&(Request("dt"))&"'")
  Response.Write("<script language=javascript>alert('���óɹ���');</script>")
end if%><HTML><HEAD><META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK href="inc/admin.css" type=text/css rel=stylesheet><META content="MSHTML 6.00.2800.1126" name=GENERATOR></head><body>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse"><form name="form1" method="post" action="admin_copyright.asp"><tr><td class=zonelovess align="center" colspan="2">��վ��Ȩ��Ϣ����</td></tr><tr><td class=zoneloveqs colspan="2">*ȫ��֧��HTML����</td></tr><tr class=zoneloveds><td width="80">ѧУ�ſ���</td><td width="*"><textarea name="gy" cols="60" rows="8" class='zonelove'><%=gy%></textarea></td></tr><tr class=zoneloveds><td width="80">��ϵ���ǣ�</td><td width="*"><textarea name="lx" cols="60" rows="8" class='zonelove'><%=lx%></textarea></td></tr><tr class=zoneloveds><td width="80">�쵼�´ǣ�</td><td width="*"><textarea name="gg" cols="60" rows="8" class='zonelove'><%=gg%></textarea></td></tr><tr class=zoneloveds><td width="80">��֯������</td><td width="*"><textarea name="dt" cols="60" rows="8" class='zonelove'><%=dt%></textarea></td></tr><tr class=zoneloveqs><td align="center" colspan="2"><input type="submit" name="Submit" value="ȷ���ύ"><input type="reset" name="Reset" value="�����д"></td></tr><input type="hidden" name="action" value="modpass"><input type="hidden" name="MM_insert" value="true"></form></table>
<%
rs.close
set rs=nothing
end if
%>