<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim adminname
dim adminpwd
if request.QueryString("action")="search" then
	dim word,engine
	word = request.Form("word")
	engine = request.Form("search")
	Select case engine
		
	end select
end if
if request("action")="adminlogin" then
if request("CheckCode")<>Session("CheckCode") then
 Response.Write("<script language=javascript>alert('��������ȷ����֤�룡');this.location.href='admin.asp';</script>") 
 session("adminlogin")=""
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
Response.Write("<script language=javascript>alert('���ύ��·�����󣬽�ֹ��վ���ⲿ�ύ�����벻Ҫ�Ҹò�����');this.location.href='admin.asp';</script>") 
response.end
end if
adminname=MD5(trim(replace(request("adminname"),"'","")))
adminpwd=MD5(trim(replace(request("adminpwd"),"'","")))

if adminname="" and adminpwd="" then
	   	   Response.Write("<script language=javascript>alert('�������û��������룡');this.location.href='admin.asp';</script>")
		   Response.End
end if
sql="select * from admin where admin_name='"&adminname&"' and admin_password='"&adminpwd&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
   rs.close
   set rs=nothing
	   	   Response.Write("<script language=javascript>alert('��������û��������벻��ȷ!');this.location.href='admin.asp';</script>")
		   Response.End
else
   session("adminlogin")=sessionvar
   session("issuper")=rs("admin_id")
   session.timeout=50
   Session("CheckCode")=""
   rs.close
   set rs=nothing
end if
elseif request("action")="logout" then
  session("adminlogin")=""
  session("issuper")=""
  Response.write "<script>window.document.location.href='./admin.asp';</script>"
end if
if session("adminlogin")=sessionvar then
frame=request("frame")
if frame="" then
%>
<html>
<head>
<title>�氮�ռ�<%=webvar%>��̨����ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset rows='*' id='Frame' cols='185,*' framespacing='0' frameborder='no' border='0'><frame src='?frame=menu' scrolling='auto' id='menu' name='menu' noresize marginwidth='5' marginheight='5'><frame src='?frame=main' name='main' id='main' scrolling='auto' noresize marginwidth='0' marginheight='0'></frameset>
��<noframes>
��<body>
����<p>��ҳʹ���˿�ܽṹ�����������������֧�������뽫�������������ΪIE5.0����ߵİ汾��</p>
��</body>
��</noframes>
</html>
<%elseif frame="menu" then%>
<html>
<head>
<title>����˵�</title>
<style type=text/css>
body  { background:#C3C7D2; font:Verdana 12px; 
SCROLLBAR-FACE-COLOR: #C3C7D2; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #C3C7D2; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #C3C7D2; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #C3C7D2;
}
table  { border:0px; }
td  { font:normal 12px ����;}
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:0px; left:8px; color:#000000; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:0px; left:8px; color:#999999; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showmenu_item(sid)
{
	which = eval("menu_item" + sid);
	if (which.style.display == "none")
	{
		var i = 1
		while(i<8){
			eval("menu_item"+ i +".style.display=\"none\";");
			eval("menuTitle"+ i +".background=\"img/title_bg_show.gif\";");
			i++;
		}
		eval("menu_item" + sid + ".style.display=\"\";");
		eval("menuTitle"+ sid + ".background=\"img/title_bg_hide.gif\";")
	}else{
		eval("menu_item" + sid + ".style.display=\"none\";");
		eval("menuTitle"+ sid + ".background=\"img/title_bg_show.gif\";")
	}
}
</SCRIPT>
  <table width="158" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="158" height="38" background="img/title.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="151" height="16"></td>
        </tr>
        <tr>
          <td><div align="center"><font color="#FFFFFF"><strong></strong></font></div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      
    <td height="25" class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=img/title_bg_quit.gif bgcolor="#7898E0"><span>&nbsp;<a href="?frame=main" target="main"><strong>������ҳ</strong></a> 
      | <a href="?action=logout" target="_top"><strong>�˳�</strong></a></span></td>
    </tr>
  </table>
  <br>
  <table cellpadding=0 cellspacing=0 width=158>
    <tr>
      
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle1 onclick="showmenu_item(1)"><span>�������ù���</span> 
    </td>
    </tr>
    <tr>
      <td style="display:none;" id='menu_item1'><div class=sec_menu style="width:158">
          <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="4"></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_config.asp" target=main>��վ����</a> | <a href="admin_admin.asp?action=modpass" target=main>�޸�����</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="book.asp" target=book>���Թ���</a> | <a href="admin_admin.asp?action=recount" target=main>��ҳ����</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_diary.asp" target="main">�������</a> | <a href="admin_diary.asp?action=newdiary" target="main">������</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_news.asp" target="main" title="���Ź���">���Ź���</a> | <a href="admin_news.asp?action=newnews#newnews" target="main">��������</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_vote.asp?action=vote" target="main">ͶƱ����</a> | <a href="admin_vote.asp?action=newvote" target="main">����ͶƱ</a></td>
            </tr>
<tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_copyright.asp" target="main">����ѧУ������Ϣ</a></td>
            </tr>
          </table>
      </div>
          <div  style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=20></td>
              </tr>
            </table>
        </div></td>
    </tr>
  </table>
      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle2 onclick="showmenu_item(2)"><span>����ƽ̨����</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item2'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=djcat" target="main">�������</a> | <a href="admin_dj.asp?action=newdjcat" target="main">��ӷ���</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=dj" target="main">���ֹ���</a> | <a href="admin_dj.asp?action=newdj" target="main">�������</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=djerror" target="main">�������ӹ���</a></td>
                </tr>
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
      </table>
      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle3 onclick="showmenu_item(3)"><span>����ģ�����</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item3'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_art.asp" target="main">�������</a> | <a href="admin_art.asp?action=newartcat" target="main">��ӷ���</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_art.asp?action=cat" target="main">���¹���</a> | <a href="admin_art.asp?action=newart" target="main">�������</a></td>
                </tr>
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
      </table>
      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle4 onclick="showmenu_item(4)"><span>������ع���</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item4'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_down.asp" target="main">�������</a> | <a href="admin_down.asp?action=newcat" target="main">��ӷ���</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_down.asp?action=list" target="main">�������</a> | <a href="admin_down.asp?action=newsoft" target="main">������</a></td>
                </tr>
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
      </table>
      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle7 onclick="showmenu_item(7)"><span>У԰������</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item7'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_pic.asp?action=piccat" target="main">�������</a> | <a href="admin_pic.asp?action=newpiccat" target="main">��ӷ���</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_pic.asp?action=pic" target="main">ͼƬ����</a> | <a href="admin_pic.asp?action=newpic" target="main">���ͼƬ</a></td>
                </tr>
              </table>
            </div>
              </td>
        </tr>
      </table>


<table cellpadding=0 cellspacing=0 width=158>
        <tr>
    
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle9 onclick="showmenu_item(9)"><span>���߱�������</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item9'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="baoming/list.asp" target="main">���߱�������</a> </td>
                </tr>
                
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
  </table>




      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
    
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle5 onclick="showmenu_item(5)"><span>��ʦ��ɹ���</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item5'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_web.asp?action=cscat" target="main">�������</a> | <a href="admin_web.asp?action=newcscat" target="main">��ӷ���</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_web.asp?action=sites" target="main">��ʦ����</a> | <a href="admin_web.asp?action=newcs" target="main">��ӽ�ʦ</a></td>
                </tr>
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
  </table>
      <table cellpadding=0 cellspacing=0 width=158>
        <tr>
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle6 onclick="showmenu_item(6)"><span>�������ӹ���</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item6'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_link.asp" target="main">�������</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle">���� <a href="admin_link.asp?action=link" target="main">����</a>|<a href="admin_link.asp?action=check" target="main">���</a>|<a href="admin_link.asp?action=lk" target="main">������</a></td>
                </tr>
              </table>
          </div>
              <div  style="width:158">
                <table cellpadding=0 cellspacing=0 align=center width=135>
                  <tr>
                    <td height=20></td>
                  </tr>
                </table>
            </div></td>
        </tr>
      </table>
      
               </td>
          </tr>
      </table></td>
    </tr>
  </table>

</body>
</html>
<%elseif frame="main" then%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<script>if(top==self)top.location="admin.asp" </script>
</head><body>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
  <tr> 
    <td class=zonelovess align="center" colspan="2">
      ��Ϣͳ��</td>
  </tr>
  <tr> 
    <td class=zoneloveqs colspan="2">��������<a href="http://www.zonelove.net" target=_blank> �氮�ռ�</a> ��Ȩ�� <%=webname%> ʹ�ã���ǰʹ�ð汾Ϊ <%=webvar%></td>
  </tr>
 <tr>
    <td width="50%" class="zoneloveds">&nbsp;�������������еĶ˿ڣ�<%=Request.ServerVariables("server_port")%></td>
    <td class="zoneloveds">&nbsp;�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="zoneloveds"> &nbsp;���������ƣ�<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td class="zoneloveds"> &nbsp;������IP��<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;վ������·����<%=Request.ServerVariables("path_translated")%> </td>
    <td class="zoneloveds">&nbsp;����·����<%=Request.ServerVariables("server_name")%></td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;������Application������<%=Application.Contents.Count%> ��</td>
    <td class="zoneloveds">&nbsp;������Session������<%=Session.Contents.Count%> ��</td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;��������ǰʱ�䣺<%=now()%></td>
    <td class="zoneloveds">&nbsp;�ű����ӳ�ʱʱ�䣺<%=Server.ScriptTimeout%> ��</td>
  </tr>
</table>
<br>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
  <tr> 
    <td class=zonelovess align="center" colspan="2">
      ��ݲ���</td>
  </tr>
 <tr>
    <td width="50%" class="zoneloveds">&nbsp;������ǵ�һ��ʹ�ã��벻Ҫ������<input type="submit" name="Submit" value="������վ����" onclick="window.location.href='admin_config.asp'">�������Լ��������԰���ɴ˿�ʼ��</td>
  </tr>
  <tr>
    <td class="zoneloveds"> &nbsp;��ҳͳ�Ƴ����밴���� <input type="submit" name="Submit" value="����ͳ����Ϣ" onclick="window.location.href='admin_admin.asp?action=recount'"> ��Ҫ�ǵľ���������Ϣ������ͳ�Ƴ���</td>
  </tr>
</table>
<br>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
  <tr> 

  </tr>
  </form>
</table>
<br>
</body>
</html>
<%end if%>
<%else%>
<HTML><HEAD><TITLE>��̨����<%webvar%>�����½</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<script language=javascript>
  function xxg()
  {
      if (document.form1.adminname.value==""){
	      alert("�������û�����")
		  document.form1.adminname.focus();
		  return false
		    }
	  if (document.form1.adminpwd.value==""){
	      alert("���������룿");
		  document.form1.adminpwd.focus();
		  return false
		  } 
          if (document.form1.CheckCode.value==""){
	      alert("��������֤�룿");
		  document.form1.CheckCode.focus();
		  return false
		  }
		  return true
  }
  function reset_form()
  {
   document.form1.adminname.value="";
   document.form1.adminpwd.value="";
   document.form1.VerifyCode.value="";
   document.form1.adminname.focus;
  }
</script>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))><br><br><br> 
<div align="center">
	<table border="0" width="538" cellspacing="0" cellpadding="0" id="table1" height="358">
		<tr>
			<td width="538" background="img/admin.gif">��<div align="center">
	<table width="280" border="0" cellpadding="0" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="admin.asp?action=adminlogin" onsubmit="return xxg()">
  <tr> 
    <td class=zonelovess align="center"  colspan="2">����Ա���</td>
  </tr>
  <tr class=zoneloveds> 
    <td class=zoneloveds width="80">&nbsp;<b>�� ����</b></td><td class=zoneloveds width="200">
	<input type="text" name="adminname" class="textarea" size="20" autocomplete="off" style="border: 5px dotted #FF0000"></td>
  <tr> 
    <td class=zoneloveds>&nbsp;<b>�� �룺</b></td><td class=zoneloveds>
	<input type="password" name="adminpwd" class="textarea" size="20" style="border: 5px dotted #FF0000"></td>
  </tr>
  </tr>
<tr> 
            <td class=zoneloveds>&nbsp;<b>��֤��:</b></td>
            <td class=zoneloveds>
			<input name="CheckCode" type="text" size="14" autocomplete="off" style="border: 5px dotted #FF0000">
              &nbsp;<b><img src="inc/Code.asp" align="absmiddle"></b></td>
          </tr>
  <tr>
    <td class=zoneloveqs colspan="2" align="center"><input type="submit" name="Submit" value="��¼" class="button"></td>
  </tr>
</table></div><p>��</td>
		</tr>
	</table>
</div><p>��</td>
		</tr>
	</table>
</div>
</body></html>
<%end if%>