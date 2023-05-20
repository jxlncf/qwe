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
 Response.Write("<script language=javascript>alert('请输入正确的认证码！');this.location.href='admin.asp';</script>") 
 session("adminlogin")=""
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
Response.Write("<script language=javascript>alert('你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！');this.location.href='admin.asp';</script>") 
response.end
end if
adminname=MD5(trim(replace(request("adminname"),"'","")))
adminpwd=MD5(trim(replace(request("adminpwd"),"'","")))

if adminname="" and adminpwd="" then
	   	   Response.Write("<script language=javascript>alert('请输入用户名或密码！');this.location.href='admin.asp';</script>")
		   Response.End
end if
sql="select * from admin where admin_name='"&adminname&"' and admin_password='"&adminpwd&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
   rs.close
   set rs=nothing
	   	   Response.Write("<script language=javascript>alert('您输入的用户名和密码不正确!');this.location.href='admin.asp';</script>")
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
<title>真爱空间<%=webvar%>后台管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset rows='*' id='Frame' cols='185,*' framespacing='0' frameborder='no' border='0'><frame src='?frame=menu' scrolling='auto' id='menu' name='menu' noresize marginwidth='5' marginheight='5'><frame src='?frame=main' name='main' id='main' scrolling='auto' noresize marginwidth='0' marginheight='0'></frameset>
　<noframes>
　<body>
　　<p>本页使用了框架结构，但是您的浏览器不支持它。请将您的浏览器升级为IE5.0或更高的版本！</p>
　</body>
　</noframes>
</html>
<%elseif frame="menu" then%>
<html>
<head>
<title>管理菜单</title>
<style type=text/css>
body  { background:#C3C7D2; font:Verdana 12px; 
SCROLLBAR-FACE-COLOR: #C3C7D2; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #C3C7D2; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #C3C7D2; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #C3C7D2;
}
table  { border:0px; }
td  { font:normal 12px 宋体;}
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
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
      
    <td height="25" class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=img/title_bg_quit.gif bgcolor="#7898E0"><span>&nbsp;<a href="?frame=main" target="main"><strong>管理首页</strong></a> 
      | <a href="?action=logout" target="_top"><strong>退出</strong></a></span></td>
    </tr>
  </table>
  <br>
  <table cellpadding=0 cellspacing=0 width=158>
    <tr>
      
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle1 onclick="showmenu_item(1)"><span>基本设置管理</span> 
    </td>
    </tr>
    <tr>
      <td style="display:none;" id='menu_item1'><div class=sec_menu style="width:158">
          <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="4"></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_config.asp" target=main>网站设置</a> | <a href="admin_admin.asp?action=modpass" target=main>修改密码</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="book.asp" target=book>留言管理</a> | <a href="admin_admin.asp?action=recount" target=main>首页更新</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_diary.asp" target="main">公告管理</a> | <a href="admin_diary.asp?action=newdiary" target="main">新添公告</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_news.asp" target="main" title="新闻管理">新闻管理</a> | <a href="admin_news.asp?action=newnews#newnews" target="main">新添新闻</a></td>
            </tr>
            <tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_vote.asp?action=vote" target="main">投票管理</a> | <a href="admin_vote.asp?action=newvote" target="main">新添投票</a></td>
            </tr>
<tr>
              <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_copyright.asp" target="main">设置学校其它信息</a></td>
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
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle2 onclick="showmenu_item(2)"><span>视听平台管理</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item2'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=djcat" target="main">分类管理</a> | <a href="admin_dj.asp?action=newdjcat" target="main">添加分类</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=dj" target="main">音乐管理</a> | <a href="admin_dj.asp?action=newdj" target="main">添加音乐</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_dj.asp?action=djerror" target="main">错误连接管理</a></td>
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
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle3 onclick="showmenu_item(3)"><span>文章模块管理</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item3'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_art.asp" target="main">分类管理</a> | <a href="admin_art.asp?action=newartcat" target="main">添加分类</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_art.asp?action=cat" target="main">文章管理</a> | <a href="admin_art.asp?action=newart" target="main">添加文章</a></td>
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
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle4 onclick="showmenu_item(4)"><span>软件下载管理</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item4'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_down.asp" target="main">分类管理</a> | <a href="admin_down.asp?action=newcat" target="main">添加分类</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_down.asp?action=list" target="main">软件管理</a> | <a href="admin_down.asp?action=newsoft" target="main">添加软件</a></td>
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
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle7 onclick="showmenu_item(7)"><span>校园相册管理</span> 
    </td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item7'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_pic.asp?action=piccat" target="main">分类管理</a> | <a href="admin_pic.asp?action=newpiccat" target="main">添加分类</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" alt width="15" height="20" border="0" align="absmiddle"><a href="admin_pic.asp?action=pic" target="main">图片管理</a> | <a href="admin_pic.asp?action=newpic" target="main">添加图片</a></td>
                </tr>
              </table>
            </div>
              </td>
        </tr>
      </table>


<table cellpadding=0 cellspacing=0 width=158>
        <tr>
    
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle9 onclick="showmenu_item(9)"><span>在线报名管理</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item9'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="baoming/list.asp" target="main">在线报名管理</a> </td>
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
    
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle5 onclick="showmenu_item(5)"><span>教师风采管理</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item5'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_web.asp?action=cscat" target="main">分类管理</a> | <a href="admin_web.asp?action=newcscat" target="main">添加分类</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_web.asp?action=sites" target="main">教师管理</a> | <a href="admin_web.asp?action=newcs" target="main">添加教师</a></td>
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
          
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="img/title_bg_show.gif" id=menuTitle6 onclick="showmenu_item(6)"><span>友情链接管理</span></td>
        </tr>
        <tr>
          <td style="display:none" id='menu_item6'><div class=sec_menu style="width:158">
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="4"></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle"><a href="admin_link.asp" target="main">分类管理</a></td>
                </tr>
                <tr>
                  <td height="20"><img src="img/bullet.gif" width="15" height="20" align="absmiddle">链接 <a href="admin_link.asp?action=link" target="main">管理</a>|<a href="admin_link.asp?action=check" target="main">审核</a>|<a href="admin_link.asp?action=lk" target="main">黑名单</a></td>
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
      信息统计</td>
  </tr>
  <tr> 
    <td class=zoneloveqs colspan="2">本程序由<a href="http://www.zonelove.net" target=_blank> 真爱空间</a> 授权给 <%=webname%> 使用，当前使用版本为 <%=webvar%></td>
  </tr>
 <tr>
    <td width="50%" class="zoneloveds">&nbsp;服务器正在运行的端口：<%=Request.ServerVariables("server_port")%></td>
    <td class="zoneloveds">&nbsp;脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="zoneloveds"> &nbsp;服务器名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td class="zoneloveds"> &nbsp;服务器IP：<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;站点物理路径：<%=Request.ServerVariables("path_translated")%> </td>
    <td class="zoneloveds">&nbsp;虚拟路径：<%=Request.ServerVariables("server_name")%></td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;服务器Application数量：<%=Application.Contents.Count%> 个</td>
    <td class="zoneloveds">&nbsp;服务器Session数量：<%=Session.Contents.Count%> 个</td>
  </tr>
  <tr>
    <td class="zoneloveds">&nbsp;服务器当前时间：<%=now()%></td>
    <td class="zoneloveds">&nbsp;脚本连接超时时间：<%=Server.ScriptTimeout%> 秒</td>
  </tr>
</table>
<br>
<table width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
  <tr> 
    <td class=zonelovess align="center" colspan="2">
      快捷操作</td>
  </tr>
 <tr>
    <td width="50%" class="zoneloveds">&nbsp;如果你是第一次使用，请不要忘了先<input type="submit" name="Submit" value="更改网站资料" onclick="window.location.href='admin_config.asp'">，打造自己的网络家园，由此开始！</td>
  </tr>
  <tr>
    <td class="zoneloveds"> &nbsp;首页统计出错？请按这里 <input type="submit" name="Submit" value="更新统计信息" onclick="window.location.href='admin_admin.asp?action=recount'"> ，要记的经常更新信息，以免统计出错！</td>
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
<HTML><HEAD><TITLE>后台管理—<%webvar%>管理登陆</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<script language=javascript>
  function xxg()
  {
      if (document.form1.adminname.value==""){
	      alert("请输入用户名？")
		  document.form1.adminname.focus();
		  return false
		    }
	  if (document.form1.adminpwd.value==""){
	      alert("请输入密码？");
		  document.form1.adminpwd.focus();
		  return false
		  } 
          if (document.form1.CheckCode.value==""){
	      alert("请输入验证码？");
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
			<td width="538" background="img/admin.gif">　<div align="center">
	<table width="280" border="0" cellpadding="0" class=zonelovebk style="border-collapse: collapse">
<form name="form1" method="post" action="admin.asp?action=adminlogin" onsubmit="return xxg()">
  <tr> 
    <td class=zonelovess align="center"  colspan="2">管理员入口</td>
  </tr>
  <tr class=zoneloveds> 
    <td class=zoneloveds width="80">&nbsp;<b>用 户：</b></td><td class=zoneloveds width="200">
	<input type="text" name="adminname" class="textarea" size="20" autocomplete="off" style="border: 5px dotted #FF0000"></td>
  <tr> 
    <td class=zoneloveds>&nbsp;<b>密 码：</b></td><td class=zoneloveds>
	<input type="password" name="adminpwd" class="textarea" size="20" style="border: 5px dotted #FF0000"></td>
  </tr>
  </tr>
<tr> 
            <td class=zoneloveds>&nbsp;<b>认证码:</b></td>
            <td class=zoneloveds>
			<input name="CheckCode" type="text" size="14" autocomplete="off" style="border: 5px dotted #FF0000">
              &nbsp;<b><img src="inc/Code.asp" align="absmiddle"></b></td>
          </tr>
  <tr>
    <td class=zoneloveqs colspan="2" align="center"><input type="submit" name="Submit" value="登录" class="button"></td>
  </tr>
</table></div><p>　</td>
		</tr>
	</table>
</div><p>　</td>
		</tr>
	</table>
</div>
</body></html>
<%end if%>