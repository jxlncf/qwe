<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<!--#include file="inc/FORMAT.asp"-->
<!--#include file="inc/UBBshow.asp"-->
<%
if session("adminlogin")<>sessionvar then
  Response.Write("<script language=javascript>alert('你尚未登录，或者超时了！请重新登录');this.top.location.href='admin.asp';</script>")
  response.end
else
if Request.form("MM_insert") then
if request.form("action")="newnews" then
sql="select * from news"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim keyword,title,content,color
keyword=trim(replace(request.form("news_keyword"),"'",""))
title=trim(replace(request.form("news_title"),"'",""))
content=rtrim(replace(request.form("news_content"),"'",""))
content=trim(replace(content," ","&nbsp;"))
color=request.form("news_color")
if keyword="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择校内新闻类型！');history.back(1);</script>")
else
  rs("news_keyword")=keyword
end if
if title="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入校内新闻的标题！');history.back(1);</script>")
else
  rs("news_title")=title
end if
if content="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入校内新闻内容！');history.back(1);</script>")
else
  rs("news_content")=content
end if
if color="" then
  rs("news_color")=""
else
  rs("news_color")=color
end if
if not founderr then
rs.update
rs.close
set rs=nothing
  sql="update allcount set newscount = newscount + 1"
  conn.execute(sql)
response.redirect "admin_news.asp"
else
call diserror()
response.end
end if
end if
if request.form("action")="editnews" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的校内新闻id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from news where news_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
keyword=trim(replace(request.form("news_keyword"),"'",""))
title=trim(replace(request.form("news_title"),"'",""))
content=rtrim(replace(request.form("news_content"),"'",""))
content=trim(replace(content," ","&nbsp;"))
color=request.form("news_color")
if keyword="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须选择校内新闻类型！');history.back(1);</script>")
else
  rs("news_keyword")=keyword
end if
if title="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入校内新闻的标题！');history.back(1);</script>")
else
  rs("news_title")=title
end if
if content="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须输入校内新闻内容！');history.back(1);</script>")
else
  rs("news_content")=content
end if
if color="" then
  rs("news_color")=""
else
  rs("news_color")=color
end if
if not founderr then
rs.update
rs.close
set rs=nothing
response.redirect "admin_news.asp"
else
call diserror()
response.end
end if
end if
if request.form("action")="delnews" then
if request.Form("id")="" then
  founderr=true
  Response.Write("<script language=javascript>alert('你必须指定操作的对象！');history.back(1);</script>")
else
  if not isInteger(request.form("id")) then
    founderr=true
    Response.Write("<script language=javascript>alert('非法的校内新闻id参数。');history.back(1);</script>")
  end if
end if
if founderr then
  call diserror()
  response.End
end if
sql="select * from news where news_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=nothing
  sql="update allcount set newscount = newscount - 1"
  conn.execute(sql)
response.redirect "admin_news.asp"
end if
end if
dim totalnews,Currentpage,totalpages,i
sql="select * from news order by news_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<script language="JavaScript">
<!--
function insertstr(str)
{document.form1.news_content.value=document.form1.news_content.value+str;}
var Quote = 0;
var Bold  = 0;
var Italic = 0;
var Underline = 0;
var Code = 0;
var Center = 0;
var Strike = 0;
var Sound = 0;
var Swf = 0;
var Ra = 0;
var Rm = 0;
var Marquee = 0;
var Fly = 0;
var fanzi=0;
var text_enter_url      = "请输入连接网址";
var text_enter_iFrame      = "请输入连接网址";
var text_enter_txt      = "请输入连接说明";
var text_enter_image    = "请输入图片网址";
var text_enter_sound    = "请输入声音文件网址";
var text_enter_swf      = "请输入FLASH动画网址";
var text_enter_ra      = "请输入Real音乐网址";
var text_enter_rm      = "请输入Real影片网址";
var text_enter_wmv      = "请输入Media影片网址";
var text_enter_wma      = "请输入Media音乐网址";
var text_enter_mov      = "请输入QuickTime音乐网址";
var text_enter_sw      = "请输入shockwave音乐网址";
var text_enter_email    = "请输入邮件网址";
var error_no_url        = "您必须输入网址";
var error_no_txt        = "您必须连接说明";
var error_no_title      = "您必须输入首页标题";
var error_no_email      = "您必须输入邮件网址";
var error_no_gset       = "必须正确按照各式输入！";
var error_no_gtxt       = "必须输入文字！";
var text_enter_guang1   = "文字的长度、颜色和边界大小";
var text_enter_guang2   = "要产生效果的文字！";
function commentWrite(NewCode) {
document.form1.news_content.value+=NewCode;
document.form1.news_content.focus();
return;
}
function storeCaret(text) { 
	if (text.createTextRange) {
		text.caretPos = document.selection.createRange().duplicate();
	}
        if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('文章正在发布，请耐心等待！');return false;}this.document.form.submit();}
}
function AddText(text) {
	if (document.form1.news_content.createTextRange && document.form1.news_content.caretPos) {      
		var caretPos = document.form1.news_content.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else document.form1.news_content.value += text;
	document.form1.news_content.focus(caretPos);
}
function inputs(str)
{
AddText(str);
}
function Curl() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_url, "http://");
var enterTxT   = prompt(text_enter_txt, enterURL);
if (!enterURL)    {
FoundErrors += "\n" + error_no_url;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_txt;
}
if (FoundErrors)  {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[URL="+enterURL+"]"+enterTxT+"[/URL]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cimage() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_image, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[IMG]"+enterURL+"[/IMG]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function CiFrame() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_iFrame, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[frame]"+enterURL+"[/frame]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cemail() {
var emailAddress = prompt(text_enter_email,"");
if (!emailAddress) { alert(error_no_email); return; }
var ToAdd = "[EMAIL]"+emailAddress+"[/EMAIL]";
commentWrite(ToAdd);
}
function Ccode() {
if (Code == 0) {
ToAdd = "[CODE]";
document.form.code.value = " 代码*";
Code = 1;
} else {
ToAdd = "[/CODE]";
document.form.code.value = " 代码 ";
Code = 0;
}
commentWrite(ToAdd);
}
function Cquote() {
fontbegin="[QUOTE]";
fontend="[/QUOTE]";
fontchuli();
}
function Cbold() {
fontbegin="[B]";
fontend="[/B]";
fontchuli();
}
function Citalic() {
fontbegin="[I]";
fontend="[/I]";
fontchuli();
}
function Cunder() {
fontbegin="[U]";
fontend="[/U]";
fontchuli();
}
function Ccenter() {
fontbegin="[center]";
fontend="[/center]";
fontchuli();
}
function Cstrike() {
fontbegin="[strike]";
fontend="[/strike]";
fontchuli();
}
function Csound() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_sound, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[SOUND]"+enterURL+"[/SOUND]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}

function Cswf() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_swf, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[FLASH]"+enterURL+"[/FLASH]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cra() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_ra, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[RA]"+enterURL+"[/RA]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Crm() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_rm, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[RM=500,350]"+enterURL+"[/RM]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cwmv() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_wmv, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[MP=500,350]"+enterURL+"[/MP]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}

function Cfanzi() {
fontbegin="[xray]";
fontend="[/xray]";
fontchuli();
}

function Cwma() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_wma, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[wma]"+enterURL+"[/wma]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cmov() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_mov, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[QT=500,350]"+enterURL+"[/QT]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cdir() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_sw, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[DIR=500,350]"+enterURL+"[/DIR]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}
function Cmarquee() {
fontbegin="[move]";
fontend="[/move]";
fontchuli();
}
function Cfly() {
fontbegin="[fly]";
fontend="[/fly]";
fontchuli();
}

function paste(text) {
	if (opener.document.form1.news_content.createTextRange && opener.document.form1.news_content.caretPos) {      
		var caretPos = opener.document.form1.news_content.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else opener.document.form1.news_content.value += text;
	opener.document.form1.news_content.focus(caretPos);
}

function showsize(size){
fontbegin="[size="+size+"]";
fontend="[/size]";
fontchuli();
}

function showfont(font){
fontbegin="[face="+font+"]";
fontend="[/face]";
fontchuli();
}

function showcolor(color){
fontbegin="[color="+color+"]";
fontend="[/color]";
fontchuli();
}

function fontchuli(){
if ((document.selection)&&(document.selection.type == "Text")) {
var range = document.selection.createRange();
var ch_text=range.text;
range.text = fontbegin + ch_text + fontend;
} 
else {
document.form1.news_content.value=fontbegin+document.form1.news_content.value+fontend;
document.form1.news_content.focus();
}
}

function Cguang() {
var FoundErrors = '';
var enterSET   = prompt(text_enter_guang1, "255,red,2");
var enterTxT   = prompt(text_enter_guang2, "文字");
if (!enterSET)    {
FoundErrors += "\n" + error_no_gset;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_gtxt;
}
if (FoundErrors)  {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[glow="+enterSET+"]"+enterTxT+"[/glow]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}

function Cying() {
var FoundErrors = '';
var enterSET   = prompt(text_enter_guang1, "255,blue,1");
var enterTxT   = prompt(text_enter_guang2, "文字");
if (!enterSET)    {
FoundErrors += "\n" + error_no_gset;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_gtxt;
}
if (FoundErrors)  {
alert("错误！"+FoundErrors);
return;
}
var ToAdd = "[SHADOW="+enterSET+"]"+enterTxT+"[/SHADOW]";
document.form1.news_content.value+=ToAdd;
document.form1.news_content.focus();
}

ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form1.submit();}}
}
function DoTitle(addTitle) { 
var revisedTitle; 
var currentTitle = document.form1.subject.value; 
revisedTitle = currentTitle+addTitle; 
document.form1.subject.value=revisedTitle; 
document.form1.subject.focus(); 
return; }

function insertsmilie(smilieface){

	document.form1.news_content.value+=smilieface;
}
//-->
</script>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<%if request.QueryString("action")="" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<tr class=zonelovess> <td colspan="4">校内新闻管理</font></td></tr>
<tr align="center" class=zoneloveqs><td width="10%">编号</td><td width="60%">标题&amp;发表时间</td><td width="10%">评论</td><td width="20%">操作</td></tr>
<%if not rs.eof then
rs.movefirst
rs.pagesize=newsperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   end if
else
   currentpage=1
end if
   totalnews=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*newsperpage<totalnews then
	       rs.move(currentpage-1)*newsperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   end if
   end if
   if (totalnews mod newsperpage)=0 then
      totalpages=totalnews\newsperpage
   else
      totalpages=totalnews\newsperpage+1
   end if
i=0
do while not rs.eof and i<newsperpage%>
<tr class=zoneloveds><td align="center"><%=rs("news_id")%></td><td><a href="shownews.asp?news_id=<%=rs("news_id")%>" target="_blank"><%=rs("news_title")%></a> <span class="date"><%=rs("news_date")%></span></td><td align="center"><font color=red><%=rs("reviewcount")%></font></td><td align="center"> <a href="admin_view.asp?action=news&news_id=<%=rs("news_id")%>">评论</a> <a href="admin_news.asp?id=<%=rs("news_id")%>&action=editnews#editnews">编辑</a> <a href="admin_news.asp?id=<%=rs("news_id")%>&action=delnews#delnews">删除</a></td></tr>
<%i=i+1
rs.movenext
loop
else
if rs.eof and rs.bof then%>
<tr align="center" class=zoneloveds><td colspan="4">当前没有校内新闻！</td></tr>
<%end if
end if%>
</table>
<table align="center" width="98%" align="center" border="0" cellspacing="0" cellpadding="0"><form name="form2" method="post" action="admin_news.asp">
<tr class=zoneloveds><td align="right"><%=currentpage%> /<%=totalpages%>页,<%=totalnews%>条记录/<%=newsperpage%>篇每页. 
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
<a href="admin_news.asp?page=<%=i%>"><%=i%></a> 
<%end if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
end if%>
              <a href="admin_news.asp?page=<%=page%>" title="下一页">>></a> 
              <%end if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%end if
if request.QueryString("action")="newnews" then%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td><a name="newnews">新的校内新闻</a></font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>标题-<select name="news_color"><%call zonelovecolor()%></select>
              <input type="text" name="news_title" class="textarea" size="50">  <select name="news_keyword">
<option value="0">校内新闻类型</option>
<option value="站内公告">站内公告</option>
<option value="本校新闻">本校新闻</option>
<option value="校内快讯">校内快讯</option>
<option value="其它类型">其它类型</option>
</select>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td>
              <textarea name="news_content" style="display:none"></textarea>
<iframe id="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=news_content&style=blue" frameborder="0" scrolling="No" width="550" height="320"></iframe>
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center">
              <input type="submit" name="Submit2" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
 <input type="button" value="返回"  onClick="location.href='admin_news.asp'" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="newnews">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.QueryString("action")="editnews" then
if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from news where news_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td><a name="editnews">修改校内新闻</a></font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>标题-<select name="news_color"><%call zonelovecolor()%></select>
              <input type="text" name="news_title" class="textarea" size="50" value="<%=rs("news_title")%>">  <select name="news_keyword">
<option value="0">校内新闻类型</option>
<option value="站内公告"<%if rs("news_keyword")="站内公告" then response.write "selected"%>>站内公告</option>
<option value="本校新闻"<%if rs("news_keyword")="本校新闻" then response.write "selected"%>>本校新闻</option>
<option value="校内快讯"<%if rs("news_keyword")="校内快讯" then response.write "selected"%>>校内快讯</option>
<option value="其它类型"<%if rs("news_keyword")="其它类型" then response.write "selected"%>>其它类型</option>
</select>
            </td>
          </tr>
                   <tr class=zoneloveds> 
            <td> <textarea name="news_content" style="display:none"><%=Server.HTMLEncode(rs("news_content"))%></textarea>
<iframe id="eWebEditor2" src="eWebEditor/ewebeditor.htm?id=news_content&style=blue" frameborder="0" scrolling="No" width="550" height="320"></iframe> 
            </td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
 <input type="button" value="返回"  onClick="location.href='admin_news.asp'" class="button"></td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("news_id")%>">
		  <input type="hidden" name="action" value="editnews">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%end if
if request.QueryString("action")="delnews" then
if request.querystring("id")="" then
  Response.Write("<script language=javascript>alert('请指定操作的对象！');history.back(1);</script>")
  response.end
else
  if not isinteger(request.querystring("id")) then
    Response.Write("<script language=javascript>alert('非法的校内新闻ID参数！');history.back(1);</script>")
	response.end
  end if
end if
sql="select * from news where news_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr class=zonelovess> 
            <td><a name="delnews" id="delnews">删除校内新闻</a></font></td>
          </tr>
          <tr class=zoneloveds> 
            <td>标题-<%=rs("news_title")%> 校内新闻类型- <%=rs("news_keyword")%>
            </td>
          </tr>
          <tr class=zoneloveds> 
            <td><%=ubb2html(formatStr(autourl(rs("news_content"))), true, true)%> 　</td>
          </tr>
          <tr class=zoneloveqs> 
            <td align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
 <input type="button" value="返回"  onClick="location.href='admin_news.asp'" class="button"></td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("news_id")%>">
          <input type="hidden" name="action" value="delnews">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%end if
rs.close
set rs=nothing
end if
%>