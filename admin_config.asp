<!--#include file="inc/config.asp"-->
<%
if session("adminlogin")<>sessionvar then
  founderr=true
Response.Write("<script language=javascript>alert('您尚未登陆或登陆超时，请重新登陆!！');this.top.location.href='admin.asp';</script>")
  response.end
else%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/admin.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) >
<%
select case Request("menu")
case ""
config
case "config"
config
case "configok"
configok
end select
sub config%>
<table width=98% align="center" border="1" cellspacing="0" cellpadding="4" class=zonelovebk style="border-collapse: collapse">
<form method="post" action="?menu=configok">
<tr class=zonelovess>
<td colspan="2" height="25">网站基本设置</td>
</tr>
<tr class=zoneloveqs>
<td colspan="2">本页用于网站首页的基本信息及个性化选择</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">网站名称：</td>
<td width="*" height="25"><input size="30" name="webname" value="<%=webname%>"> 如：真爱空间校网</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">网站地址：</td>
<td width="*" height="25"><input size="30" value="<%=weburl%>" name="weburl"> 如：http://www.zonelove.net</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">站长邮箱：</td>
<td width="*" height="25"><input size="30" value="<%=webmail%>" name="webmail"> 如：zonelove8@163.com</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">站长名字：</td>
<td width="*" height="25"><input size="30" value="<%=webceo%>" name="webceo"> 如：Zonelove</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">替换名字：</td>
<td width="*" height="25"><input size="30" value="<%=noceo%>" name="noceo"> 如：真爱空间</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">站长代码：</td>
<td width="*" height="25"><input size="30" value="<%=ceopass%>" name="ceopass"> 如：695215，建议使用数字</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">LOGO：</td>
<td width="*" height="25"><input size="30" value="<%=weblogo%>" name="weblogo"> 如：img/logo.gif</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">默认风格：</td>
<td width="*" height="25"><input size="30" value="<%=webbanner%>" name="webbanner"> 如：blue （请输入skin目录下的风格目录）</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">音乐开关：</td>
<td width="*" height="25"><input size="30" value="<%=musicoff%>" name="musicoff"> 如："1"为开,"0"为关</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">对联广告：</td>
<td width="*" height="25"><input size="30" value="<%=ad%>" name="ad"> 对联广告开关设置："1"为开,"0"为关</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">备案信息：</td>
<td width="*" height="25"><input size="30" value="<%=beian%>" name="beian"> 网站在ICP的备案信息,若没无,请留空</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">内容屏蔽：</td>
<td width="*" height="25"><textarea name="Ly_In" cols="40" rows="6" class='zonelove'><%=Ly_In%></textarea>  用于屏蔽留言和评论的不法内容，用|隔开</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">系统变量：</td>
<td width="*" height="25"><input size="30" value="<%=sessionvar%>" name="sessionvar"> 如：www.sina.com 请一定要修改，最好改乱点！</td>
</tr>

<tr class=zoneloveqs>
<td colspan="2">首页幻灯片图组显示信息</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图一的显示位置及名称:</td>
<td width="*" height="25"><input size="30" value="<%=pic1%>" name="pic1"> 默认：img/top/1.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图一的显示信息：</td>
<td width="*" height="25"><input size="30" value="<%=tit1%>" name="tit1"> 默认：我校*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图一的超级链接位置</td>
<td width="*" height="25"><input size="30" value="<%=pic1_lnk%>" name="pic1_lnk"> 默认：wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图二的显示位置及名称:</td>
<td width="*" height="25"><input size="30" value="<%=pic2%>" name="pic2"> 默认：img/top/2.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图二的显示信息：</td>
<td width="*" height="25"><input size="30" value="<%=tit2%>" name="tit2"> 默认：我校*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图二的超级链接位置</td>
<td width="*" height="25"><input size="30" value="<%=pic2_lnk%>" name="pic2_lnk"> 默认：wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图三的显示位置及名称:</td>
<td width="*" height="25"><input size="30" value="<%=pic3%>" name="pic3"> 默认：img/top/3.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图三的显示信息：</td>
<td width="*" height="25"><input size="30" value="<%=tit3%>" name="tit3"> 默认：我校*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图三的超级链接位置:</td>
<td width="*" height="25"><input size="30" value="<%=pic3_lnk%>" name="pic3_lnk"> 默认：wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图四的显示位置及名称:</td>
<td width="*" height="25"><input size="30" value="<%=pic4%>" name="pic4"> 默认：img/top/4.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图四的显示信息：</td>
<td width="*" height="25"><input size="30" value="<%=tit4%>" name="tit4"> 默认：我校*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图四的超级链接位置</td>
<td width="*" height="25"><input size="30" value="<%=pic4_lnk%>" name="pic4_lnk"> 默认：wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图五的显示位置及名称:</td>
<td width="*" height="25"><input size="30" value="<%=pic5%>" name="pic5"> 默认：img/top/5.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图五的显示信息：</td>
<td width="*" height="25"><input size="30" value="<%=tit5%>" name="tit5"> 默认：我校*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">图五的超级链接位置</td>
<td width="*" height="25"><input size="30" value="<%=pic5_lnk%>" name="pic5_lnk"> 默认：wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">首页风格样式的选择</td>
<td width="*" height="25"><input size="30" value="<%=skin%>" name="skin">"1"为显示图片,"2"为显示FLASH,若取消此栏目请留空</td>
</tr>
<tr class=zoneloveqs>
<td colspan="2" align="center"><input type="submit" value=" 更 新 数 据 "></td>
</tr>
</table>
<%
end sub
Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >inc/config.asp</font>文件内容替换成框中内容"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub
sub configok
sessionvar = replace(Trim(Request.Form("sessionvar")),CHR(34),"'")
Ly_In = replace(Trim(Request.Form("Ly_In")),CHR(34),"'")
webname = replace(Trim(Request.Form("webname")),CHR(34),"'")
musicoff = replace(Trim(Request.Form("musicoff")),CHR(34),"'")
ad = replace(Trim(Request.Form("ad")),CHR(34),"'")
beian = replace(Trim(Request.Form("beian")),CHR(34),"'")
weburl = replace(Trim(Request.Form("weburl")),CHR(34),"'")
webmail = replace(Trim(Request.Form("webmail")),CHR(34),"'")
weblogo = replace(Trim(Request.Form("weblogo")),CHR(34),"'")
webbanner = replace(Trim(Request.Form("webbanner")),CHR(34),"'")
webceo = replace(Trim(Request.Form("webceo")),CHR(34),"'")
noceo = replace(Trim(Request.Form("noceo")),CHR(34),"'")
ceopass = replace(Trim(Request.Form("ceopass")),CHR(34),"'")
indexnews = replace(Trim(Request.Form("indexnews")),CHR(34),"'")
indexarticle = replace(Trim(Request.Form("indexarticle")),CHR(34),"'")
indexsoft = replace(Trim(Request.Form("indexsoft")),CHR(34),"'")
indexdj = replace(Trim(Request.Form("indexdj")),CHR(34),"'")
indexpic = replace(Trim(Request.Form("indexpic")),CHR(34),"'")
indexjs = replace(Trim(Request.Form("indexjs")),CHR(34),"'")
indexcoolsite = replace(Trim(Request.Form("indexcoolsite")),CHR(34),"'")
indexfriendlink = replace(Trim(Request.Form("indexfriendlink")),CHR(34),"'")
indexdiary = replace(Trim(Request.Form("indexdiary")),CHR(34),"'")
newsperpage = replace(Trim(Request.Form("newsperpage")),CHR(34),"'")
diaryperpage = replace(Trim(Request.Form("diaryperpage")),CHR(34),"'")
softperpage = replace(Trim(Request.Form("softperpage")),CHR(34),"'")
articleperpage = replace(Trim(Request.Form("articleperpage")),CHR(34),"'")
picperpage = replace(Trim(Request.Form("picperpage")),CHR(34),"'")
djperpage = replace(Trim(Request.Form("djperpage")),CHR(34),"'")
coolsitesperpage = replace(Trim(Request.Form("coolsitesperpage")),CHR(34),"'")
jsperpage = replace(Trim(Request.Form("jsperpage")),CHR(34),"'")
newsoft = replace(Trim(Request.Form("newsoft")),CHR(34),"'")
toparticlenum = replace(Trim(Request.Form("toparticlenum")),CHR(34),"'")
topsoftnum = replace(Trim(Request.Form("topsoftnum")),CHR(34),"'")
topdjnum = replace(Trim(Request.Form("topdjnum")),CHR(34),"'")
topjsnum = replace(Trim(Request.Form("topjsnum")),CHR(34),"'")
toppicnum = replace(Trim(Request.Form("toppicnum")),CHR(34),"'")
topcoolsitenum = replace(Trim(Request.Form("topcoolsitenum")),CHR(34),"'")
bestdj = replace(Trim(Request.Form("bestdj")),CHR(34),"'")
bestpic = replace(Trim(Request.Form("bestpic")),CHR(34),"'")
bestjs = replace(Trim(Request.Form("bestjs")),CHR(34),"'")
bestcoolsite = replace(Trim(Request.Form("bestcoolsite")),CHR(34),"'")
adflperpage = replace(Trim(Request.Form("adflperpage")),CHR(34),"'")
skin= replace(trim(request.Form("skin")),chr(34),"'")
pic1=replace(Trim(Request.Form("pic1")),CHR(34),"'")
pic2=replace(Trim(Request.Form("pic2")),CHR(34),"'")
pic3=replace(Trim(Request.Form("pic3")),CHR(34),"'")
pic4=replace(Trim(Request.Form("pic4")),CHR(34),"'")
pic5=replace(Trim(Request.Form("pic5")),CHR(34),"'")
tit1=replace(Trim(Request.Form("tit1")),CHR(34),"'")
tit2=replace(Trim(Request.Form("tit2")),CHR(34),"'")
tit3=replace(Trim(Request.Form("tit3")),CHR(34),"'")
tit4=replace(Trim(Request.Form("tit4")),CHR(34),"'")
tit5=replace(Trim(Request.Form("tit5")),CHR(34),"'")
pic1_lnk=replace(Trim(Request.Form("pic1_lnk")),CHR(34),"'")
pic2_lnk=replace(Trim(Request.Form("pic2_lnk")),CHR(34),"'")
pic3_lnk=replace(Trim(Request.Form("pic3_lnk")),CHR(34),"'")
pic4_lnk=replace(Trim(Request.Form("pic4_lnk")),CHR(34),"'")
pic5_lnk=replace(Trim(Request.Form("pic5_lnk")),CHR(34),"'")
Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "dim sessionvar,WebName,weburl,webmail,weblogo,webbanner,webceo,noceo,ceopass" & VbCrLf
	TempStr = TempStr & "dim indexnews,indexarticle,indexsoft,indexdj,indexpic,indexjs,indexcoolsite,indexfriendlink,indexdiary" & VbCrLf
	TempStr = TempStr & "dim newsperpage,diaryperpage,softperpage,articleperpage,picperpage,djperpage,coolsitesperpage,jsperpage" & VbCrLf
	TempStr = TempStr & "dim newsoft,toparticlenum,topsoftnum,topdjnum,topjsnum,toppicnum,topcoolsitenum" & VbCrLf
	TempStr = TempStr & "dim bestdj,bestpic,bestjs,bestcoolsite,adflperpage" & VbCrLf
	TempStr = TempStr & "Dim Ly_Post,Ly_In,Ly_Inf,Ly_Xh,skin,musicoff,ad,beian" & VbCrLf
	TempStr = TempStr & "dim pic1,pic2,pic3,pic4,pic5,tit1,tit2,tit3,tit4,tit5,pic1_lnk,pic2_lnk,pic3_lnk,pic4_lnk,pic5_lnk" & vbcrlf

	TempStr = TempStr & "'=====网站基本信息=====" & VbCrLf
	TempStr = TempStr & "sessionvar="& Chr(34) & sessionvar & Chr(34) &"                   '设置变量，变量不可以为NO，否则后台无法登陆" & VbCrLf
	TempStr = TempStr & "webname="& Chr(34) & webname & Chr(34) &"                   '设置站点名称" & VbCrLf
	TempStr = TempStr & "weburl="& Chr(34) & weburl & Chr(34) &"            '设置网站地址" & VbCrLf
	TempStr = TempStr & "webmail="& Chr(34) & webmail & Chr(34) &"       '设置站长EMAIL" & VbCrLf
	TempStr = TempStr & "weblogo="& Chr(34) & weblogo & Chr(34) &"  '设置logo代码" & VbCrLf
	TempStr = TempStr & "webbanner="& Chr(34) & webbanner & Chr(34) &"                 '设置banner代码" & VbCrLf
	TempStr = TempStr & "webceo="& Chr(34) & webceo & Chr(34) &"                        '设置站长名字" & VbCrLf
	TempStr = TempStr & "noceo="& Chr(34) & noceo & Chr(34) &"                        '如果有人用你设置的站长名称留言时显示的名称" & VbCrLf
	TempStr = TempStr & "ceopass="& Chr(34) & ceopass & Chr(34) &"                        '设置站长代号" & VbCrLf
	TempStr = TempStr & "Ly_In="& Chr(34) & Ly_In & Chr(34) &"                        '设置屏蔽不法内容，用|隔开" & VbCrLf
	TempStr = TempStr & "pic1="& chr(34) & pic1 & chr(34) &"              '第一张图片" & VbCrLf
	TempStr = TempStr & "tit1="& chr(34) & tit1 & chr(34) &"              '第一张图片的标题" & VbCrLf
	TempStr = TempStr & "pic1_lnk="& chr(34) &pic1_lnk & chr(34) &"      '第一张图片的链接" & VbCrLf
	TempStr = TempStr & "pic2="& chr(34) & pic2 & chr(34)&"              '第二张图片" & VbCrLf
	TempStr = TempStr & "tit2="& chr(34) & tit2 & chr(34) &"              '第二张图片的标题" & VbCrLf
	TempStr = TempStr & "pic2_lnk="& chr(34) & pic2_lnk & chr(34) &"      '第二张图片的链接" & VbCrLf
	TempStr = TempStr & "pic3="& chr(34) & pic3 & chr(34) &"              '第三张图片" & VbCrLf
	TempStr = TempStr & "tit3="& chr(34) & tit3 & chr(34) &"              '第三张图片的标题" & VbCrLf
	TempStr = TempStr & "pic3_lnk="& chr(34) & pic3_lnk & chr(34) &"      '第三张图片链接" & VbCrLf
	TempStr = TempStr & "pic4="& chr(34) & pic4 & chr(34) &"              '第四张图片" & VbCrLf
	TempStr = TempStr & "tit4="& chr(34) & tit4 & chr(34) &"              '第四张图片的标题" & VbCrLf
	TempStr = TempStr & "pic4_lnk="& chr(34) & pic4_lnk & chr(34) &"      '第四张图片的链接" & VbCrLf
	TempStr = TempStr & "pic5="& chr(34) & pic5 & chr(34) &"              '第五张图片" & VbCrLf
	TempStr = TempStr & "tit5="& chr(34) & tit5 & chr(34) &"              '第五张图片的标题" & VbCrLf
	TempStr = TempStr & "pic5_lnk="& chr(34) & pic5_lnk & chr(34) &"      '第五张图片的链接" & VbCrLf
	TempStr = TempStr & "'=====首页显示信息=====" & VbCrLf
	TempStr = TempStr & "indexnews=10             '首页显示的校内新闻的条数" & VbCrLf
	TempStr = TempStr & "musicoff="& chr(34)& musicoff &chr(34)&"  '首页背景音乐开关" & VbCrlf 
	TempStr = TempStr & "ad="& chr(34)& ad &chr(34)&"  '首页对联广告开关" & VbCrlf 
	TempStr = TempStr & "beian="& chr(34)& beian &chr(34)&"  '首页备案信息" & VbCrlf    
	TempStr = TempStr & "skin="& chr(34)&skin &chr(34)&"         '首页风格样式选择开关" & VbCrLf
	TempStr = TempStr & "indexarticle=10          '首页显示的校园内外篇数" & VbCrLf
	TempStr = TempStr & "indexsoft=10            '首页显示的软件的个数" & VbCrLf
	TempStr = TempStr & "indexdj=10              '首页显示的音乐的个数" & VbCrLf
	TempStr = TempStr & "indexpic=20              '首页显示的照片的个数" & VbCrLf
	TempStr = TempStr & "indexjs=10               '首页显示的党政文章的个数" & VbCrLf
	TempStr = TempStr & "indexcoolsite=1          '首页显示的光 荣 榜的个数" & VbCrLf
	TempStr = TempStr & "indexfriendlink=6        '首页显示的友情链接的个数" & VbCrLf
	TempStr = TempStr & "indexdiary=5             '首页显示的公告的条数" & VbCrLf
	TempStr = TempStr & "'=====每页显示几条信息=====" & VbCrLf
	TempStr = TempStr & "newsperpage=15          '每页显示的校内新闻的条数" & VbCrLf
	TempStr = TempStr & "diaryperpage=1           '显示的每日碎语的篇数" & VbCrLf
	TempStr = TempStr & "softperpage=10           '每页显示的软件个数" & VbCrLf
	TempStr = TempStr & "articleperpage=20       '每页显示校园内外的篇数" & VbCrLf
	TempStr = TempStr & "picperpage=12            '每页显示的照片的个数 " & VbCrLf
	TempStr = TempStr & "djperpage=25             '每页显示的音乐的个数" & VbCrLf
	TempStr = TempStr & "coolsitesperpage=5       '每页显示的光荣榜上的个数" & VbCrLf
	TempStr = TempStr & "jsperpage=20             '每页显示的文章个数" & VbCrLf
	TempStr = TempStr & "'=====显示热门信息=====" & VbCrLf
	TempStr = TempStr & "newsoft=10              '显示最新的程序个数" & VbCrLf
	TempStr = TempStr & "toparticlenum=10         '显示阅读次数最多的X篇文章" & VbCrLf
	TempStr = TempStr & "topsoftnum=10           '显示下载次数最多的X个软件" & VbCrLf
	TempStr = TempStr & "topdjnum=10             '显示点击次数最多的X首音乐" & VbCrLf
	TempStr = TempStr & "topjsnum=10              '显示点击次数最多的X篇文章" & VbCrLf
	TempStr = TempStr & "toppicnum=10              '显示点击次数最多的X张图片" & VbCrLf
	TempStr = TempStr & "topcoolsitenum=2         '显示点击次数最多的X个网站" & VbCrLf
	TempStr = TempStr & "'=====显示推荐信息=====" & VbCrLf
	TempStr = TempStr & "bestdj=10                '显示推荐音乐的个数" & VbCrLf
	TempStr = TempStr & "bestpic=10                '显示推荐照片的个数" & VbCrLf
	TempStr = TempStr & "bestjs=10                '显示推荐文章的个数" & VbCrLf
	TempStr = TempStr & "bestcoolsite=2           '显示光 荣 榜的个数" & VbCrLf
	TempStr = TempStr & "adflperpage=1           '后台连接每页显示数量" & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"inc/config.asp"
	If GBL_CHK_TempStr = "" Then
Response.Write "<table width=""98%"" align=""center"" border=""1"" cellspacing=""0"" cellpadding=""4"" class=zonelovebk style=""border-collapse: collapse""><tr><td class=zonelovess>基本资料更新</td></tr><tr class=zoneloveds><td align=""center"" height=""66"">&gt;更新成功&lt;</td></tr></table>"
Else
		%><table width=""98%"" align=""center"" border=""1"" cellspacing=""0"" cellpadding=""4"" class=zonelovebk style=""border-collapse: collapse""><tr><td class=zonelovess>基本资料更新</td></tr><tr class=zoneloveds><td align=""center"" height=""66"">&gt;<%=GBL_CHK_TempStr%>&lt;<br><br>
		<textarea name="fileContent" cols="80" rows="30"><%=Server.htmlencode(TempStr)%></textarea></td></tr></table><%
		GBL_CHK_TempStr = ""
	End If
End Sub
end if%>