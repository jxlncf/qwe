<!--#include file="inc/config.asp"-->
<%
if session("adminlogin")<>sessionvar then
  founderr=true
Response.Write("<script language=javascript>alert('����δ��½���½��ʱ�������µ�½!��');this.top.location.href='admin.asp';</script>")
  response.end
else%>
<HTML><HEAD><TITLE>��������</TITLE>
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
<td colspan="2" height="25">��վ��������</td>
</tr>
<tr class=zoneloveqs>
<td colspan="2">��ҳ������վ��ҳ�Ļ�����Ϣ�����Ի�ѡ��</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">��վ���ƣ�</td>
<td width="*" height="25"><input size="30" name="webname" value="<%=webname%>"> �磺�氮�ռ�У��</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">��վ��ַ��</td>
<td width="*" height="25"><input size="30" value="<%=weburl%>" name="weburl"> �磺http://www.zonelove.net</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">վ�����䣺</td>
<td width="*" height="25"><input size="30" value="<%=webmail%>" name="webmail"> �磺zonelove8@163.com</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">վ�����֣�</td>
<td width="*" height="25"><input size="30" value="<%=webceo%>" name="webceo"> �磺Zonelove</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">�滻���֣�</td>
<td width="*" height="25"><input size="30" value="<%=noceo%>" name="noceo"> �磺�氮�ռ�</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">վ�����룺</td>
<td width="*" height="25"><input size="30" value="<%=ceopass%>" name="ceopass"> �磺695215������ʹ������</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">LOGO��</td>
<td width="*" height="25"><input size="30" value="<%=weblogo%>" name="weblogo"> �磺img/logo.gif</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">Ĭ�Ϸ��</td>
<td width="*" height="25"><input size="30" value="<%=webbanner%>" name="webbanner"> �磺blue ��������skinĿ¼�µķ��Ŀ¼��</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">���ֿ��أ�</td>
<td width="*" height="25"><input size="30" value="<%=musicoff%>" name="musicoff"> �磺"1"Ϊ��,"0"Ϊ��</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">������棺</td>
<td width="*" height="25"><input size="30" value="<%=ad%>" name="ad"> ������濪�����ã�"1"Ϊ��,"0"Ϊ��</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">������Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=beian%>" name="beian"> ��վ��ICP�ı�����Ϣ,��û��,������</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">�������Σ�</td>
<td width="*" height="25"><textarea name="Ly_In" cols="40" rows="6" class='zonelove'><%=Ly_In%></textarea>  �����������Ժ����۵Ĳ������ݣ���|����</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ϵͳ������</td>
<td width="*" height="25"><input size="30" value="<%=sessionvar%>" name="sessionvar"> �磺www.sina.com ��һ��Ҫ�޸ģ���ø��ҵ㣡</td>
</tr>

<tr class=zoneloveqs>
<td colspan="2">��ҳ�õ�Ƭͼ����ʾ��Ϣ</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼһ����ʾλ�ü�����:</td>
<td width="*" height="25"><input size="30" value="<%=pic1%>" name="pic1"> Ĭ�ϣ�img/top/1.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼһ����ʾ��Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=tit1%>" name="tit1"> Ĭ�ϣ���У*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼһ�ĳ�������λ��</td>
<td width="*" height="25"><input size="30" value="<%=pic1_lnk%>" name="pic1_lnk"> Ĭ�ϣ�wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ������ʾλ�ü�����:</td>
<td width="*" height="25"><input size="30" value="<%=pic2%>" name="pic2"> Ĭ�ϣ�img/top/2.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ������ʾ��Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=tit2%>" name="tit2"> Ĭ�ϣ���У*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ���ĳ�������λ��</td>
<td width="*" height="25"><input size="30" value="<%=pic2_lnk%>" name="pic2_lnk"> Ĭ�ϣ�wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ������ʾλ�ü�����:</td>
<td width="*" height="25"><input size="30" value="<%=pic3%>" name="pic3"> Ĭ�ϣ�img/top/3.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ������ʾ��Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=tit3%>" name="tit3"> Ĭ�ϣ���У*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ���ĳ�������λ��:</td>
<td width="*" height="25"><input size="30" value="<%=pic3_lnk%>" name="pic3_lnk"> Ĭ�ϣ�wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ�ĵ���ʾλ�ü�����:</td>
<td width="*" height="25"><input size="30" value="<%=pic4%>" name="pic4"> Ĭ�ϣ�img/top/4.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ�ĵ���ʾ��Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=tit4%>" name="tit4"> Ĭ�ϣ���У*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ�ĵĳ�������λ��</td>
<td width="*" height="25"><input size="30" value="<%=pic4_lnk%>" name="pic4_lnk"> Ĭ�ϣ�wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ�����ʾλ�ü�����:</td>
<td width="*" height="25"><input size="30" value="<%=pic5%>" name="pic5"> Ĭ�ϣ�img/top/5.jpg</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ�����ʾ��Ϣ��</td>
<td width="*" height="25"><input size="30" value="<%=tit5%>" name="tit5"> Ĭ�ϣ���У*******</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">ͼ��ĳ�������λ��</td>
<td width="*" height="25"><input size="30" value="<%=pic5_lnk%>" name="pic5_lnk"> Ĭ�ϣ�wab.asp</td>
</tr>
<tr class=zoneloveds>
<td height="25" width="80">��ҳ�����ʽ��ѡ��</td>
<td width="*" height="25"><input size="30" value="<%=skin%>" name="skin">"1"Ϊ��ʾͼƬ,"2"Ϊ��ʾFLASH,��ȡ������Ŀ������</td>
</tr>
<tr class=zoneloveqs>
<td colspan="2" align="center"><input type="submit" value=" �� �� �� �� "></td>
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
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ�������ʹ��FTP�ȹ��ܣ���<font color=Red >inc/config.asp</font>�ļ������滻�ɿ�������"
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

	TempStr = TempStr & "'=====��վ������Ϣ=====" & VbCrLf
	TempStr = TempStr & "sessionvar="& Chr(34) & sessionvar & Chr(34) &"                   '���ñ���������������ΪNO�������̨�޷���½" & VbCrLf
	TempStr = TempStr & "webname="& Chr(34) & webname & Chr(34) &"                   '����վ������" & VbCrLf
	TempStr = TempStr & "weburl="& Chr(34) & weburl & Chr(34) &"            '������վ��ַ" & VbCrLf
	TempStr = TempStr & "webmail="& Chr(34) & webmail & Chr(34) &"       '����վ��EMAIL" & VbCrLf
	TempStr = TempStr & "weblogo="& Chr(34) & weblogo & Chr(34) &"  '����logo����" & VbCrLf
	TempStr = TempStr & "webbanner="& Chr(34) & webbanner & Chr(34) &"                 '����banner����" & VbCrLf
	TempStr = TempStr & "webceo="& Chr(34) & webceo & Chr(34) &"                        '����վ������" & VbCrLf
	TempStr = TempStr & "noceo="& Chr(34) & noceo & Chr(34) &"                        '��������������õ�վ����������ʱ��ʾ������" & VbCrLf
	TempStr = TempStr & "ceopass="& Chr(34) & ceopass & Chr(34) &"                        '����վ������" & VbCrLf
	TempStr = TempStr & "Ly_In="& Chr(34) & Ly_In & Chr(34) &"                        '�������β������ݣ���|����" & VbCrLf
	TempStr = TempStr & "pic1="& chr(34) & pic1 & chr(34) &"              '��һ��ͼƬ" & VbCrLf
	TempStr = TempStr & "tit1="& chr(34) & tit1 & chr(34) &"              '��һ��ͼƬ�ı���" & VbCrLf
	TempStr = TempStr & "pic1_lnk="& chr(34) &pic1_lnk & chr(34) &"      '��һ��ͼƬ������" & VbCrLf
	TempStr = TempStr & "pic2="& chr(34) & pic2 & chr(34)&"              '�ڶ���ͼƬ" & VbCrLf
	TempStr = TempStr & "tit2="& chr(34) & tit2 & chr(34) &"              '�ڶ���ͼƬ�ı���" & VbCrLf
	TempStr = TempStr & "pic2_lnk="& chr(34) & pic2_lnk & chr(34) &"      '�ڶ���ͼƬ������" & VbCrLf
	TempStr = TempStr & "pic3="& chr(34) & pic3 & chr(34) &"              '������ͼƬ" & VbCrLf
	TempStr = TempStr & "tit3="& chr(34) & tit3 & chr(34) &"              '������ͼƬ�ı���" & VbCrLf
	TempStr = TempStr & "pic3_lnk="& chr(34) & pic3_lnk & chr(34) &"      '������ͼƬ����" & VbCrLf
	TempStr = TempStr & "pic4="& chr(34) & pic4 & chr(34) &"              '������ͼƬ" & VbCrLf
	TempStr = TempStr & "tit4="& chr(34) & tit4 & chr(34) &"              '������ͼƬ�ı���" & VbCrLf
	TempStr = TempStr & "pic4_lnk="& chr(34) & pic4_lnk & chr(34) &"      '������ͼƬ������" & VbCrLf
	TempStr = TempStr & "pic5="& chr(34) & pic5 & chr(34) &"              '������ͼƬ" & VbCrLf
	TempStr = TempStr & "tit5="& chr(34) & tit5 & chr(34) &"              '������ͼƬ�ı���" & VbCrLf
	TempStr = TempStr & "pic5_lnk="& chr(34) & pic5_lnk & chr(34) &"      '������ͼƬ������" & VbCrLf
	TempStr = TempStr & "'=====��ҳ��ʾ��Ϣ=====" & VbCrLf
	TempStr = TempStr & "indexnews=10             '��ҳ��ʾ��У�����ŵ�����" & VbCrLf
	TempStr = TempStr & "musicoff="& chr(34)& musicoff &chr(34)&"  '��ҳ�������ֿ���" & VbCrlf 
	TempStr = TempStr & "ad="& chr(34)& ad &chr(34)&"  '��ҳ������濪��" & VbCrlf 
	TempStr = TempStr & "beian="& chr(34)& beian &chr(34)&"  '��ҳ������Ϣ" & VbCrlf    
	TempStr = TempStr & "skin="& chr(34)&skin &chr(34)&"         '��ҳ�����ʽѡ�񿪹�" & VbCrLf
	TempStr = TempStr & "indexarticle=10          '��ҳ��ʾ��У԰����ƪ��" & VbCrLf
	TempStr = TempStr & "indexsoft=10            '��ҳ��ʾ������ĸ���" & VbCrLf
	TempStr = TempStr & "indexdj=10              '��ҳ��ʾ�����ֵĸ���" & VbCrLf
	TempStr = TempStr & "indexpic=20              '��ҳ��ʾ����Ƭ�ĸ���" & VbCrLf
	TempStr = TempStr & "indexjs=10               '��ҳ��ʾ�ĵ������µĸ���" & VbCrLf
	TempStr = TempStr & "indexcoolsite=1          '��ҳ��ʾ�Ĺ� �� ��ĸ���" & VbCrLf
	TempStr = TempStr & "indexfriendlink=6        '��ҳ��ʾ���������ӵĸ���" & VbCrLf
	TempStr = TempStr & "indexdiary=5             '��ҳ��ʾ�Ĺ��������" & VbCrLf
	TempStr = TempStr & "'=====ÿҳ��ʾ������Ϣ=====" & VbCrLf
	TempStr = TempStr & "newsperpage=15          'ÿҳ��ʾ��У�����ŵ�����" & VbCrLf
	TempStr = TempStr & "diaryperpage=1           '��ʾ��ÿ�������ƪ��" & VbCrLf
	TempStr = TempStr & "softperpage=10           'ÿҳ��ʾ���������" & VbCrLf
	TempStr = TempStr & "articleperpage=20       'ÿҳ��ʾУ԰�����ƪ��" & VbCrLf
	TempStr = TempStr & "picperpage=12            'ÿҳ��ʾ����Ƭ�ĸ��� " & VbCrLf
	TempStr = TempStr & "djperpage=25             'ÿҳ��ʾ�����ֵĸ���" & VbCrLf
	TempStr = TempStr & "coolsitesperpage=5       'ÿҳ��ʾ�Ĺ��ٰ��ϵĸ���" & VbCrLf
	TempStr = TempStr & "jsperpage=20             'ÿҳ��ʾ�����¸���" & VbCrLf
	TempStr = TempStr & "'=====��ʾ������Ϣ=====" & VbCrLf
	TempStr = TempStr & "newsoft=10              '��ʾ���µĳ������" & VbCrLf
	TempStr = TempStr & "toparticlenum=10         '��ʾ�Ķ���������Xƪ����" & VbCrLf
	TempStr = TempStr & "topsoftnum=10           '��ʾ���ش�������X�����" & VbCrLf
	TempStr = TempStr & "topdjnum=10             '��ʾ�����������X������" & VbCrLf
	TempStr = TempStr & "topjsnum=10              '��ʾ�����������Xƪ����" & VbCrLf
	TempStr = TempStr & "toppicnum=10              '��ʾ�����������X��ͼƬ" & VbCrLf
	TempStr = TempStr & "topcoolsitenum=2         '��ʾ�����������X����վ" & VbCrLf
	TempStr = TempStr & "'=====��ʾ�Ƽ���Ϣ=====" & VbCrLf
	TempStr = TempStr & "bestdj=10                '��ʾ�Ƽ����ֵĸ���" & VbCrLf
	TempStr = TempStr & "bestpic=10                '��ʾ�Ƽ���Ƭ�ĸ���" & VbCrLf
	TempStr = TempStr & "bestjs=10                '��ʾ�Ƽ����µĸ���" & VbCrLf
	TempStr = TempStr & "bestcoolsite=2           '��ʾ�� �� ��ĸ���" & VbCrLf
	TempStr = TempStr & "adflperpage=1           '��̨����ÿҳ��ʾ����" & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"inc/config.asp"
	If GBL_CHK_TempStr = "" Then
Response.Write "<table width=""98%"" align=""center"" border=""1"" cellspacing=""0"" cellpadding=""4"" class=zonelovebk style=""border-collapse: collapse""><tr><td class=zonelovess>�������ϸ���</td></tr><tr class=zoneloveds><td align=""center"" height=""66"">&gt;���³ɹ�&lt;</td></tr></table>"
Else
		%><table width=""98%"" align=""center"" border=""1"" cellspacing=""0"" cellpadding=""4"" class=zonelovebk style=""border-collapse: collapse""><tr><td class=zonelovess>�������ϸ���</td></tr><tr class=zoneloveds><td align=""center"" height=""66"">&gt;<%=GBL_CHK_TempStr%>&lt;<br><br>
		<textarea name="fileContent" cols="80" rows="30"><%=Server.htmlencode(TempStr)%></textarea></td></tr></table><%
		GBL_CHK_TempStr = ""
	End If
End Sub
end if%>