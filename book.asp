<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<!--#include file="mdb.asp"-->
<!--#include file="inc/inc.asp"-->
<!--#include file="inc/config.asp"-->
<%
start="���Ա�"
Ly_Inf = split(Ly_In,"|")
If Request.Form<>"" Then
	For Each Ly_Post In Request.Form
		For Ly_Xh=0 To Ubound(Ly_Inf)
			If Instr(LCase(Request.Form(Ly_Post)),Ly_Inf(Ly_Xh))<>0 Then
				Response.Write "<Script Language=JavaScript>alert('�Բ���,������������а���ϵͳ����ֹ�ַ�,���ʵ���ٷ���');javascript :history.back();</Script>"
				Response.End
			End If
		Next
	Next
End If
dim founderr
founderr=false
if request("page")<>"" then
  if not isInteger(request("page")) then
    founderr=true
    Response.Write "<script language=javascript>alert('�����Ƿ���');javascript:history.back();</script>"
  end if
end if
if request("keyword")<>"" then
  if instr(request("keyword"),"'")>0 then
    founderr=true
    Response.Write "<script language=javascript>alert('���������Ƿ�');javascript:history.back();</script>"
  end if
end if
page = Request("page")
action = Request.QueryString("action")
action_e = Request.Form("action_e")
call head
call menu
Response.Write "<TABLE id=navsub cellSpacing=0 cellPadding=0 align=center><TBODY><TR><TD class=l></TD><TD height=""23"" align=""right"">"
call Main_Menu()
Response.Write "&nbsp;</TD><TD class=r></TD></TR></TBODY></TABLE>" & vbCrLf
Response.Write "<TABLE id=middle cellSpacing=0 cellPadding=0 align=center boder=""0"" height=300>" & vbCrLf
Response.Write "  <TBODY>" & vbCrLf
Response.Write "  <TR vAlign=top align=left>" & vbCrLf
Response.Write "    <TD width=200>" & vbCrLf
Response.Write "<DIV class=member>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""�� �� ˵ ��"",""ml"");</SCRIPT>" & vbCrLf
Response.Write "<div style=""LINE-HEIGHT: 180%"">" & vbCrLf
Response.Write "���������ϵ��£������л����񹲺͹��ĸ����йط��ɷ���" & vbCrLf
Response.Write "<br>���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·�������" & vbCrLf
Response.Write "<br>����վ������Ա��Ȩ������ɾ�����Ͻ�����е���������" & vbCrLf
Response.Write "<br>����վ��Ȩ����վ��ת�ػ�������������" & vbCrLf
Response.Write "<br>�����뱾���Լ��������Ѿ��Ķ���������������" & vbCrLf	
Response.Write "</div>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetabletop(""�� �� �� ��"",""ml"");</SCRIPT>" & vbCrLf
Response.Write "<form name=""form2"" method=""post"" action=""book.asp""><div align=center><input type=""radio"" name=""select"" value=""name"" checked>����&nbsp;<input type=""radio"" name=""select"" value=""words"">����&nbsp;<input type=""radio"" name=""select"" value=""reply"">�ظ�<br><input type='text' name='keyword'  size='15' value='�����ؼ���' maxlength='50' onFocus='this.select();' class=""zonelove"">&nbsp;<input type='submit' name='search'  value='����' onmouseover=""this.className='boton'"" onmouseout=""this.className='botoff'"" class=""botoff"">" & vbCrLf
Response.Write "</div></form>" & vbCrLf
Response.Write "<SCRIPT type=text/javascript>zonelovetablebottom(""mr"");</SCRIPT>"
Response.Write "</TD>" & vbCrLf
Response.Write "<TD>" & vbCrLf
Response.Write "<DIV class=mframe>" & vbCrLf
%><SCRIPT type=text/javascript>zonelovetable("<%if request("keyword")<>"" then%>�������<%else%>���Ա�<%end if%>");</SCRIPT><%
Response.Write "<table width=""100%"" border=""1"" align=""center"" cellspacing=""0"" cellpadding=""3"" bgcolor=""#FFFFFF"" bordercolor=""#f0f0f0"" style=""border-collapse: collapse"" frame=lhs>"
Select Case action_e
	Case ""
	Case "Add_New"
if int(request("VerifyCode"))<>int(Session("GetCode")) then
 Response.Write("<script language=javascript>alert('��������ȷ����֤�룡');window.document.location.href='book.asp?action=Add_New';</script>") 
Response.End 
end if 
Call Add_New_Execute()
	Case "Edit"
if session("adminlogin")<>sessionvar then
Response.write "<script language = 'javascript'>alert('����δ��½����,�޷����б༭����!');window.document.location.href='book.asp';</script>"
else
		Call Edit_Execute()
end if

End Select

Select Case action		
	Case ""		
		Call View_Words()		
	Case "Add_New"
		Call Add_New()
	Case "Edit"
		Call Edit()
	Case "View_Words"	
		Call View_Words()	
	Case "Delete"
if session("adminlogin")<>sessionvar then
Response.write "<script language = 'javascript'>alert('����δ��½����,�޷�ɾ������!');window.document.location.href='book.asp';</script>"
else
		Call Delete()
		Call View_Words()
end if				
End Select
Call footer()
Sub Main_Menu()
Response.Write "<a href='?action=Add_New'>���������</a>&nbsp;<a href='?action=View_Words'>���鿴����</a>&nbsp;"
End Sub 
Sub View_Words() 
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Sql="Select * From words Order By date Desc"
                if request("keyword")<>"" then
                sql="select * from words where "&request("select")&" like '%"&request("keyword")&"%'order by date DESC"
                 end if
		Rs.Open Sql,Conn,1,1
		TotalRecord=Rs.RecordCount
if not rs.eof then
		Rs.PageSize = 10
		PageSize = Rs.PageSize
		PageCount=Rs.PageCount
                if page="" Then  
			Page = 1
		Else
			Rs.AbsolutePage = page
		End If
i=0
    do while i < PageSize And not rs.eof 
NO=TotalRecord-i-(PageSize*Page)+PageSize
%><tr bgcolor="#ffffff"><td colspan="2" height="10"></td></tr><tr bgcolor="#ffffff"><td rowspan=2 width=90 align="center" valign="top"><%If Rs("qq")<>"" Then%><a target="_blank" href="http://qqshow.qq.com/cgi-bin/qqshow_user_info?uin=<%=Rs("qq")%>"><img width="70" height="113" border="0" src="http://qqshow-user.tencent.com/<%=Rs("qq")%>/11/00/"></a><%else%><img src=img/<%=Rs("sex")%>.gif><%End If%><br><a href="book.asp?select=name&keyword=<%=Rs("name")%>" title='����[<%=Rs("name")%>]�������������'><font class="book"><%=Rs("name")%></font></a></td><td height=20><table border=0 cellpadding=0 cellspacing=0 width="100%"><tbody><tr><td valign=bottom width="40%"><%=NO%>¥��<%=Rs("date")%></td><td align=right width="60%"><%If Rs("city")<>"" Then%><img align=absBottom title="<%=Rs("name")%>����<%=Rs("city")%>" border=0 height=16 src="img/book_city.gif" width=16>&nbsp;<%End If%><%If Rs("email")<>"" Then%><a href="mailto:<%=Rs("email")%>"><img align=absBottom border=0 height=16 src="img/book_email.gif" width=16 title="��<%=Rs("name")%>д��"></a>&nbsp;<%End If%><%If Rs("qq")<>"" Then%><img align=absBottom title="<%=Rs("name")%>��QQ������<%=Rs("qq")%>" border=0 height=16 src="img/book_qq.gif" width=16>&nbsp;<%End If%><%If Rs("uc")<>"" Then%><img align=absBottom title="<%=Rs("name")%>��uc������<%=Rs("uc")%>" border=0 height=16 src="img/book_qq.gif" width=16>&nbsp;<%End If%><%If Rs("web")<>"" Then%><a href="<%=Rs("web")%>" target="_blank"><img align=absBottom border=0 height=16 src="img/book_web.gif" width=16 title="ȥ<%=Rs("name")%>����ҳ����"></a>&nbsp;<%End If%><img align=absBottom title="<%=Rs("name")%>��IP��ַΪ<%=Rs("ip")%>" height=16 src="img/book_ip.gif" width=16>&nbsp;<% If session("adminlogin")=sessionvar Then %><a href="book.asp?action=Delete&id=<%=Rs("id")%>"><img align=absBottom border=0 title="��ɾ����" height=16 src="img/book_Delete.gif" width=16></a>&nbsp;<a href="book.asp?action=Edit&id=<%=Rs("id")%>"><img align=absBottom border=0 title="����˻ظ���" height=16 src="img/book_edit.gif" width=16></a>&nbsp;<%End if%></td></tr></tbody></table></td></tr><tr><td bgcolor="#ffffff" height=80><table border=0 cellpadding=0 cellspacing=2 width=100% height=100% style='table-layout:fixed; word-break:break-all; line-height:150%'><tbody><tr><td><%If Rs("admin")="0" Then%>[<%if Rs("title")="1" then%><font color=#0000FF>����</font><%elseif Rs("title")="2" then%><font color=#FF00FF>����</font><%elseif Rs("title")="3" then%><font color=#FF7F50>����</font><%elseif Rs("title")="4" then%><font color=#228B22>����</font><%elseif Rs("title")="5" then%><font color=#1E90FF>����</font><%End If%>]<%=Ubb(unHtml(Rs("words")))%><%elseif session("adminlogin")=sessionvar Then %><center><font color=#1E90FF>������������ ����������������� �ԩ���������</font></center>[<%if Rs("title")="1" then%><font color=#0000FF>����</font><%elseif Rs("title")="2" then%><font color=#FF00FF>����</font><%elseif Rs("title")="3" then%><font color=#FF7F50>����</font><%elseif Rs("title")="4" then%><font color=#228B22>����</font><%elseif Rs("title")="5" then%><font color=#1E90FF>����</font><%End If%>]<%=Ubb(unHtml(Rs("words")))%><%else%><center><font color=#1E90FF>������������ ����������������� �ԩ���������</font></center><%End If%></td></tr><%If Rs("reply")<>"" Then%><tr><td valign="bottom"><div style="margin:0px 3px;border:1px solid #CCCCCC;padding:5px;background:#FDFDDF ;line-height:150%;" ><font color=#ccccccc><b><%=webceo%>�ظ��ڣ�</b><%=Rs("redate")%></font><br><%=Ubb(unHtml(Rs("reply")))%></div></td></tr><%End If%></table></td></tr><%
		 
		rs.movenext
		i=i+1
    	loop
else
if rs.eof and rs.bof then
%>
<tr><td align=middle height="80" colSpan=2><%if request("keyword")<>"" then%>û���ҵ�����[<b><font color=red><%=request("keyword")%></font></b>]�����ԣ�<%else%>û���κ����ԣ���ӭ����Ϊ��վ��һ�������ߣ�<%end if%></td></tr>
<%Response.Write "<tr align=""center""><td bgcolor=""#F0f0f0"" height=""1"" colspan=""2""></td></tr></TABLE></td></tr></TABLE>"
end if
end if
		Rs.Close
		Set Rs = Nothing
dim n
n= TotalRecord \ PageSize
Response.Write "<tr bgcolor=""#ffffff""><td colspan=""2"" height=""10""></td></tr>" & vbCrLf
Response.Write "<TR bgColor=#FFFFFF><TD colspan=""2"">" & vbCrLf
Response.Write "<TABLE width=""100%"" align=center border=0>" & vbCrLf
Response.Write "<TBODY>" & vbCrLf
Response.Write "<TR>" & vbCrLf
Response.Write "<TD align=middle width=""35%"" height=25><IMG height=14 src=""img/so.gif"" width=14 align=absMiddle> ��[<font color=""#FF6666"">"&TotalRecord&"</font>]�����ԡ���" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "[<font color=""#FF6666"">"&n+1&"</font>]ҳ</TD>" & vbCrLf
Response.Write "<TD width=""40%"" align=middle><IMG height=11 src=""img/lt.gif"" width=11 align=absMiddle>"
if Page<2 then
response.write "<font color='999966'>��ҳ ��һҳ</font> "
else
response.write "<a href=book.asp?select="&request("select")&"&keyword="&request("keyword")&">��ҳ</a> "
response.write "<a href=book.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&Page-1&">��һҳ</a> "
end if
if n-page<0 then
response.write "<font color='999966'>��һҳ βҳ</font>"
else
response.write "<a href=book.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&(Page+1)
response.write ">��һҳ</a> <a href=book.asp?select="&request("select")&"&keyword="&request("keyword")&"&page="&n+1&">βҳ</a>"
end if
Response.Write "<IMG height=11 src=""img/rt.gif"" width=11 align=absMiddle></TD>" & vbCrLf
Response.Write "<form name=""form1"" method=""post"" action=""book.asp?select="&request("select")&"&keyword="&request("keyword")&"""><TD align=middle width=""25%"">" & vbCrLf
Response.Write "<select name=""page"" class=""zonelove"">"
For p = 1 To PageCount
if request("page")=cstr(p) then
%>
<option value=<%=p%> selected>��<%=p%>ҳ</option> 
<%else%>
<option value=<%=p%>>��<%=p%>ҳ</option> 
<%end if
Next
Response.Write "</select><input type=""submit"" name=""Submit2"" value=""ת��"" onmouseover=""this.className='boton'"" onmouseout=""this.className='botoff'"" class=""botoff""></TD></FORM>" & vbCrLf
Response.Write "</TR>" & vbCrLf
Response.Write "</TABLE>" & vbCrLf
Response.Write "<tr align=""center""><td bgcolor=""#F0f0f0"" height=""1"" colspan=""2""></td></tr>" & vbCrLf
Response.Write "</td></tr></table><br></td></tr></table>"
End Sub
Sub Add_New()
Response.Write "<form name=""form"" method=""post"" action=""book.asp"" onsubmit=""return checkBook();"">" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td width=""80"">��������:</td>" & vbCrLf
Response.Write "<td width=""300""><input type=""text"" name=""name"" maxlength=""14"" size=""20"" class='zonelove' autocomplete=""off"">   <font color=red>*</font></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td>�����Ա�:</td>" & vbCrLf
Response.Write "<td><input type=""radio"" name=""SEX"" value=""boy"" checked>˧�� <input type=""radio"" name=""SEX"" value=""girl"">��Ů</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td>�������䣺</td>" & vbCrLf
Response.Write "<td><input type=""text"" name=""email"" maxlength=""30"" size=""20"" class='zonelove'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td>��Ѹ QQ:</td>" & vbCrLf
Response.Write "<td><input type=""text"" name=""qq"" maxlength=""15"" size=""20"" class='zonelove'> ��д���������ʾ���QQ��</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td>���� UC:</td>" & vbCrLf
Response.Write "<td><input type=""text"" name=""uc"" maxlength=""15"" size=""20"" class='zonelove'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td>������ҳ:</td>" & vbCrLf
Response.Write "<td><input type=""text"" name=""web"" maxlength=""40"" size=""20"" class='zonelove'> �������http://</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr> <td>��������:</td>" & vbCrLf
Response.Write "<td><input type=""text"" name=""city"" maxlength=""40"" size=""20"" class='zonelove'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr><td>����ѡ��:</td>" & vbCrLf
Response.Write "<td><input type=""radio"" name=""title"" value=""1"" checked><font color=#0000FF>����</font> <input type=""radio"" name=""title"" value=""2""><font color=#FF00FF>����" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "</font> <input type=""radio"" name=""title"" value=""3""><font color=#FF7F50>����</font> <input type=""radio"" name=""title"" value=""4""><font color=#228B22>����</font>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<input type=""radio"" name=""title"" value=""5""><font color=#1E90FF>����</font></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td valign=""top"">��������: <br><font color=red>*֧��UBB</font></td>" & vbCrLf
Response.Write "<td  valign=""bottom""><textarea name=""words"" cols=""40"" rows=""6"" class='zonelove' onkeydown=""bookshowLen(this)"" onkeyup=""bookshowLen(this)""></textarea>&nbsp;������<input type='text' value='0' id='wordsLen' size='2' style='border-width:0;background:transparent;'><font title='���ڻ�С��255���֣�'>��255</font></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td valign=""middle"">��֤�룺</td>" & vbCrLf
Response.Write "<td valign=""top"">" & vbCrLf
Response.Write "<input name=""VerifyCode"" type=""text"" size=""10""> &nbsp;<img src=""inc/lyCode.asp""> <input type=""radio"" name=""admin"" value=""1"" checked></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td align=""center""  height=""40"" colspan=""2"">" & vbCrLf
Response.Write "<input type=""hidden"" name=""action_e"" value=""Add_New""> <input type=""submit"" name=""Submit"" value=""�ύ"" onmouseover=""this.className='boton'""" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "onmouseout=""this.className='botoff'"" class=""botoff"">" & vbCrLf
Response.Write "        <input type=""reset"" name=""Submit2"" value=""��д"" onmouseover=""this.className='boton'"" onmouseout=""this.className='botoff'"" class=""botoff"">" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</form>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</table>"
End Sub
Sub Edit()
Set Rs = Server.CreateObject("ADODB.RecordSet")
Sql="Select * From words Where id="&Request.QueryString("id")
Rs.Open Sql,Conn,1,1
Response.Write "<form name=""reply"" method=""post"" action=""book.asp"">" & vbCrLf
Response.Write "<tr bgcolor=""#ffffff"">" & vbCrLf
Response.Write " <TD align=""center"">������������</TD></tr>" & vbCrLf
Response.Write "<tr bgcolor=""#ffffff"">" & vbCrLf
Response.Write "<td align=""center""><textarea name=""words"" cols=""80"" rows=""5"">"&Rs("words")&"</textarea></td></tr>" & vbCrLf
Response.Write "<tr bgcolor=""#ffffff"">" & vbCrLf
Response.Write " <TD align=""center"">�ظ�����</TD></tr>" & vbCrLf
Response.Write "<tr bgcolor=""#ffffff"">" & vbCrLf
Response.Write "<td align=""center""><textarea name=""reply"" cols=""80"" rows=""5"" class='zonelove'>"&Rs("reply")&"</textarea></td></tr>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td valign=""middle"">�Ƿ����أ�" & vbCrLf
Response.Write "<input type=""radio"" name=""admin"" value=""0"""
If Rs("admin")="0" Then
Response.Write "checked" & vbCrLf
end if
Response.Write "> ��" & vbCrLf
Response.Write "<input type=""radio"" name=""admin"" value=""1"""
If Rs("admin")="1" Then
Response.Write "checked" & vbCrLf
end if
Response.Write "> ��"
Response.Write "&nbsp;&nbsp;<font color=#009900>*</font> ѡ�����غ󣬴�����ֻ�й���Ա���Կ�����</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr align=""center""><td colspan=""2""><input type=""hidden"" name=""action_e"" value=""Edit"">" & vbCrLf
Response.Write "        <input type=""hidden"" name=""id"" value="""&Request.QueryString("id")&""">" & vbCrLf
Response.Write "        <input type=""submit"" name=""Submit"" value=""ȷ��"" id=""Submit"" onmouseover=""this.className='boton'"" onmouseout=""this.className='botoff'"" class=""botoff""></td></tr>" & vbCrLf
Response.Write "</form>" & vbCrLf
Response.Write "<tr align=""center""><td bgcolor=""#F0f0f0"" height=""1""></td></tr>" & vbCrLf
Response.Write "</table><br>" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</table>"
End Sub 
Sub Add_New_Execute()
	If Request.Form("name")="" Then
	Response.Write "<script language=javascript>alert('��������Ϊ�գ�');javascript:history.back();</script>"
	Response.End
	End If
	If Len(Request.Form("name"))>10 Then
	Response.Write "<script language=javascript>alert('�������̫���ˣ�');javascript:history.back();</script>"
	Response.End
	End If
	If Request.Form("email")<>"" Then
	If instr(Request.Form("email"),"@")=0 or instr(Request.Form("email"),"@")=1 or instr(Request.Form("email"),"@")=len(email) then
	Response.Write "<script language=javascript>alert('���������ʽ��д����ȷ��');javascript:history.back();</script>"
	Response.End
	End If
	End If
        If Request.Form("qq")<>"" Then
	If Len(Request.Form("qq"))<5 or Len(Request.Form("qq"))>9 or Not IsNumeric(Request.Form("qq")) then
	Response.Write "<script language=javascript>alert('QQ�������\n\n�� QQ����ֻ�������֡�\n\n�� QQ���벻�ܵ���4λ�� \n\n�� QQ���벻�ܸ���9λ��\n\n�� �����û��QQ������Բ��');javascript:history.back();</script>"
	Response.End
	End If
	End If
        If Request.Form("uc")<>"" Then
	If Len(Request.Form("uc"))<5 or Len(Request.Form("uc"))>9 or Not IsNumeric(Request.Form("uc")) then
	Response.Write "<script language=javascript>alert('UC�������\n\n�� UC����ֻ�������֡�\n\n�� UC���벻�ܵ���4λ�� \n\n�� UC���벻�ܸ���9λ��\n\n�� �����û��UC������Բ��');javascript:history.back();</script>"
	Response.End
	End If
	End If
	If Request.Form("words")="" Then
	Response.Write "<script language=javascript>alert('�������ݲ���Ϊ�գ�');javascript:history.back();</script>"
	Response.End
	End If
        If Request.Form("words")<>"" Then
	If Len(Request.Form("words"))<4 or Len(Request.Form("words"))>255 then
	Response.Write "<script language=javascript>alert('�������ݴ�����ʾ��\n\n�� ���Բ��õ���4���֣�\n\n�� ���Բ��ø���255���֣���ƪ�����뵽��̳����');javascript:history.back();</script>"
	Response.End
	End If
	End If
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From words"
	Rs.Open Sql,Conn,2,3
	Rs.AddNew
        if Request.Form("name")=webceo then
        Rs("name")=noceo
        elseif Request.Form("name")=ceopass then
        Rs("name")=webceo
        else
        Rs("name")=Server.HTMLEncode(Request.Form("name"))
        End If
	Rs("sex")=Server.HTMLEncode(Request.Form("sex"))
        Rs("qq")=Server.HTMLEncode(Request.Form("qq"))
        Rs("uc")=Server.HTMLEncode(Request.Form("uc"))
	Rs("city")=Server.HTMLEncode(Request.Form("city"))
	Rs("web")=Server.HTMLEncode(Request.Form("web"))
	Rs("email")=Server.HTMLEncode(Request.Form("email"))
        Rs("admin")=Server.HTMLEncode(Request.Form("admin"))
        Rs("title")=Server.HTMLEncode(Request.Form("title"))
	Rs("words")=Server.HTMLEncode(Request.Form("words"))
	Rs("date")=Now()
        Rs("ip")=request.servervariables("remote_addr")
	Rs.Update
	Rs.Close
        Session("GetCode")=""
	Set Rs = Nothing

End Sub
Sub Delete()
	Conn.Execute("Delete * From words Where id="&Request.QueryString("id"))
End Sub
Sub Edit_Execute()
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From words Where id="&Request.Form("id")
	Rs.Open Sql,Conn,2,3
	Rs("words") = Server.HTMLEncode(Request.Form("words"))
	Rs("reply") = Server.HTMLEncode(Request.Form("reply"))
        Rs("admin") = Server.HTMLEncode(Request.Form("admin"))
        Rs("redate") = Now()
	Rs.Update
	Rs.Close
	Set Rs=Nothing
End Sub
Conn.Close
Set Conn = Nothing
function unHtml(content)
unHtml=content
if content <> "" then
unHtml=replace(unHtml,"<","&lt;")
unHtml=replace(unHtml,">","&gt;")
unHtml=replace(unHtml,chr(34),"&quot;")
unHtml=replace(unHtml,chr(13),"<br>")
unHtml=replace(unHtml,chr(32),"&nbsp;")
end if
end function
function ubb(content)
ubb=content
    nowtime=now()
    UBB=Convert(ubb,"code")
    UBB=Convert(ubb,"html")
    UBB=Convert(ubb,"url")
    UBB=Convert(ubb,"color")
    UBB=Convert(ubb,"font")
    UBB=Convert(ubb,"size")
    UBB=Convert(ubb,"quote")
    UBB=Convert(ubb,"email")
    UBB=Convert(ubb,"img")
    UBB=Convert(ubb,"swf")

    UBB=AutoURL(ubb)
    ubb=replace(ubb,"[b]","<b>",1,-1,1)
    ubb=replace(ubb,"[/b]","</b>",1,-1,1)
    ubb=replace(ubb,"[i]","<i>",1,-1,1)
    ubb=replace(ubb,"[/i]","</i>",1,-1,1)
    ubb=replace(ubb,"[u]","<u>",1,-1,1)
    ubb=replace(ubb,"[/u]","</u>",1,-1,1)
    ubb=replace(ubb,"[blue]","<font color='#000099'>",1,-1,1)
    ubb=replace(ubb,"[/blue]","</font>",1,-1,1)
    ubb=replace(ubb,"[red]","<font color='#990000'>",1,-1,1)
    ubb=replace(ubb,"[/red]","</font>",1,-1,1)
    for i=1 to 28
    ubb=replace(ubb,"{:em"&i&"}","<IMG SRC=emot/emotface/em"&i&".gif></img>",1,6,1)
    ubb=replace(ubb,"{:em"&i&"}","",1,-1,1)
    next
    ubb=replace(ubb,"["&chr(176),"[",1,-1,1)
    ubb=replace(ubb,chr(176)&"]","]",1,-1,1)
    ubb=replace(ubb,"/"&chr(176),"/",1,-1,1)
end function
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function
function Convert(ubb,CovT)
cText=ubb
startubb=1
do while Covt="url" or Covt="color" or Covt="font" or Covt="size"
startubb=instr(startubb,cText,"["&CovT&"=",1)
if startubb=0 then exit do
endubb=instr(startubb,cText,"]",1)
if endubb=0 then exit do
Lcovt=Covt
startubb=startubb+len(lCovT)+2
text=mid(cText,startubb,endubb-startubb)
codetext=replace(text,"[","["&chr(176),1,-1,1)
codetext=replace(codetext,"]",chr(176)&"]",1,-1,1)
codetext=replace(codetext,"/","/"&chr(176),1,-1,1)
select case CovT
    case "color"
	cText=replace(cText,"[color="&text&"]","<font color='"&text&"'>",1,1,1)
	cText=replace(cText,"[/color]","</font>",1,1,1)
    case "font"
	cText=replace(cText,"[font="&text&"]","<font face='"&text&"'>",1,1,1)
	cText=replace(cText,"[/font]","</font>",1,1,1)
    case "size"
	if IsNumeric(text) then
	if text>6 then text=6
	if text<1 then text=1
	cText=replace(cText,"[size="&text&"]","<font size='"&text&"'>",1,1,1)
	cText=replace(cText,"[/size]","</font>",1,1,1)
	end if
    case "url"
	cText=replace(cText,"[url="&text&"]","<a href='"&codetext&"' target=_blank>",1,1,1)
	cText=replace(cText,"[/url]","</a>",1,1,1)
    case "email"
	cText=replace(cText,"["&CovT&"="&text&"]","<a href=mailto:"&text&">",1,1,1)
	cText=replace(cText,"[/"&CovT&"]","</a>",1,1,1)
end select
loop
startubb=1
do
startubb=instr(startubb,cText,"["&CovT&"]",1)
if startubb=0 then exit do
endubb=instr(startubb,cText,"[/"&CovT&"]",1)
if endubb=0 then exit do
Lcovt=Covt
startubb=startubb+len(lCovT)+2
text=mid(cText,startubb,endubb-startubb)
codetext=replace(text,"[","["&chr(176),1,-1,1)
codetext=replace(codetext,"]",chr(176)&"]",1,-1,1)
codetext=replace(codetext,"/","/"&chr(176),1,-1,1)
select case CovT
    case "url"
	cText=replace(cText,"["&CovT&"]"&text,"<a href='"&codetext&"' target=_blank>"&codetext,1,1,1)
	cText=replace(cText,"<a href='"&codetext&"' target=_blank>"&codetext&"[/"&CovT&"]","<a href="&codetext&" target=_blank>"&codetext&"</a>",1,1,1)
    case "email"
	cText=replace(cText,"["&CovT&"]","<a href=mailto:"&text&">",1,1,1)
	cText=replace(cText,"[/"&CovT&"]","</a>",1,1,1)
    case "html"
	codetext=replace(codetext,"<br>",chr(13),1,-1,1)
	codetext=replace(codetext,"&nbsp;",chr(32),1,-1,1)
	Randomize
	rid="temp"&Int(100000 * Rnd)
	cText=replace(cText,"[html]"&text,"����Ƭ�����£�<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext,1,1,1)
	cText=replace(cText,"����Ƭ�����£�<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext&"[/html]","����Ƭ�����£�<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext&"</TEXTAREA><INPUT onclick=runEx('"&rid&"') type=button value=���д˶δ��� name=Button1 class='Tips_bo'> <INPUT onclick=JM_cc('"&rid&"') type=button value=���Ƶ��ҵļ����� name=Button2 class='Tips_bo'>",1,1,1)
    case "img"
	cText=replace(cText,"[img]"&text,"<a href="&chr(34)&"about:<img src="&codetext&" border=0>"&chr(34)&" target=_blank><img src="&codetext,1,1,1)
	cText=replace(cText,"[/img]"," vspace=2 hspace=2 border=0 alt=::���ͼƬ���´����д�::></a>",1,1,1)
    case "code"
	cText=replace(cText,"[code]"&text,"��������Ϊ�������<hr noshade>"&codetext,1,1,1)
	cText=replace(cText,"��������Ϊ�������<hr noshade>"&codetext&"[/code]","��������Ϊ�������<hr noshade>"&codetext&"<hr noshade>",1,1,1)
    case "quote"
	atext=replace(text,"[img]","",1,-1,1)
	atext=replace(atext,"[/img]","",1,-1,1)
	atext=replace(atext,"[swf]","",1,-1,1)
	atext=replace(atext,"[/swf]","",1,-1,1)
	atext=replace(atext,"[html]","",1,-1,1)
	atext=replace(atext,"[/html]","",1,-1,1)
	atext=SplitWords(atext,350)
	atext=replace(atext,chr(32),"&nbsp;",1,-1,1)
	cText=replace(cText,"[quote]"&text,"<blockquote><hr noshade>"&atext,1,1,1)
	cText=replace(cText,"<blockquote><hr noshade>"&atext&"[/quote]","<blockquote><hr noshade>"&atext&"<hr noshade></blockquote>",1,1,1)
    case "swf"
	cText=replace(cText,"[swf]"&text,"ӰƬ��ַ:<br>"&text&"<br><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='500'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='500'>",1,1,1)
	cText=replace(cText,"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='500'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='500'>"&"[/swf]","<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='500'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='500'>"&"</embed></object>",1,1,1)
end select
loop
Convert=cText
end function
function AutoURL(ubb)
cText=ubb
startubb=1
do
startubb=1
endubb_a=0
endubb_b=0
endubb=0
startubb=instr(startubb,cText,"http://",1)
if startubb=0 then exit do
endubb_b=instr(startubb,cText,"<",1)
endubb_a=instr(startubb,cText,"&nbsp;",1)
endubb=endubb_a
if endubb=0 then
endubb=endubb_b
end if
if endubb_b<endubb and endubb_b>0 then
endubb=endubb_b
end if
if endubb=0 then
lenc=ctext
endubb=len(lenc)+1
end if
if startubb>endubb then exit do
text=mid(cText,startubb,endubb-startubb)
codetext=text
urllink="<a href='"&codetext&"' target=_blank>"&codetext&"</a> "
urllink=replace(urllink,"/","/"&chr(176),1,-1,1)
cText=replace(cText,text,urllink,1,1,1)
loop
AutoURL=cText
end function
%>
</body>
</html>