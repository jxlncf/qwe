<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="inc/config.asp"-->
<!--#include file="mdb.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style></head>

<body>
<table>  <tr>
        <td><div align="center">
          <script type="text/javascript">

imgUrl1="<%=pic1%>";
imgtext1="<%=tit1%>"
imgLink1=escape("<%=pic1_lnk%>");
imgUrl2="<%=pic2%>";
imgtext2="<%=tit2%>"
imgLink2=escape("<%=pic2_lnk%>");
imgUrl3="<%=pic3%>";
imgtext3="<%=tit3%>"
imgLink3=escape("<%=pic3_lnk%>");
imgUrl4="<%=pic4%>";
imgtext4="<%=tit4%>"
imgLink4=escape("<%=pic4_lnk%>");
imgUrl5="<%=pic5%>";
imgtext5="<%=tit5%>"
imgLink5=escape("<%=pic5_lnk%>");

 var focus_width=218
 var focus_height=150
 var text_height=22
 var swf_height = focus_height+text_height
 
 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
 
 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="img/photo.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
 document.write('<embed src="pixviewer.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#F0F0F0" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>'); 
            </script>
        </div></td>
      </tr></table>
</body>
</html>
