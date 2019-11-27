<% Layout = Application(ScriptName & "Layout") %>
<table border="0" cellpadding="0" cellspacing="0"  width="206" height="100">
	<tr>
		<td height="200"  background="images/back_center_novo.gif" width="206"  colspan="3" align="center" valign="middle">
		<script language="JavaScript">
 //By Paul Davis - www.kaosweaver.com 
   var j,d="",l="",m="",p="",q="",z="",list= new Array()

   
    list[list.length]='http://www.madeiraonline.com.br/banners/portotech.gif';
  
    //list[list.length]='http://www.madeiraonline.com.br/banners/logo_simpi.jpg';
  
    //list[list.length]='http://www.madeiraonline.com.br/banners/mariomoraes.jpg';
    
    //list[list.length]='http://www.madeiraonline.com.br/banners/nota_informacao.gif';


   j=parseInt(Math.random()*list.length);
  j=(isNaN(j))?0:j;
    document.write("<img name='seqSlideShow' src='"+list[j]+"' border=0 >");
function seqSlideShow(t,l) {
  x=document.seqSlideShow;
  j=l;
  j++;
  if (j==list.length) j=0;
  x.src=list[j];
  setTimeout("seqSlideShow("+t+","+j+")",t);
  }
 </script>		
		<script language="JavaScript"> seqSlideShow(3000,0); </script></td>
	</tr>
    <tr>  
      <td height="62" background="images/back_center_novo.gif" width="206"  colspan="3" align="center" valign="middle"><img src="banners/paintball.gif" border="1"></td>
	  </tr>
		<tr>
      <td height="62" background="images/back_center_novo.gif" width="206"  colspan="3" align="center" valign="middle"><img src="banners/rondoniadinamica.gif" border="1"></td>
    <tr>
    </tr>
      <td height="62" background="images/back_center_novo.gif" width="206"  colspan="3" align="center" valign="middle"><img src="banners/vejahoje.gif" border="1"></td>
    </tr>
    <tr>  
      <td height="62" background="images/back_center_novo.gif" width="206"  colspan="3" align="center" valign="middle"><img src="banners/imagemnews.gif" border="1"></td>
	  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0"  width="206" height="100">
	<tr>
		<td height="206"  background="images/back_center_novo.gif" width="206" align="center" valign="middle"><!--#include file="destaques7.asp"--></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0"  width="206" height="100">
	<tr>
		<td height="206"  background="images/back_center_novo.gif" width="206" align="center" valign="middle"><!--#include file="pub_center7.asp"--></td>
	</tr>
</table>