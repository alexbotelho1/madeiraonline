<meta http-equiv="refresh" content="180"><html><head><!--#include file="titles.asp"-->
<script type="text/JavaScript">
function popup(pagina,w,h,scroll)
{
	var t = (screen.height/2)-(h/2);
	var l = (screen.width/2)-(w/2);

    window.open(pagina,'popup','top=' +t+ ',left=' +l+ ',width=' +w+ ',height=' +h+ ',scrollbars=' + scroll );
}
</script>
</head>
<!--#include file="styles_novo.asp"-->
<body leftmargin=0 topmargin=0 marginwidth=0 marginheigh=0 style="word-spacing: 0; text-align: center; margin: 5">
<table height="100%" cellspacing="0" cellpadding="0" border="0">
  <tr>
		<td height="100%" valign="top">
			<table border="0" background="0" cellpadding="0" cellspacing="0" width="980" height="100%">
  				<tr><!--#include file="cabecalho_logo_novo.asp"--></tr>
  				<tr><!--#include file="cabecalho_menu_novo.asp"--></tr>				
  				<tr>
    				<td valign="top" height="100%" width="122" class="fundoesquerdo"><!--#include file="menu_novo.asp"--></td>	
    				<td valign="top" height="100%" width="736" class="fundocentro">
              <table border="0" cellpadding="0" cellspacing="0"  width="736" height="22">
                <tr>
                  <td height="22" background="images/center_line.gif" align="center" style="border-left:0px solid #000000; font-style:normal; font-variant:normal; font-weight:900; border-right-width:0; border-top-width:1; border-bottom-width:1" width="150">&nbsp;</td>
                  <td height="22" background="images/center_center.gif" align="center" style=font:900 width="436">      
					<script language="JavaScript">
<!--
var months=new Array(13);
months[1]="Janeiro";
months[2]="Fevereiro";
months[3]="Mar&ccedil;o";
months[4]="Abril";
months[5]="Maio";
months[6]="Junho";
months[7]="Julho";
months[8]="Agosto";
months[9]="Setembro";
months[10]="Outubro";
months[11]="Novembro";
months[12]="Dezembro";
var time=new Date();
var lmonth=months[time.getMonth() + 1];
var date=time.getDate();
year=time.getFullYear();
var today = new Date();
var hrs = today.getHours();
document.write("<left>");
document.write("<font face=Verdana size=1 color=#000000>");
if (hrs < 6)
document.write("Bom Dia ! -");
else if (hrs < 12)
document.write("Bom Dia ! -");
else if (hrs < 18)
document.write("Boa Tarde ! -");
else
document.write("Boa Noite ! -");   
document.write(" Porto Velho, ");
document.write(date + "  de ");
document.write(lmonth + " de  " + year + "</left>");
//-->
					</script>		
                  </td>
                  <td height="22" background="images/center_line.gif" align="center" style="border-right:0px solid #000000; font-style:normal; font-variant:normal; font-weight:900; border-left-width:1; border-top-width:1; border-bottom-width:1" width="150">&nbsp;</td>
                </tr>
              </table>
    				
    				
    				
    				
    				
    				
              <table border="0" background="0" cellpadding="0" cellspacing="0" width="530" height="100%">
                <tr>
                  <td valign="top" height="100%" width="530" class="fundocentro"><!--#include file="corpo_index_novo.asp"--></td>
                  <td valign="top" height="100%" width="206" class="fundocentro"><!--#include file="corpo_index_novo2.asp"-->
                    <table border="0" background="0" cellpadding="0" cellspacing="0" width="206" height="100%">
                      <tr>
                        <td valign="top" height="100%" width="206" class="fundocentro"><!--#include file="horoscopo_novo.asp"--></td>
                      </tr>	
                    </table>
                  </td>
                </tr>
              </table>
            </td>
            <td valign="top" height="100%" width="122" class="fundodireito"><!--#include file="facilidades_novo.asp"--></td>          
          </tr>
          <tr>             
            <td height="22" colspan="3"><!--#include file="rodape_princ_novo.asp"--></td>
          </tr>
			</table>
		</td>
	</tr>
</table>
</body></html>