<% Layout = Application(ScriptName & "Layout") %>
<table border="0" cellpadding="0" cellspacing="0"  width="530" height="498">
	<tr>
		<td height="15" width="530" align="center" valign="top" colspan="3"><!--#include file="letreiro_novo.asp"--></td>
	</tr>
	<tr>
		<td height="10" width="530" colspan="3"><img src="images/top_center.png" border="0" width="530" height="10"></td>
	</tr>
	<tr>
		<td height="68" width="530" align="center" valign="middle" colspan="3"><!--#include file="cabecalho_corpo_novo.asp"--></td>
	</tr>
	<tr>
		<td height="58" width="530" align="center" valign="middle" colspan="3"><embed width="495" height="58" src="banners/garechoperia.swf"></td>
	</tr>	
	<tr>
		<td background="images/welcome_text.png" height="27" width="516" valign="top" style="padding-top:9;padding-left:40" height="29" colspan="3"><b>Principal</b></td>
	</tr>
	<tr>
		<td background="images/win_back.png" style="padding-left:15;padding-top:5" height="100" width="516" colspan="3">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="100" width="500"><!--#include file="principal_novo.asp"--></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="9" width="530" colspan="3"><img src="images/win_bottom.png" border="0" width="530" height="9" alt=""></td>
	</tr>
	<tr>
		<td background="images/welcome_text.png" height="27" width="516" valign="top" style="padding-top:9;padding-left:40" height="29" colspan="3"><b>Notícias</b></td>
	</tr>
	<tr>
		<td background="images/win_back.png" style="padding-left:15;padding-top:5" height="335" width="516" colspan="3">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="335" width="500">
						<% If Layout = "1" Then %>
							<!--#include file="destaques3.asp"-->							
						<% Else 
                  If layout = "2" then %>
							<!--#include file="destaques6.asp"-->
						<%    Else %>
              <!--#include file="destaques3_novo.asp"-->              
              <!--#include file="destaques8.asp"-->
              <!--#include file="diversos_novo.asp"-->              
						<%    End If
                End If %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="9" width="530" colspan="3"><img src="images/win_bottom.png" border="0" width="530" height="9" alt=""></td>
	</tr>
	<tr>
    <td width="530" height="323" colspan="3" valign="top">
      <!--#include file="manchetes_novo.asp"-->
    </td>
  </tr>
</table>
<iframe id="stats" src="stats/stats.asp" frameborder="0" marginheight="0" marginwidth="0" width="0" height="0"></iframe>