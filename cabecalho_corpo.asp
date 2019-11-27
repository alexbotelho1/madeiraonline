<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = 1
	objConn.Open strConn
	
		objRs.Open "SELECT IDNews, NewsHeadline, NewsCategory FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1 %>
<div align="center">
  <center>
  <table cellpadding="0" cellspacing="0" width="490">
  	<tr>
    	<td width="490" height="20" colspan="7" valign="middle" bgcolor="#CC0000">
        <p align="center"><b><font color="#FFFFFF" face="Tahoma" size="2">Manchetes dos principais jornais de Rondônia</font></b></td>
    </tr>
    <tr>
      <td width="115" height="100%">
      <!-- INÍCIO DO ALTOMADEIRA -->
	<table width="115" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	  	<tr>
	    	<td style="border-style: none; border-width: medium"><font color="#4C5B63" face="Tahoma"><b>
	    	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			  		<tr>
			    		<td align="center"><div align="center">
                          <font color="#CC0000" size="1" face="Verdana"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
			  		</tr>
				</table>		  
		  	<% Else %>			
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			 		<tr>                  
			   			<td align="justify" valign="top">							
				 		<% 	HowMany = 0
				 			Do While Not objRs.EOF And HowMany < objRs.PageSize
				 											
								If objRs("NewsCategory") = 84 Then %>
						<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><img border="0" src="images/jornal1_menu.jpg" width="115" height="21">
						<font color="#000000" size="1" face="Verdana">&nbsp;<% = objRs("NewsHeadline") %></font></a><br>
						<center><b><font color="#FF0000" size="2" face="Verdana">- - - - - - - - - -</font></b></center>
						
						<% 	HowMany = HowMany + 1
								
								End If					
							
							objRs.MoveNext 				 
				 			Loop %>
			   			</td>
					</tr>	 
				</table>				      
			<%  End If %>
	    	</td>
		</tr>
	</table>
      </td>
      <td width="10" height="100%">&nbsp;</td>
      <td width="115" height="100%">
<!-- INÍCIO DO DIÁRIO DA AMAZÔNIA -->
	<table width="115" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	  	<tr>
	    	<td style="border-style: none; border-width: medium"><font color="#4C5B63" face="Tahoma"><b>
	    	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			  		<tr>
			    		<td align="center"><div align="center">
                          <font color="#CC0000" size="1" face="Verdana"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
			  		</tr>
				</table>		  
		  	<% Else %>			
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			 		<tr>                  
			   			<td align="justify" valign="top">							
				 		<% 	HowMany = 0
				 		    objRs.MoveFirst
				 			Do While Not objRs.EOF And HowMany < objRs.PageSize
				 											
								If objRs("NewsCategory") = 85 Then %>
						<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><img border="0" src="images/jornal2_menu.jpg" width="115" height="21">
						<font color="#000000" size="1" face="Verdana">&nbsp;<% = objRs("NewsHeadline") %></a></font><br>
						<center><b><font color="#FF0000" size="2" face="Verdana">- - - - - - - - - -</font></b></center>
						<% 	HowMany = HowMany + 1
								
								End If					
							
							objRs.MoveNext 				 
				 			Loop %>
			   			</td>
					</tr>	 
				</table>			      
			<%  End If %>
	    	</td>
		</tr>
	</table>
      </td>
      <td width="10" height="100%">&nbsp;</td>
      <td width="115" height="100%">
<!-- INÍCIO DO O ESTADÃO DO NORTE -->
	<table width="115" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	  	<tr>
	    	<td style="border-style: none; border-width: medium"><font color="#4C5B63" face="Tahoma"><b>
	    	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			  		<tr>
			    		<td align="center"><div align="center">
                          <font color="#CC0000" size="1" face="Verdana"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
			  		</tr>
				</table>		  
		  	<% Else %>			
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			 		<tr>                  
			   			<td align="justify" valign="top">							
				 		<% 	HowMany = 0
				 			objRs.MoveFirst
				 			Do While Not objRs.EOF And HowMany < objRs.PageSize
				 											
								If objRs("NewsCategory") = 86 Then %>
						<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><img border="0" src="images/jornal4_menu.jpg" width="115" height="21">
						<font color="#000000" size="1" face="Verdana">&nbsp;<% = objRs("NewsHeadline") %></a></font><br>
						<center><b><font color="#FF0000" size="2" face="Verdana">- - - - - - - - - -</font></b></center>
						<% 	HowMany = HowMany + 1
								
								End If					
							
							objRs.MoveNext 				 
				 			Loop %>
			   			</td>
					</tr>	 
				</table>			      
			<%  End If %>
	    	</td>
		</tr>
	</table>
      </td>
      <td width="10" height="100%">&nbsp;</td>
      <td width="115" height="100%"> 
<!-- INÍCIO DO FOLHA DE RONDÔNIA -->
	<table width="115" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	  	<tr>
	    	<td style="border-style: none; border-width: medium"><font color="#4C5B63" face="Tahoma"><b>
	    	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			  		<tr>
			    		<td align="center"><div align="center">
                          <font color="#CC0000" size="1" face="Verdana"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
			  		</tr>
				</table>		  
		  	<% Else %>			
				<table border="0" cellpadding="0" cellspacing="0" width="115" align="center">
			 		<tr>                  
			   			<td align="justify" valign="top">							
				 		<% 	HowMany = 0
				 			objRs.MoveFirst
				 			Do While Not objRs.EOF And HowMany < objRs.PageSize
				 											
								If objRs("NewsCategory") = 87 Then %>
						<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><img border="0" src="images/jornal3_menu.jpg" width="115" height="21">
						<font color="#000000" size="1" face="Verdana">&nbsp;<% = objRs("NewsHeadline") %></a></font><br>
						<center><b><font color="#FF0000" size="2" face="Verdana">- - - - - - - - - -</font></b></center>
						<% 	HowMany = HowMany + 1
								
								End If					
							
							objRs.MoveNext 				 
				 			Loop %>
			   			</td>
					</tr>	 
				</table>				      
			<%  End If %>
	    	</td>
		</tr>
	</table>
      </td>
   </tr>
</table></center>
</div>
<%	objRs.Close
	objConn.Close 
	Set objConn = Nothing
	Set objRs= Nothing %>