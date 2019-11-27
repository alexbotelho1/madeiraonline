<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = Application(ScriptName & "MainNews")
	objConn.Open strConn

	objRs.Open "SELECT IDNews, NewsPostar, NewsImage, NewsText, NewsSemFoto, NewsCodigoAlbum, NewsHeadline, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1 %>
<body topmargin="0" leftmargin="0"><center>
<table border="0" cellpadding="0" cellspacing="0" width="450" height="80">
<% 	quantidade = objRs.PageSize/2
	coluna = 0
	Do While Not objRs.EOF And coluna < quantidade %>
	<tr>
      	<td width="220" height="100%" valign="top">
<!-- LADO ESQUERDO -->
			<table border="0" cellspacing="0" cellpadding="0" width="220" height="100%">
				<tr>
           			<td width="220" height="100%" class="fundodestaquecorpo">
		<%  noticia = 0
			Do While Not objRs.EOF And noticia < 1
				If objRs("NewsPostar") = "2" and objRs("NewsSemFoto") = "1" Then %>
               			<table width="220" border="0" cellpadding="0" cellspacing="0" height="100%">
                  			<tr>
                    			<td width="220" height="25" valign="top" class="fundodestaquetitulo"><p style="margin-left: 3; margin-right: 3" align="left"><strong><img src="images/arrow.gif" align="middle" hspace="5"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></strong></td>
                 			</tr>
         					<tr>
                    			<td width="220" height="40">
                    				<table border="0" cellpadding="2" cellspacing="2" align="left" width="220" height="100%">
                      					<tr>
                        					<td width="220" height="100%" valign="top"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
                    <% NewsText = Smilies(objRs("NewsText"))
					If Len(NewsText) > Application(ScriptName & "MaxCharactersDestaques") Then %> <% = CutNewsText2(NewsText) %>
                    	<% ThereIsMore = True
				    Else %>
						   					</font>
                    	<% = NewsText %>
                        <% ThereIsMore = False
					End If  %>
                        					</td>
                 						</tr>
                    				</table>
                    			</td>
                  			</tr>
                  			<tr>
                    			<td width="220" height="15" align="right"><p style="margin-left: 5; margin-right: 5">                    										                      											
                    				<% if objRs("NewsCodigoAlbum") > 0 then %>
            						<img border="0" src="images/foto.gif" width="14" height="11">
            						<a class="NewsLink" href="album.asp?codalbum=<% = objRs("NewsCodigoAlbum") %>">        			
            						<font color="#186997" size="1" face="Verdana"><b>Veja Álbum</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            						<% end if %>
                    				<a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">
                               		<font color="#186997" size="1" face="Verdana"><b>Veja mais...</b></font></a>                                                                                										                  													
                    			</td>
                  			</tr>
                		</table>
           <%	   noticia = noticia + 1
				End If
			objRs.MoveNext
			Loop %>
		  			</td>
       			</tr>
 			</table>
		</td>
      	<td width="10" height="100%">&nbsp;</td>
      	<td width="220" height="100%" valign="top">
<!-- LADO DIREITO -->
      		<table border="0" cellspacing="0" cellpadding="0" width="220" height="100%">
       			<tr>
           			<td width="220" height="100%" class="fundodestaquecorpo">
		<%  noticia = 0
			Do While Not objRs.EOF And noticia < 1
				If objRs("NewsPostar") = "2" and objRs("NewsSemFoto") = "1" Then %>
                		<table width="220" border="0" cellpadding="0" cellspacing="0" height="100%">
                  			<tr>
                    			<td width="220" height="25" valign="top" class="fundodestaquetitulo"><p style="margin-left: 3; margin-right: 3" align="left"><strong><img src="images/arrow.gif" align="middle" hspace="5"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></strong></td>
                 			</tr>
                 			<tr>
                    			<td width="220" height="40">
                    				<table border="0" cellpadding="2" cellspacing="2" width="220" height="100%">
                        					<td width="220" height="100%" valign="top"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
                    <% NewsText = Smilies(objRs("NewsText"))                          				     			     
					If Len(NewsText) > Application(ScriptName & "MaxCharactersDestaques") Then %> <% = CutNewsText2(NewsText) %>
                    	<% ThereIsMore = True				     	
				    Else %>
						   					</font>
                    	<% = NewsText %>
                        <% ThereIsMore = False
					End If  %>
                     						</td>
       									</tr>
                    				</table>
                    			</td>
                  			</tr>
                  			<tr>
                    			<td width="220" height="15" align="right"><p style="margin-left: 5; margin-right: 5">                     											
                      				<% if objRs("NewsCodigoAlbum") > 0 then %>
            						<img border="0" src="images/foto.gif" width="14" height="11">
            						<a class="NewsLink" href="album.asp?codalbum=<% = objRs("NewsCodigoAlbum") %>">
            						<font color="#186997" size="1" face="Verdana"><b>Veja Álbum</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            						<% end if %>
                      				<a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">
                               		<font color="#186997" size="1" face="Verdana"><b>Veja mais...</b></font></a>
             					</td>
                  			</tr>
                		</table>              
           <%	   noticia = noticia + 1
				End If
			objRs.MoveNext	
			Loop %>
					</td>
       			</tr>
 			</table>
		</td>
		<tr>
      <td width="450" height="10" colspan="2"></td>
    </tr>
	</tr>
<%  coluna = coluna + 1
	Loop %>
</table>
</center>
<%	objRs.Close
	objConn.Close 
	Set objRs = Nothing
	Set objConn = Nothing %>