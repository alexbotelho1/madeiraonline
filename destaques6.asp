<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = Application(ScriptName & "MainNews")
	objConn.Open strConn

	objRs.Open "SELECT IDNews, NewsPostar, NewsImage, NewsText, NewsSemFoto, NewsCodigoAlbum, NewsHeadline, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1 %>
<table border="0" cellpadding="0" cellspacing="0" width="500" height="300">
    <tr>
      <td width="245" height="300" class="fundodestaquecorpo">
		<%  noticialateral = 0
			Do While Not objRs.EOF And noticialateral < 1
				If objRs("NewsPostar") = "3" Then %>      
		<table border="0" cellpadding="0" cellspacing="0" width="100%" height="300" style="border-collapse: collapse; border: 0px solid #000000">
			<tr>
				<td width="240" height="180" align="center" style="border-style: solid; border-width: 0; margin-top:0; margin-bottom:0" bordercolor="#000066" valign="top"><p style="margin-top: 0; margin-bottom: 0"><img src="<% = objRs("NewsImage") %>" align="center" width="240" height="180" hspace="0"></td>
		    </tr>
		    <tr>
			  	<td width="100%" height="35" align="left" class="fundodestaquetitulo" valign="top"><strong><img src="images/arrow.gif" align="absmiddle" width="14" height="15"><a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></strong></td>
		    </tr>
		    <tr>
		    	<td valign="top" align="justify" width="100%" height="60"><font face="Verdana" style="color:000000" size="2"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
                    <% NewsText = Smilies(objRs("NewsText"))
					If Len(NewsText) > Application(ScriptName & "MaxCharacters") Then %>
						<% = CutNewsText(NewsText) %>
                    	<% ThereIsMore = Tru  
				    Else %>
					</font>
                    	<% = NewsText %>
                        <% ThereIsMore = False											   
					End If %>             												
          		</td>
		    </tr>
		    <tr>
		    	<td valign="top" width="100%" height="15" align="right">
            		<% if objRs("NewsCodigoAlbum") > 0 then %>
            			<img border="0" src="images/foto.gif" width="14" height="11">
            			<a href="album.asp?codalbum=<% = objRs("NewsCodigoAlbum") %>">                     			
            			<font color="#186997" size="1" face="Verdana"><b>Veja Álbum</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            		<% end if %>
            			<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">
            			<font color="#186997" size="1" face="Verdana"><b>Veja mais...</b></font></a>                      												
            	</td>
		    </tr> 
		</table>	
        	<% noticialateral = noticialateral + 1				
				End If				
			objRs.MoveNext					
			Loop %>
  			<table border="0" cellpadding="0" cellspacing="0" width="100%" height="1">
    			<tr>
      				<td width="100%" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    			</tr>
  			</table> 					    
      </td>
      <td width="5" height="300">&nbsp;</td>
      <td width="250" height="310" valign="top">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" height="155">
				<tr>
           			<td width="100%" height="100%" class="fundodestaquecorpo">
		<%  objRs.MoveFirst
			noticiapequena = 0
			Do While Not objRs.EOF And noticiapequena < 1
				If objRs("NewsPostar") = "2" Then %>
               			<table width="100%" cellpadding="0" cellspacing="0" height="100%" style="border-collapse: collapse; border: 0px solid #000000">
                  			<tr>
                    			<td width="100%" height="30" valign="top" class="fundodestaquetitulo"><p style="margin-left: 3; margin-right: 3" align="left"><img src="images/arrow.gif" align="middle" hspace="5"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></td>
                 			</tr>
         					<tr>
                    			<td width="100%" height="65">
                    				<table border="0" cellpadding="2" cellspacing="2" align="left" width="100%" height="65">
                      					<tr>
                     				<% If objRs("NewsSemFoto") = 0 Then %>
                        					<td width="80" height="60" align="center" style="border-style: solid; border-width: 0" bordercolor="#000066"><img src="<% = objRs("NewsImage") %>" class="News" align="left" width="80" height="65" hspace="0"></td>
                    				<% End If %>
                        					<td width="100%" height="65" valign="top"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
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
                    			<td width="100%" height="15" align="right"><p style="margin-left: 5; margin-right: 5">                    										                      											
                    				<% if objRs("NewsCodigoAlbum") > 0 then %>
            						<img border="0" src="images/foto.gif" width="14" height="11">
            						<a href="album.asp?codalbum=<% = objRs("NewsCodigoAlbum") %>">        			
            						<font color="#186997" size="1" face="Verdana"><b>Veja Álbum</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            						<% end if %>
                    				<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">
                               		<font color="#186997" size="1" face="Verdana"><b>Veja mais...</b></font></a>                                                                                										                  													
                    			</td>
                  			</tr>
                		</table>
           <%	noticiapequena = noticiapequena + 1
           		End If
			objRs.MoveNext
			Loop %>
		  			</td>
       			</tr>
 			</table>
  			<table border="0" cellpadding="0" cellspacing="0" width="100%" height="8">
    			<tr>
      				<td width="100%" height="8"><span style="font-size: 1pt">&nbsp;</span></td>
    			</tr>
  			</table> 	 						
      		<table border="0" cellspacing="0" cellpadding="0" width="100%" height="155">
       			<tr>
           			<td width="100%" height="100%" class="fundodestaquecorpo">
		<%  noticiapequena = 0
			Do While Not objRs.EOF And noticiapequena < 1
				If objRs("NewsPostar") = "2" Then %>
               			<table width="100%" cellpadding="0" cellspacing="0" height="100%" style="border-collapse: collapse; border: 0px solid #000000">
                  			<tr>
                    			<td width="100%" height="30" valign="top" class="fundodestaquetitulo"><p style="margin-left: 3; margin-right: 3" align="left"><img src="images/arrow.gif" align="middle" hspace="5"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></td>
                 			</tr>
         					<tr>
                    			<td width="100%" height="65">
                    				<table border="0" cellpadding="2" cellspacing="2" align="left" width="100%" height="65">
                      					<tr>
                     				<% If objRs("NewsSemFoto") = 0 Then %>
                        					<td width="80" height="60" align="center" style="border-style: solid; border-width: 0" bordercolor="#000066"><img src="<% = objRs("NewsImage") %>" class="News" align="left" width="80" height="65" hspace="0"></td>
                    				<% End If %>
                        					<td width="100%" height="65" valign="top"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
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
                    			<td width="100%" height="15" align="right"><p style="margin-left: 5; margin-right: 5">                    										                      											
                    				<% if objRs("NewsCodigoAlbum") > 0 then %>
            						<img border="0" src="images/foto.gif" width="14" height="11">
            						<a href="album.asp?codalbum=<% = objRs("NewsCodigoAlbum") %>">        			
            						<font color="#186997" size="1" face="Verdana"><b>Veja Álbum</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            						<% end if %>
                    				<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">
                               		<font color="#186997" size="1" face="Verdana"><b>Veja mais...</b></font></a>                                                                                										                  													
                    			</td>
                  			</tr>
                		</table>                 		              
           <%	noticiapequena = noticiapequena + 1
           		End If
			objRs.MoveNext
			Loop %>
					</td>
       			</tr>
 			</table>      
      </td>
    </tr>
    <tr>
      <td width="500" height="180" colspan="3" class="fundodestaquecorpo">
		<%  objRs.MoveFirst
			noticiagrande = 0
			Do While Not objRs.EOF And noticiagrande < 1
				If objRs("NewsPostar") = "4" Then %>
			<table cellpadding="2" cellspacing="2" align="left" width="500" height="180" style="border-collapse: collapse; border: 0px solid #000000">
				<tr>
            		<td width="240" height="180" align="center" style="border-style: solid; border-width: 0; margin-top:0; margin-bottom:0" bordercolor="#000066" rowspan="3" valign="top"><p style="margin-top: 0; margin-bottom: 0"><img src="<% = objRs("NewsImage") %>" align="center" width="240" height="180" hspace="0"></td>
               		<td width="100%" height="30" align="left" valign="top" class="fundodestaquetitulo"><strong><img src="images/arrow.gif" align="absmiddle" width="14" height="15"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></strong></td>
               	</tr>
               	<tr>
          			<td valign="top" width="100%" height="130"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><p align="left" style="margin-left: 5; margin-right: 5">&nbsp;&nbsp;
                    	<% NewsText = Smilies(objRs("NewsText"))
						If Len(NewsText) > Application(ScriptName & "MaxCharacters") Then %>
							<% = CutNewsText(NewsText) %>
                    		<% ThereIsMore = True				     	
				    	Else %>
						</font>
                    		<% = NewsText %>
                        	<% ThereIsMore = False											   
						End If %>             												
          			</td>
          		</tr>
          		<tr>
            		<td valign="top" width="100%" height="15" align="right">
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
           <%	noticiagrande = noticiagrande + 1
           		End If
			objRs.MoveNext
			Loop %>
      </td>
    </tr>
</table>
<%	objRs.Close
	objConn.Close 
	Set objRs = Nothing
	Set objConn = Nothing %>