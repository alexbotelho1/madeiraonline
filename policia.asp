<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = Application(ScriptName & "MainNews")
	objConn.Open strConn
	
	objRs.Open "SELECT IDNews, NewsPostar, NewsCategory, NewsImage, NewsText, NewsHeadline, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1 %>
<center>
<table width="118" height="140" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%" height="100%">
		<% If objRs.PageCount = 0 Then %>
			<table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
            	<tr>
                   	<td width="100%" height="100%" valign="top"><i>Não existe notícia lançada!</i></td>
				</tr>
			</table>
		<% Else

 			noticia = 0
			Do While Not objRs.EOF AND noticia < 1
			
				If objRs("NewsCategory") = "79" Then
				If objRs("NewsPostar") = "0" Then %>				
			<table width="100%" height="110" border="0" cellpadding="0" cellspacing="0" class="fundolateralcorpo">
				<tr>
					<td width="100%" height="20" align="left"><H2>Polícia</H2></td>
				</tr>
				<tr>
                   	<% If Len(objRs("NewsImage")) > 0 And Application(ScriptName & "NewsImages") = 1 And Application(ScriptName & "NewsImagesAlign") = "left" Then %>
                    <td width="100%" height="70" align="center" valign="middle"><img src="<% = objRs("NewsImage") %>" align="center" width="80" height="65" hspace="0" border="2"></td>
                   	<% End If %>
           		</tr>
           		<tr>
					<td valign="top" width="100%" height="30"><H2><a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><% = Smilies(objRs("NewsHeadline"))%></a></H2></td>
 				</tr>
			</table>
				<%  noticia = noticia + 1
					idnoticia = objRs("IDNews")
				End If
				End If
			objRs.MoveNext
			Loop %>
			<table width="100%" height="10" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
			<% noticia = 0
			objRs.MoveFirst
			Do While Not objRs.EOF AND noticia < 2
				If objRs("NewsCategory") = "79" and idnoticia <> objRs("IDNews") Then %>
				<tr>
  					<td width="100%" height="10"><P style="margin-left: 5; margin-right: 0"><img src="images/seta-preta.gif" width="8" height="10"><font face="Verdana" size="1"><a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><% = Smilies(objRs("NewsHeadline"))%></a></b></td>
           		</tr>             		
					<% if noticia < 1 then %>
				<tr>
      				<td width="118" height="1" align="center"><font style="font-size: 8pt">- - - - - - - - - - - - -</font></td>
    			</tr>
					<% End If
					noticia = noticia + 1
				End If			
			objRs.MoveNext					
			Loop %>
			</table>		                
	<% End If %>
		</td>
	</tr>
</table></center>
<%	objRs.Close
	objConn.Close 
	Set objRs = Nothing
	Set objConn = Nothing %>