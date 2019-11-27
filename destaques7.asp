<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = Application(ScriptName & "MainNews")
	objConn.Open strConn

	objRs.Open "SELECT IDNews, NewsPostar, NewsImage, NewsText, NewsSemFoto, NewsCodigoAlbum, NewsHeadline, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1 
	objRs.MoveFirst
	
    pularnoticia = 0
    
		Do While Not objRs.EOF And pularnoticia < 1
      If objRs("NewsPostar") = "1" and objRs("NewsSemFoto") = 0 Then
        pularnoticia = pularnoticia + 1
      End If
      objRs.MoveNext
		Loop
      
	 %>
<table border="0" cellpadding="0" cellspacing="0" width="200" height="1">
   <tr>
   		<td width="100%" height="100%" align="center"><span style="font-size: 1pt"><font style="font-size: 4pt">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -</font></span></td>
   </tr>
</table> 	
<%			noticiapequena = 0
			Do While Not objRs.EOF And noticiapequena < 8
				If objRs("NewsPostar") = "1" and objRs("NewsSemFoto") = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="200" height="70">
    <tr>
      <td width="200" height="100%">		
        <table width="100%" cellpadding="0" cellspacing="0" height="100%" style="border-collapse: collapse; border: 0px solid #000000">
         <tr>
          <td width="100%" height="65">
            <table border="0" cellpadding="2" cellspacing="2" align="left" width="100%" height="65">
              <tr>
                <% If objRs("NewsSemFoto") = 0 Then %>
                  <td width="80" height="60" align="center" style="border-style: solid; border-width: 0" bordercolor="#000066"><img src="<% = objRs("NewsImage") %>" class="News" align="left" width="80" height="65" hspace="0"></td>
                <% End If %>
                  <td width="100%" height="65" valign="middle"><p align="left" style="margin-left: 5; margin-right: 5"><a class="NewsLink" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font face="Verdana" style="color:000000" size="1" style="text-decoration: none"><% = Smilies(objRs("NewsHeadline"))%></font></a></td>
              </tr>
             </table>
          </td>
         </tr>
       </table>
		</td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="200" height="1">
   <tr>
   		<td width="100%" height="100%" align="center"><span style="font-size: 1pt"><font style="font-size: 4pt">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -</font></span></td>
   </tr>
</table> 	 						
<%	  noticiapequena = noticiapequena + 1
      End If 
			objRs.MoveNext
			Loop %>
<table border="0" cellpadding="0" cellspacing="0" width="200" height="5">
   <tr>
   		<td width="100%" height="100%" align="center"><span style="font-size: 1pt"><font style="font-size: 4pt">&nbsp;</font></span></td>
   </tr>
</table>
<%  noticialateral = 0
			Do While Not objRs.EOF And noticialateral < 1
				If objRs("NewsPostar") = "3" Then %>      
		<table border="0" cellpadding="0" cellspacing="0" width="100%" height="206" style="border-collapse: collapse; border: 0px solid #000000">
			<tr>
				<td width="206" height="155" align="center" style="border-style: solid; border-width: 0; margin-top:0; margin-bottom:0" bordercolor="#000066" valign="top"><p style="margin-top: 0; margin-bottom: 0"><img src="<% = objRs("NewsImage") %>" align="center" width="194" height="146" hspace="0"></td>
		    </tr>
		    <tr>
			  	<td width="100%" height="35" align="left" valign="top"><strong><img src="images/arrow.gif" align="absmiddle" width="14" height="15"><a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font color="#186997" face="Tahoma" size="2"><% = Smilies(objRs("NewsHeadline"))%></font></a></strong></td>
		    </tr>
		    <tr>
		    	<td valign="top" align="justify" width="100%" height="60"><p align="left" style="margin-left: 5; margin-right: 5"><a style="text-decoration: none; color: #000000; font-size: 7pt" href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>">&nbsp;&nbsp;
           <% NewsText = Smilies(objRs("NewsText"))
					If Len(NewsText) > Application(ScriptName & "MaxCharacters") Then %>
						<% = CutNewsText(NewsText) %>
            <% ThereIsMore = Tru  
				    Else %>
					</a></font>
                    	<% = NewsText %>
                        <% ThereIsMore = False											   
					End If %>             												
          		</td>
		    </tr>
		</table>
<table border="0" cellpadding="0" cellspacing="0" width="200" height="100%">
   <tr>
   		<td width="100%" height="100%" align="center"><span style="font-size: 1pt"><font style="font-size: 4pt">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -</font></span></td>
   </tr>
</table>
        	<% noticialateral = noticialateral + 1				
				End If				
			objRs.MoveNext					
			Loop 	
  objRs.Close
	objConn.Close 
	Set objRs = Nothing
	Set objConn = Nothing %>
