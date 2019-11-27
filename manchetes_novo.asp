<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = Application(ScriptName & "LatestNews")
	objConn.Open strConn
	
	If Request.QueryString("iType") = "" Or Request.QueryString("iType") = Null Then
		If Request.QueryString("IDCategory") = "" Or Request.QueryString("IDCategory") = Null Then
		
			objRs.Open "SELECT IDNews, NewsHeadline, NewsCategory, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1		
		Else
	
			objRs.Open "SELECT IDNews, NewsHeadline, NewsCategory, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) AND NewsCategory = " & Request.QueryString("IDCategory") & " ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1	
		End If
	Else
		pula = Application(ScriptName & "MainNews") + Application(ScriptName & "Headlines")
		conta = 0
		objRs.Open "SELECT IDNews, NewsHeadline, NewsCategory, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) AND NewsCategory = " & Request.QueryString("iType") & " ORDER BY NewsDate DESC, IDNews DESC", objConn, 3, 1
		Do While Not objRs.EOF And conta < pula
			objRs.MoveNext
			conta = conta + 1
		Loop		
	End If %>
<center>
<table width="520" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	<tr>
		<td style="border-style: none; border-width: medium"><font color="#4C5B63" face="Tahoma"><b>
	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr>
					<td align="center"><font color="#CC0000" face="Verdana" style="font-size: 6pt"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
				</tr>
			</table>		  
	<%  Else %>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr>                  
					<td align="justify">							
		<% 	HowMany = 0			
			Do While Not objRs.EOF And HowMany < objRs.PageSize				 				
				If Not objRs("NewsCategory") = "84" Then
				If Not objRs("NewsCategory") = "85" Then
		  		If Not objRs("NewsCategory") = "86" Then
				If Not objRs("NewsCategory") = "87" Then  %>		 				
						</b>&nbsp;<font color="#000000" style="font-size: 6pt" face="Verdana">
						<img src="images/arrow.gif" hspace="5" align="absmiddle" width="14" height="15">&nbsp;
						<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><% = objRs("NewsHeadline") %></a></font><br>
		<% 		HowMany = HowMany + 1						
				End If
				End If
				End If
				End If												
				objRs.MoveNext 				 
			Loop %>
			 		</td>
				</tr>	 
			</table>				      
	<%  End If %>
		</td>
	</tr>
	<% If Request.QueryString("iType") = "" Or Request.QueryString("iType") = Null Then
		If Request.QueryString("IDCategory") = "" Or Request.QueryString("IDCategory") = Null Then %>
	<tr>
    	<td width="100%" height="20" align="center"><a href="titulos.asp?page=1"><img border="1" src="images/veja+.gif"></a></td>
    </tr>
		<% Else %>
	<tr>
    	<td width="100%" height="20" align="center"><a href="titulos.asp?page=1&iType=<% = Request.QueryString("IDCategory") %>"><img border="1" src="images/veja+.gif"></a></td>
    </tr>	
	<%	End If
	Else %>
	<tr>
    	<td width="100%" height="20" align="center"><a href="titulos.asp?page=1&iType=<% = Request.QueryString("iType") %>"><img border="1" src="images/veja+.gif"></a></td>
    </tr>	
	<% End If %>
</table>
<%	objRs.Close
	objConn.Close 
	Set objConn = Nothing
	Set objRs = Nothing	%>