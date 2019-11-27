<%	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.PageSize = 5
	objConn.Open strConn	
	If Request.QueryString("iType") = "" Or Request.QueryString("iType") = Null Then		
		objRs.Open "SELECT IDNews, NewsRead, NewsHeadline, NewsCategory, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) ORDER BY NewsDate DESC", objConn, 3, 1
	Else	
		objRs.Open "SELECT IDNews, NewsRead, NewsHeadline, NewsCategory, NewsDate FROM News WHERE releaseDate <= DATE() AND (expireDate >= DATE() OR expireDate IS NULL) AND NewsCategory = " & Request.QueryString("iType") & " ORDER BY NewsDate DESC", objConn, 3, 1
	End If %>
<body topmargin="0" style="text-align: center; text-indent: 0; line-height: 100%; line-width: 115; word-spacing: 0; margin: 0">
	<table width="118" height="65" border="0" cellspacing="0" cellpadding="0" align="center" style="border-width: 0">
	  	<tr>
	    	<td style="border-style: none; border-width: medium" height="65" width="118"><font color="#4C5B63" face="Tahoma"><b>
	    	<% If (objRs.BOF And objRs.EOF) Or objRs.PageCount = 0 Then %>
				<table border="0" cellpadding="0" cellspacing="0" width="118" height="30" align="center">
			  		<tr>
			    		<td align="center" width="118" height="30"><div align="center"><font color="#CC0000" face="Verdana" style="font-size: 6pt"><b><i>N&atilde;o h&aacute; nada no momento!</i></b></font></div></td>
			  		</tr>
				</table>		  
		  	<%  Else 
		  			If Len(Request.QueryString("page")) > 0 Then  			   
						objRs.AbsolutePage = Request.QueryString("page")					
					Else				
						objRs.AbsolutePage = 1					
					End If %>			
				<table cellpadding="0" cellspacing="0" height="35" width="118" class="fundocincomais">
			 		<tr>                  
			   			<td width="118" height="35" style="margin-left: 10; margin-right: 10">							
				 		<% 	HowMany = 0
				 			leituras = 20
				 			Do While Not objRs.EOF And HowMany < objRs.PageSize
				 				If objRs("NewsCategory") = 65 OR objRs("NewsCategory") = 66 OR objRs("NewsCategory") = 89 OR objRs("NewsCategory") = 90 Then
		
				 				Else
				 					If objRs("NewsRead") = leituras Then %>					
										<a href="ler_noticia.asp?IDNews=<% = objRs("IDNews") %>"><font class="textocincomais" style="font-size: 7pt" face="Verdana">&nbsp;<% = objRs("NewsHeadline") %></font></a><br><center><hr width="100"></center>
						<% 				HowMany = HowMany + 1
										leituras = leituras - 3
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
	</table>
</body>
<%	objRs.Close
	objConn.Close 
	Set objConn = Nothing
	Set objRs = Nothing	%>