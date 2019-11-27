<!--#include file="connDUnews.asp" -->
<%
set rsCount = Server.CreateObject("ADODB.Recordset")
rsCount.ActiveConnection = MM_connDUnews_STRING
rsCount.Source = "SELECT (SELECT COUNT(*) FROM News WHERE NEWS_APPROVED=1) AS News_COUNT, COUNT(*)   AS Category_COUNT  FROM Category"
rsCount.CursorType = 0
rsCount.CursorLocation = 2
rsCount.LockType = 3
rsCount.Open()
rsCount_numRows = 0
%>
<body>
<!--#include file="pub_right_novo.asp" -->
<table border="0" cellpadding="0" cellspacing="0" width="122" height="61">
	<tr>
		<td width="122" height="22" background="images/text6.gif" valign="top" style="padding-top:5;padding-left:30; border-right-width:1; border-top-width:1; border-bottom-width:1"><font class=bold>5 mais lidas</font></td>
	</tr>
	<tr>
    	<td width="122" height="39" align="center" valign="top"><!--#include file="manchetes_right.asp" --></td>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="122" height="50">
	<tr>
		<td width="122" height="22" background="images/text6.gif" valign="top" style="padding-top:5;padding-left:30; border-right-width:1; border-top-width:1; border-bottom-width:1"><font class=bold>Categorias</font></td>
	</tr>
    </tr>
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>
	<tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="veiculo.asp" --></td>
    </tr>
    </tr>
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>      
    <tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="municipio.asp" --></td>
    </tr>
    <tr>
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>  
    <tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="editorial.asp" --></td>
    </tr>        
    <tr>
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>    
    <tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="esporte.asp" --></td>
    </tr>
    <tr>    
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>      
    <tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="geral.asp" --></td>
    </tr>        
    <tr>    
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>
        <tr>
    	<td width="122" height="10" align="center" valign="top"><!--#include file="economia.asp" --></td>
    </tr>        
    <tr>    
    	<td width="122" height="5"><span style="font-size: 1pt">&nbsp;</span></td>
    </tr>
</table>