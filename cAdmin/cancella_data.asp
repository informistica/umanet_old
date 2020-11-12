<!-- calcola_risultato_MODBC3.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	<style>
	<!--
	 li.MsoNormal
		{mso-style-parent:"";
		margin-bottom:.0001pt;
		font-size:12.0pt;
		font-family:"Times New Roman";
		margin-left:0cm; margin-right:0cm; margin-top:0cm}
	-->
	</style>
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
   <% Response.Buffer=True 
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
 
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                         
   idperiodo=Request.QueryString("idperiodo")
   divid=Request.QueryString("divid")
   id_classe=Request.QueryString("id_classe")
   

		      QuerySQL = "delete  from [dbo].[3PERIODI] where ID_Periodo=" &cint(idperiodo) &";"  
              ConnessioneDB.Execute QuerySQL
			 ' response.Redirect "../home.asp"
		 





On Error Resume Next
If Err.Number = 0 Then
	if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
	
Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>   
	 
	
	
	</body>
	</html>
	