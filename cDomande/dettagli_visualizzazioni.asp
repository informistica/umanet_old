<!-- esegui_test_MODBC3.asp -->

<%@ Language=VBScript %>
<% Response.Buffer=True %>

<HTML>
<HEAD>
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

<TITLE>VISUALIZZAZIONI</TITLE>
</HEAD>
<BODY> 

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

 Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato

    'StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
  
 Capitolo=Request.QueryString("Stato") 
 Paragrafo=Request.QueryString("Cartella")
 ID_Mod=Request.QueryString("ID_Mod")  
 ID_Paragrafo= Request.QueryString("Id_Paragrafo") 
 cod=Request.QueryString("cod") 
 tipo=Request.QueryString("tipo") ' =0 mostro tutte le visualizzazioni del capitolo; =1 solo quelle del paragrafo 
  		'response.write("Stato"&stato)				
%>
 
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color="#FF0000">VISUALIZZAZIONI:</font></b> 
</p> <!-- stampa il titolo del test -->

  <table border="0" align=center width="60%">
		<tr>
			<td colspan=3 align=center>
			  <font color="#000000"><b><h3><%=Capitolo%></h3></font>
			</td>
		</tr>
		<tr>
			<td colspan=3 align=center>
			  <font color="#000000"><b><h4><%=Paragrafo%></h4></font>
			</td>
		</tr>
	</table>
 <%
 if tipo=0 then
		QuerySQL="SELECT * " &_
		" FROM Elenco_tutte_visualizzazioni " &_
		" WHERE (Visualizzazioni.CodiceAllievo='"&cod&"' AND Moduli.ID_Mod='"&ID_Mod&"') order by Moduli.ID_Mod;"
 else 
        QuerySQL="SELECT * " &_
		" FROM Elenco_tutte_visualizzazioni " &_
		" WHERE (Visualizzazioni.CodiceAllievo='"&cod&"' AND Moduli.ID_Mod='"&ID_Mod&"' AND Paragrafi.ID_Paragrafo='"&ID_Paragrafo&"') ;"
 end if 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
 
 <table border="1" align=center width="60%">
 <%
 
  Do until rsTabella.EOF%>
  		<tr>
			<td><%=rsTabella(0)%></font></td>
			<td><a href="spiegazione_test_img.asp?davisualizzazioni=1&cartella=<%=rsTabella.fields("cartella")%>&CodiceDomanda=<%=rsTabella.fields("CodiceDomanda")%>&CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>&Modulo=<%=rsTabella.fields("ID_Mod")%>"><%=rsTabella(1)%></a></font></td>
		</tr>  
 
    <%
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 %>
		
	</table>
	<br>
<%   
  
    

 

  
 
 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
   


  </div>
  </div>
</BODY>
</HTML>