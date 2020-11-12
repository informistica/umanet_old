<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
 <%Function domandaplus()
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 Cartella=rsTabella.fields("Cartella")
	 Modulo=rsTabella.fields("ID_Mod")
	 'Paragrafo=rsTabella(15)
	 Paragrafo=rsTabella.fields("Titolo")
	' response.write("PARAGRAFO="&Paragrafo)
	 Id=rsTabella.fields("CodiceDomanda")
	'homesito="/anno_2010-2011_ITC/ECDL"
	 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	'response.write(url)
	objTextFile.Close
End Function %>
<% Response.Buffer=True %>

<html>
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<%
  
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
   
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO,i
  Dim ConnessioneDB,rsTabella, QuerySQL,CodiceTest,StringaConnessione
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<% 
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
 Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  tipo=Request.QueryString("tipo")
  sint=Request.QueryString("sint") ' se è valorizzato =1 non mostro la risposta esatta, né la data, né l'autore serve
  sint=1 'lo metto a 1 per l'esportazione dei pdf  poi andrà tolto 
  criterio=""
  if strcomp(tipo,"Vero/Falso")=0 then
    criterio="and VF=1"
  end if
  
   if strcomp(tipo,"risposta chiusa singola")=0 then
    criterio="and VF=0 and Multiple=0"

  end if
  
   if strcomp(tipo,"risposta chiusa multipla")=0 then
    criterio="and Multiple=1"

  end if
  
    CodiceAllievo=Request.QueryString("CodiceAllievo")
if clng(Stato)=1 then	
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Mod='"&Modulo&"' and Segnalata=0 " & criterio
 else
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"&Codice_Test&"' and Segnalata=0 " & criterio
 end if
 
 ' mi faccio passare la query dalla pagina precedente di spiegazione che ha più filtri
 
  QuerySQL=Request.QueryString("QuerySQL")
 
 			

%>
   

<body bgcolor="#FFFFFF">
<div id="container">  

<%

 'response.write(QuerySQL)	
 if (InStr(QuerySQL,"drop")=0) and (InStr(QuerySQL,"delete")=0) and (QuerySql<>"") then
Set rsTabella = ConnessioneDB.Execute(QuerySQL)	

	
else

if clng(Stato)=1 then	
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Mod='"&Modulo&"' and Segnalata=0 " & criterio
 else
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"&Codice_Test&"' and Segnalata=0 " & criterio
 end if
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)	


end if
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
%>

<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%><center>
  <H4>Domande del Test non ancora disponibili!<h4></h4>
<% Else%>
<center>
  <h2><%=Capitolo%></h2> 
  <h3><%=Paragrafo%></h3> 
  <h4>Tipologia <%=tipo%></h4> 
 </center>
  <%
  i=1 'inizializza la variabile i (contatore delle domande totali) 
  k=1 'inizializza la variabile k (contatore delle domande per paragrafi)
  Do until rsTabella.EOF
  	  	if strcomp(titoloParagrafo,rsTabella(0))<>0 then
		 i=1
	       titoloParagrafo=rsTabella(0) 
			  if (i=1) then%>
				 <b><center>      <%=rsTabella(0)%>  </center></b>
				 <hr>
			   <% 
			  end if
	     end if
				if (rsTabella("Id_Sottoparagrafo")<>"") then
					if (StrComp(Sottoparagrafo, rsTabella("Id_Sottoparagrafo")) <> 0)  then
					   querySqlSotto="select Titolo,Id_Sottoparagrafo from Sottoparagrafi where Id_Sottoparagrafo='"&rsTabella("Id_Sottoparagrafo")&"'"
					 set rsTabellaSotto=ConnessioneDB.execute (querySqlSotto)
				
					   Sottoparagrafo=rsTabellaSotto("Id_Sottoparagrafo")						
						%>
						<b><center><i>  <%=rsTabellaSotto("Titolo")%>  </i></center></b> 
					 <%end if%>
				<%end if%>  
  <%  ID=rsTabella("CodiceDomanda")
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
   url=Replace(url,"\","/")     
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll
objTextFile.Close   ' la soluzione seguente la rimuovo e dirò di copiare ed incollare la domanda plus nella spiegazione
%>
  <div>
  <table>
		<tr>
			<td style="width:auto"><b> <%=k&") "&rsTabella("Quesito") %>  </b></td>
			<td> </td>
            <td> </td>             
		</tr>
		<% if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
	    <tr><td colspan="3"><p>
			 <%
			 Response.write(domandaplus())%> <br></td></tr><br>
        <%end if %> 
		<tr>
			<td colspan=3>			 
			<p >
			 <%
			 Response.write(sReadAll)%>  <br>
		     </td>		 
		</tr>
	</table>
    </div>
	<br>
<%    
       i = i+ 1 
	   k=k+1
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
</div>
</body>
</html>