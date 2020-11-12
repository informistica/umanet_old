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
<title>Stampa frasi</title>
<link rel="stylesheet" type="text/css" href="../stile.css">
 <link rel="stylesheet" type="text/css" href="css_stampa.css">
<style>
<!--
 li.MsoNormal
	{mso-style-parent:"";
	margin-bottom:.0001pt;
	font-size:14.0pt;
	font-family:"Times New Roman";
	margin-left:0cm; margin-right:0cm; margin-top:0cm}
<!--Il layout della pagina in fase di stampa sarà quello di un normale foglio A4 con margini di 3cm su tutti i lati.-->
@page {size: 210mm 297mm; margin: 30mm;}
<!--e vitare che l'interruzione avvenga nel corpo della tabella ->     
table {page-break-inside: avoid;}

<!-- larghezza del div che contiene la tabella (facoltativo): il valore potrebbe essere omesso se vale 100% oppure se viene definito altrove -->
.table-responsive {width: 95%;}

<!-- stile del bordo per la tabella (facoltativo) -->
.table-responsive table {border: #ccc solid 1px;}

<!--  istruzioni per le celle (alcune sono obbligatorie)-->
.table-responsive table td, .table-responsive table th 
{min-width: 50px; width: 24%; border: #ccc solid 1px; word-break: break-all; text-align: center; padding: 1%;}

<!--  larghezza delle immagini (facoltativo) -->
.table-responsive table td img {max-width: 50%;}
     

#content {
	float:center;
	width:95%; 
	padding:2% 2% 2% 12%;
}
 
.flex {max-width: 100%}	 
	                

</style>
<meta http-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

 <script type="text/javascript">
window.onload=function() {
'window.print();
}
</script>
 

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
  '-----
 ' Data=Request.Form("txtDATA")
'  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0

'  ID_MOD=Request.QueryString("ID_MOD")
'  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
'  if left(Cartella,1)<>"" then ' DA SISTEMARE NELLE QUERY PER I GRUPPI !!!!!!!!!!!!!
'     Classe=clng(left(Request.QueryString("Cartella"),1))
'  end if
'  
'  
 
'
 

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\wwwroot\anno_2012-2013\logStampaFrasi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL &"---" & Modulo &"-----"& Cartella&"---")
'				objCreatedFile.Close
'				

%>
   

<body bgcolor="#FFFFFF">
<div id="container">  
 <div id="bloc_destra_cont" class="contenuti_login" style="width:95%">
<%

 'response.write(QuerySQL)	
 if (InStr(QuerySQL,"drop")=0) and (InStr(QuerySQL,"delete")=0) then
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
				 <b><center> <font size="+3">  <%=rsTabella(0)%></font> </center></b>
				 <hr>
			   <% 
			  end if
	     end if
		' response.write("ciao:"&rsTabella("Id_Sottoparagrafo")&"pippo")
				if (rsTabella("Id_Sottoparagrafo")<>"") then
					' response.write("ciao1")
					if (StrComp(Sottoparagrafo, rsTabella("Id_Sottoparagrafo")) <> 0)  then
					 
					  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
					   
					   querySqlSotto="select Titolo,Id_Sottoparagrafo from Sottoparagrafi where Id_Sottoparagrafo='"&rsTabella("Id_Sottoparagrafo")&"'"
				 
					 set rsTabellaSotto=ConnessioneDB.execute (querySqlSotto)
					  ' response.write(querySqlSotto)	
					   Sottoparagrafo=rsTabellaSotto("Id_Sottoparagrafo")
						
						%>
						<b><center> <font size="+3"><%=rsTabellaSotto("Titolo")%></font> </center></b> 
					 <%end if%>
				<%end if%>  
  <%  ID=rsTabella("CodiceDomanda")
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
   url=Replace(url,"\","/")
 
                ' Set objFSO = CreateObject("Scripting.FileSystemObject")
			'	url2="C:\Inetpub\wwwroot\anno_2012-2013\logDomande.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url2, True)
				'objCreatedFile.WriteLine(url)
				'objCreatedFile.Close
'response.write(url) 
 
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
 
sReadAll = objTextFile.ReadAll
'sReadAll = url
objTextFile.Close   ' la soluzione seguente la rimuovo e dirò di copiare ed incollare la domanda plus nella spiegazione
' così da avere il livello di apprendimento comprensibile , diversamente dovrei prevedere il modo di far apparire il testo della domanda plus 
' anche nell'approfondimento di fine quiz.
'if clng(rsTabella.fields("Tipo"))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
'	    url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
'		url=Replace(url,"\","/")
'		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
'		sReadAll1 = objTextFile.ReadAll
'		objTextFile.Close
'end if
			 
%>

  
  
    
   <div class="table-responsive" style="width: 100%;">
  <table border="0"  align=center width="60%" id="blugradient1">
		<tr>
			 
			<td style="width:auto"><b><font size="+2"> <%=k&") "&rsTabella("Quesito") %> </b></td>
			<td> </td>
            <td> </td>
             
		</tr>
          <% if session("admin")=true then%>
                    <%if sint="" then%>      
                   <%if rsTabella("VF")=1 then%>
				   <%end if%>
                    <tr>
                    <td> <%if rsTabella("RispostaEsatta")=1 then response.write("Vera") else response.write("Falsa") end if%></td>                             
                    </tr>
                   <%end if%>  
                   
                    <%if ((rsTabella("VF")=0) and (rsTabella("Multiple")=0)) then%>
                    <tr><td colspan="3">1) <%=rsTabella("Risposta1")%></td></tr> 
                    <tr><td colspan="3">2) <%=rsTabella("Risposta2")%></td></tr> 
                    <tr><td colspan="3">3) <%=rsTabella("Risposta3")%></td></tr> 
                    <tr><td colspan="3">4) <%=rsTabella("Risposta4")%></td></tr> 
					<%if sint="" then%>
                    <tr><td>Risposta esatta N.<%=rsTabella("RispostaEsatta")%><td></tr> 
                    <%end if%>                          
                   <%end if%>  
        
               
                 
         <% end if %>
		
		<% if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
	    <tr><td colspan="3"><p align="center">
			 <textarea rows="<%=1+round((len(domandaplus()))/60)%>" name="TestoDomandaPlus0" value="ciao" cols="100"><%
			 
			 
			 Response.write(domandaplus())%> </textarea><br></td></tr><br>
        <%end if %>
   
		<tr>
			<td colspan=3>
			
			<p align="center">
			 <textarea rows="<%=1+round((len(sReadAll))/60)%>" name="TestoDomandaPlus" value="ciao" cols="100"><%
			 ' if clng(rsTabella(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			
			 Response.write(sReadAll)%> </textarea><br>
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