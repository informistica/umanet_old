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

<TITLE>TEORIA DEL TEST</TITLE>
</HEAD>
<BODY> 

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

 Dim ConnessioneDB, rsTabella,rsLink,QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato

    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
  ' 
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  Codice_Test=Request.QueryString("CodiceTest") 
  'response.write("Codice_Test:"& Codice_Test)
'response.Write(Modulo & " " & Capitolo & " " & Paragrafo)
 
' codice per permettere la visualizzazione solo delle proprie domande 
'QuerySQL="Select * from Setting"
'	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
'	Privato=rsTabella.fields("Privato") 
'	rsTabella.close
'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
'  
 
  Dim objFSO, objTextFile
  Dim liv(8) ' serve per indicizzare il chi,cosa,....
  liv(1)="Chi"
  liv(2)="Cosa"
  liv(3)="Dove"
  liv(4)="Quando"
  liv(5)="Come"
  liv(6)="Perchè"
  liv(7)="Quindi"
  
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  						
%>
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color="#FF0000">METAFORA DEL TOPOLINO NEL LABIRINTO:</font></b></p> 
 
<!-- 


 -->

  <table border="0" align=center width="60%">
		<tr>
			<td colspan=3 align=center>
			  <font color="#000000"><b><h3><%=Capitolo%></h3></b></font>
			</td>
		</tr>
		<tr>
			<td colspan=3 align=center>
			  <font color="#000000"><b><h4><%=Paragrafo%></h4></b></font>
			</td>
		</tr>

		
	</table>
 
 
 <table border="1" align=center width="60%">
		<tr>
			<td><font color="#0022FF"><b>Paragrafo</b></font></td>
			<td><font color="#0022FF"><b>Codice Metafora </b></font></td>
			<td><font color="#0022FF"><b>Studente</b></font></td>
		</tr>
	    <tr><td colspan=3><p align="center"><font color="#0022FF"><b>Topolino    : Soggetto che compie l'Azione</b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Formaggio   : Obiettivo</b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Fame   : Motivazione</b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Labirinto : Contesto </b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Strada   : Strategia </b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Strada_OK : Vincente</b></font></td></tr>
		<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Strada_OK : Perdente</b></font></td></tr> 
	    <tr><td colspan=3><p align="center"><font color="#0022FF"><b>Testata : Conseguenze</b></font></td></tr> 
        <tr><td colspan=3><p align="center"><font color="#0022FF"><b>Distanza : Difficoltà</b></font></td></tr> 

	</table>
	<br>
<%   
  QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close
 

costQuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, M_Topolino.CodiceMetafora, Moduli.ID_Mod, M_Topolino.Topolino, M_Topolino.Formaggio, M_Topolino.Fame, M_Topolino.Labirinto, M_Topolino.Strada, M_Topolino.Strada_OK, M_Topolino.Strada_KO, M_Topolino.Testata, M_Topolino.Distanza, M_Topolino.In_Quiz,Paragrafi.Posizione,M_Topolino.Cartella " &_
" From Elenco_Metafore_topolino"
'
' 

'if (cint(Stato)=0) or (cint(Stato0)=0) then  
 if cint(Stato)=0 then
 'Definzione codice SQl della query per ricercare le domande del paragrafo 
  
   if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato  
		QuerySQL=costQuerySQL &_
		 " Where Paragrafi.ID_Paragrafo='" & Codice_Test & "' and M_Topolino.Topolino<>'?' and M_Topolino.Cartella='"&Cartella&"'" &_   
		" ORDER BY Paragrafi.Posizione,M_Topolino.CodiceMetafora;"
	else
	    QuerySQL= costQuerySQL &_
		 " Where Paragrafi.ID_Paragrafo='" & Codice_Test & "' and M_Topolino.Topolino<>'?' and M_Topolino.Cartella='"&Cartella&"'and M_Topolino.Id_Stud='"& Session("CodiceAllievo")&"'" &_   
		" ORDER BY Paragrafi.Posizione,M_Topolino.CodiceMetafora;"
	      
	end if

else 		 

    if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato
		QuerySQL= costQuerySQL &_
		" Where Moduli.ID_Mod='" & Modulo & "' and M_Topolino.Topolino<>'?' " &_ 
		" ORDER BY Paragrafi.Posizione,M_Topolino.CodiceMetafora;"
    else
	    QuerySQL=costQuerySQL &_
		" HAVING Moduli.ID_Mod='" & Modulo & "' and M_Topolino.Topolino<>'?'  and Nodi.Id_Stud='"& Session("CodiceAllievo") &_   
		" ORDER BY Paragrafi.Posizione,M_Topolino.CodiceMetafora;"
	
	end if

end if   

 Set objFSO = CreateObject("Scripting.FileSystemObject")  
  ' 	url="C:\Inetpub\umanetroot\anno_2012-2013\logSpiegazioneTopolino0.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close 

Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
'response.write(querySQL)

 
      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%>
  <H4>Metafore della rete non ancora disponibili!</h4>
  <p><h5><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
</H5>
<% Else
  
	  i=1 'inizializza la variabile i (contatore delle domande)
	  Do until rsTabella.EOF
	  'response.Write(rsTabella(12))
		if (strcomp(rsTabella(12),"12/12/2112")<>0) then  'apro l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
					 
				 
					ID=rsTabella(3)
					url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
					
			   'Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   			url3="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logSpiegazioneTopolino.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url3, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.Close 

					 ' NB c'è una / nell'url locale
				
					' url=Server.MapPath("/ECDL") & "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
					   url1= "../" & Cartella & "/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
				
				url3=Replace(url,"\","/")
				url=url3
				
				'response.write(url)
				' Open file for reading.
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
				'sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
				%>
				<%' devo controllare se ID nodo esiste nella tabella dei link in tal caso leggo la L1 ed in quella posizione invece dell'ancora metto href
										  '0		   1		 2			3		4			5          6
			' LINK ****************
			
				QuerySql="Select LinkTopolino.ID_LinkTopolino, LinkTopolino.Id_n1, LinkTopolino.L1, LinkTopolino.Id_n2, LinkTopolino.L2, LinkTopolino.Id_Stud,LinkTopolino.Testo2 FROM LinkTopolino WHERE Id_n1="&ID&";"
				 
			
				Set rsLink = ConnessioneDB.Execute(QuerySQL)
				If rsLink.BOF=True And rsLink.EOF=True Then  ' se il nodo non compare nella tabella link allora metto tutte ancore
				%>
			
					  <table border="1"  align=center width="63%">
							<tr>
							  <td width="13%"><b>Metafora n</b>.<%=rsTabella(3)%></td>
							  <td width="18%"><%=rsTabella.fields("Titolo")%></td>
							  <td width="69%"><%=rsTabella.fields("Cognome")%></td>
							  <td><a href="6_simula_metafora_topolino.asp?CodiceMetafora=<%=rsTabella.fields("CodiceMetafora")%>" title="Esegui simulazione">Simula</a></td>
							  <td width="auto"><a href="inserisci_metafora_patente.asp?CodiceMetafora=<%=rsTabella.fields("CodiceMetafora")%>&CodiceTest=U_3_5&Capitolo=Interfaccia UWWW&Paragrafo=Navigazione nella Rete della Vita&daTopolino=1" title="Interpreta nella  Metafora della Navigazione">Invia Patente</a></td>
							</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
							<tr><td><b><a name="<%=ID%>_1">Soggetto</a></b></td><td colspan=4><p align="center"><b><%=ucase(rsTabella.fields("Topolino"))%></b></td></tr>
							<tr><td><b><a name="<%=ID%>_2">Obiettivo</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Formaggio"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_3">Motivazione</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Fame"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_4">Contesto</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Labirinto"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_5">Strategia</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Strada"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_6">Vincente</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Strada_OK"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_7">Perdente</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Strada_KO"))%></td></tr>
							<tr><td><b><a name="<%=ID%>_8">Conseguenze</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Testata"))%></td></tr>
								<tr><td><b><a name="<%=ID%>_8">Difficoltà</a></b></td><td colspan=4><p align="center"><%=ucase(rsTabella.fields("Distanza"))%></td></tr>
							<tr>
								<td colspan=4>
								<p align="center"><%=sReadAll%></td>
							</tr>
	</table>
				<br>
				<%else ' 
				
				'devo mettere href nel livello indicato 
				'************** DA SISTEMARE %> 
					
					
					<table border="1"  align=center width="62%">
							<tr>
								<td width="10%"><b>Metafora n</b>.<%=rsTabella.fields("CodiceMetafora")%></td>
							  <td width="16%"><%=rsTabella.fields("Titolo")%></td>
							  <td width="60%"><%=rsTabella.fields("Cognome")%></td>
								<td width="14%">Link to   </td>
							</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
							
							<%' per ogni livello di ogni nodo vedo i link che ha ad altri nodi, e metto una stellina per ognuno
							  ' per ogni livello controllo il rsLink, se trovo che il livello è coinvolto in un link metto href, la prima volta metto il <td> le altre aggiungo allo stesso <td>
							   for i=1 to 9
								primo=0
								primo1=0 %>
							   <tr><td><b><a name="<%=ID%>_<%=i%>" title="<%=ID%>_<%=i%>"><%=liv(i)%></a></b></td><td colspan=2><p align="center"><%=rsTabella(4+i)%> </td>
											
								<%	 rsLink.Movefirst()
									 Do until rsLink.EOF
											L1=rsLink(2)
											Id_n1=rsLink(1)
											Id_n2=rsLink(3)
											L2=rsLink(4)
											T2=rsLink(6)
										   if i=L1 then
												 if primo=0 then 
													primo=1 %>
													<td><a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">></a>
												<%else%>
													 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">></a>
												<%end if%>  
										   <%end if  
										  rsLink.Movenext()
										Loop%>
								</td></tr>
							  <% next
								
							 %>
							 
							<tr>
								<td colspan=4>
								<p align="center"><%=sReadAll%></td>
							</tr>
				</table>
				<br>	
				
				<%end if %>
			<%
			
       i = i+ 1 
	   	end if  'chiudo l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
	
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
  <div class="citazioni">
  <a href="javascript:history.back()">	Indietro </a>
  </div>

  </div>
  </div>
</BODY>
<% 'else 
  ' Response.Redirect "../home.asp"
      'end if %>
</HTML>