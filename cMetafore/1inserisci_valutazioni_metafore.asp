<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
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
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Valuta Metafore</title>
<script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

location.href="studente_domande.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"

//location.href=window.history.back();
 }
 </script>
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../Home.asp"
//location.href=window.history.back();
 }
 </script>
 
 <script language="javascript" type="text/javascript" src="../js/seleziona_stampa.js"> 
 
 </script>
 
</head>

<%
  Response.Buffer = true
 'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO,i
  Dim ConnessioneDB,rsTabella, QuerySQL, QuerySQL1,CodiceTest,StringaConnessione
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	 <!-- #include file = "../var_globali.inc" -->
<%	 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'CodiceAllievo=Request.QueryString("cod")
  'cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
   'CodiceDomanda=Request.QueryString("CodiceDomanda")
 
    'response.write(Capitolo)
  Paragrafo=Request.QueryString("Paragrafo")
   TitoloParagrafo=Request.QueryString("TitoloParagrafo")
   if Paragrafo="" then
     Paragrafo=TitoloParagrafo
	end if 
 
  'TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  'response.write("TP:"&ID_Paragrafo)
  Capitolo=Request.QueryString("Capitolo")
  Modulo=Request.QueryString("Modulo")
  if Modulo="" then
    Modulo=Capitolo
  end if
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update

  Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("ID_MOD")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  
  '07/10/14  lo tolgo perchè da errore e poi Classe serve solo nelle query da rivedere
 ' if left(Cartella,1)<>"" then
   '  Classe=Cint(left(Request.QueryString("Cartella"),1))
 ' end if
  Classe=1 ' per non lasciarlo vuoto
  '----
' CONTROLLO PER DISTINGUERE IL TIPO DI METAFORA
 
Select Case Codice_Test
	Case Cartella&"_U_2_3" ' metafora topolino METAFORA TOPOLINO
	 
		if (CodiceAllievo<>"") then  ' se sono stata chiamata dalla pagina studente_domande, valuterò solo le domande di quello studente
			 if (Nulle<>"") then ' se devo mostrare solo quelle con voto=0
					  if (Tutte<>"") then
						  QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where Voto=0 and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"
						else
							 QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where ID_MOD='"& ID_MOD &"' and Voto=0 and CodiceAllievo='"&CodiceAllievo&"';"
						end if 
			  else	    
					if (Data<>"") then ' se devo mostrare sollo quelle dopo una certa data
						 if (Tutte<>"") then
							 QuerySQL="SELECT * FROM Elenco_Metafore_Topolino WHERE " &_ 
							  "  (Data>= CONVERT(DATETIME,'" &Data  &"', 104))  and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"     
							 
					     else
							 QuerySQL="SELECT Elenco_Metafore_Topolino.*, Elenco_Metafore_Topolino.Data FROM Elenco_Metafore_Topolino WHERE "&_ 
							   "  (Data>= CONVERT(DATETIME,'" &Data  &"', 104))  and ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"';"
					     end if
					else
					    if (Tutte<>"") then
							QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where  CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"
						else' IN GENERE ESEGUE QUESTA RIGA
							 QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"';"
						end if	
					end if
				  end if 
		else ' se codiceallievo=""
			if (Gruppi<>"") and (Nulle<>"") then ' mostro le domande per gruppo solo quelle con voto =0 
		  'response.write("QUI")
			QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &" and Voto=0;"
			else
			   if (Gruppi<>"") then
				   QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &";"
			   else
				  if (Nulle<>"") then
						QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where ID_Paragrafo='"& Paragrafo &"' and Voto=0"
				 else	        
					 if (Data<>"") then	 
					   QuerySQL="SELECT Elenco_Metafore_Topolino.*, Elenco_Metafore_Topolino.Data FROM Elenco_Metafore_Topolino WHERE "&_ 
					   "  (Data>= CONVERT(DATETIME,'" &Data  &"', 104))  and  AND ID_Paragrafo='"& Paragrafo &"';"
					else' IN GENERE ESEGUE QUESTA RIGA
					  QuerySQL="SELECT * FROM Elenco_Metafore_Topolino Where ID_Paragrafo='"& Paragrafo &"'"
					end if
				  end if 
			  end if  
		
			end if 
			
		end if 
 
 Case Cartella&"_U_2_5" '  METAFORA Navigazione

		if (CodiceAllievo<>"") then  ' se sono stata chiamata dalla pagina studente_domande, valuterò solo le domande di quello studente
			 if (Nulle<>"") then ' se devo mostrare solo quelle con voto=0
					  if (Tutte<>"") then
						  QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where Voto=0 and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' order by CodiceMetafora;"
						else
							 QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where ID_MOD='"& ID_MOD &"' and Voto=0 and CodiceAllievo='"&CodiceAllievo&"' order by CodiceMetafora;"
						end if 
			else	        
					if (Data<>"") then ' se devo mostrare sollo quelle dopo una certa data
						if (Tutte<>"") then
							 QuerySQL="SELECT Elenco_Metafore_Navigazione.*, Elenco_Metafore_Navigazione.Data FROM Elenco_Metafore_Navigazione WHERE Elenco_Metafore_Navigazione.Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"#  and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' order by CodiceMetafora;"
					   else
							 QuerySQL="SELECT Elenco_Metafore_Navigazione.*, Elenco_Metafore_Navigazione.Data FROM Elenco_Metafore_Navigazione WHERE Elenco_Metafore_Navigazione.Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"#  and ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"' order by CodiceMetafora;"
					   end if
					else
					   if (Tutte<>"") then
							QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where  CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' order by CodiceMetafora;"
						else
							 QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"' order by CodiceMetafora;"
						end if	
					end if
				  end if 
		else
			if (Gruppi<>"") and (Nulle<>"") then ' mostro le domande per gruppo solo quelle con voto =0 
		  'response.write("QUI")
			QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &" and Voto=0;"
			else
			   if (Gruppi<>"") then
				   QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &";"
			   else
				  if (Nulle<>"") then
						QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where ID_Paragrafo='"& Paragrafo &"' and Voto=0"
				 else	        
					 if (Data<>"") then	 
					   QuerySQL="SELECT Elenco_Metafore_Navigazione.*, Elenco_Metafore_Navigazione.Data FROM Elenco_Metafore_Navigazione WHERE" &_  
					    " (Data>= CONVERT(DATETIME,'" &Data  &"', 104))  and  AND ID_Paragrafo='"& Paragrafo&"' order by CodiceMetafora;"
					else
					  QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione Where ID_Paragrafo='"& Paragrafo &"' order by CodiceMetafora;"
					end if
				  end if 
			  end if  
		
			end if 
			
		end if 
 
 		
 
 Case Else
 
				   ' Istruzioni di default
  End Select               				
'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_Paragrafo='"& Paragrafo &"'"

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logMetafore.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close

response.write(QuerySQL)				
Set rsTabella = ConnessioneDB.Execute(QuerySQL)		
		
QuerySQL1=QuerySQL

  ' per il privato
    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaP = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabellaP.fields("Privato") 
	rsTabellaP.close
	Condiviso=request.QueryString("Condiviso")
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) or (Condiviso=1) then  ' else è alla fine
   
%> 

<body bgcolor="#FFFFFF">

<div id="container">
<div class="immagini" style="height:auto">
<!---------->
<%if (session("Admin")=true) then ' per selezionare solo quelle con voto=0 %>
	<form method="POST" form action="1inserisci_valutazioni_metafore.asp?Nulle=1&Tutte=<%=Tutte%>&Gruppi=<%=Gruppi%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
	  <b>Seleziona metafore </b><br>
	  <br>
	<b>Da valutare </b>
	 <input type="submit" value="Voto=0" name="B1"> </p> 
	</form> 
<!-- selezione in base alla data-->
<form method="POST" form action="1inserisci_valutazioni_metafore.asp?Gruppi=<%=Gruppi%>&Tutte=<%=Tutte%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
<b>Data da:</b>
<%if data<>"" then%>
<input type="text" name="txtDATA" value="<%=Data%>" size="10">
<% else%>
<input type="text" name="txtDATA" value="gg/mm/aaaa" size="10">
<% end if%>
 <input type="submit" value="Invia" name="B1"> </p> 
</form>
</div>
<%end if %>
 
  
  <form method="POST" name="dati" action="1inserisci_valutazioni_metafore1.asp?NumRec=<%=i%>&TitoloParagrafo=<%=TitoloParagrafo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&ID_MOD=<%=ID_MOD%>&CodiceTest=<%=Codice_Test%>&CodiceAllievo=<%=CodiceAllievo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <div id="bloc_destra_cont">
 
 <br><br>
  <font size="4" color="#FF0000"><b>Valuta o modifica </b></font><strong><font color="#FF0000" size="4">le Metafore </font></strong><b><font color=#FF0000 size="4"> :</font></b>
  <br>
  <p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Modulo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (TitoloParagrafo) %></b></font> <!-- stampa il titolo del test -->
	
    <p>
	
	<!-- <div class="contenuti_login" >	-->
	<%
	i=0
	'response.write(QuerySql) 
    do while not rsTabella.eof   
	'response.write("<br>CA="&CodiceAllievo)%>
    <br><div class="contenuti_login" style="width: 1000px; height: auto;">	  
	<p><hr><br> 
			<tr><td><b><%=rsTabella(2)%></b></td></tr>
              <input type="text" name="txtCodiceMetafora<%=i%>"  tabindex="<%=(7*i)%>" value="<%=rsTabella.Fields("CodiceMetafora")%>" size="10" maxlength="250">
              <b>Codice Metafora </b> 
			  <input type="text" name="txtDATA<%=i%>" value="<%=rsTabella.Fields("Data")%>" size="8" maxlength="250">
              <b>Data</b>
			  <input type="text" name="txtOraMetafora<%=i%>" value="<%=rsTabella.Fields("Ora")%>" size="6" maxlength="250">
              <b>Ora</b> 
			  
			  <br>
          
	 
	<% Select Case Codice_Test
	Case Cartella&"_U_3_3" ' metafora topolino METAFORA TOPOLINO%>
	  
    
    
    
 <%   
    ID=rsTabella("CodiceMetafora")  
' devo controllare se CodiceMetafora esiste nella tabella dei linkNavigazione, se compare in  in tal caso leggo la L1 ed in quella posizione invece dell'ancora metto href
										  '0		   1		 2			3		4			5          6
				QuerySql="Select * FROM LinkTopolino WHERE Id_n1="&ID&";"
				Set rsLink = ConnessioneDB.Execute(QuerySQL)
				
				If rsLink.BOF=True And rsLink.EOF=True Then  ' se la metafora non compare nella tabella link allora metto tutte ancore
				%>
<p><input type="text" name="txtTopolino<%=i%>"  value="<%=rsTabella.Fields("Topolino")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250"><b><a name="<%=ID%>_1"> Topolino</a> <br>
</b></p> 

 <p><input type="text" name="txtR1Formaggio<%=i%>" value="<%=rsTabella.Fields("Formaggio")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150"><b><a name="<%=ID%>_2"> Formaggio</a> </b></p>
  <p>
	<input type="text" name="txtR1Fame<%=i%>" value="<%=rsTabella.Fields("Fame")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150"><b><a name="<%=ID%>_3"> Fame</a> </b></p>
  <p>
	<input type="text" name="txtR1Labirinto<%=i%>" value="<%=rsTabella.Fields("Labirinto")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"><b><a name="<%=ID%>_4"> Fame</a> </b></p>
   <p><input type="text" name="txtR1Strada<%=i%>" value="<%=rsTabella.Fields("Strada")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150"><b> <a name="<%=ID%>_5">Strada</a> </b></p>
  <p><input type="text" name="txtR1Strada_OK<%=i%>" value="<%=rsTabella.Fields("Strada_OK")%>" tabindex="<%=(7*i)+5%>" size="135"><b> <a name="<%=ID%>_6">Strada_OK </a></b></p>
	<p><input type="text" name="txtR1Strada_KO<%=i%>" value="<%=rsTabella.Fields("Strada_KO")%>" tabindex="<%=(7*i)+6%>" size="135"><b><a name="<%=ID%>_7"> Strada_KO</a> </b></p>
	<p><input type="text" name="txtR1Testata<%=i%>" value="<%=rsTabella.Fields("Testata")%>" tabindex="<%=(7*i)+6%>" size="135">
    <b> Testata </b></p>
	<p><input type="text" name="txtR!Distanza<%=i%>" value="<%=rsTabella.Fields("Distanza")%>" tabindex="<%=(7*i)+6%>" size="5">
    <b> Distanza </b></p>

  
    <%else ' metto gli href
	L1=rsLink(2)
	Id_n1=rsLink(1)
	Id_n2=rsLink(3)
	L2=rsLink(4)
	T2=rsLink(6)
	
	 %>
	<p><input type="text" name="txtTopolino<%=i%>"  value="<%=rsTabella.Fields("Topolino")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250">
     <b>
      <%if L1=1 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Topolino</a>
     <%else%>
     Topolino   
	 <%end if%>  
     <br></b></p>
   
     
     <b><p><input type="text" name="txtR1Formaggio<%=i%>" value="<%=rsTabella.Fields("Formaggio")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150">    
      <%if L1=2 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Formaggio</a>
     <%else%>
          Formaggio
	 <%end if%> 
    </p></b> 
    <b><p><input type="text" name="txtR1Fame<%=i%>" value="<%=rsTabella.Fields("Fame")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150">  
    <%if L1=3 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Fame </a>
     <%else%>
          Fame  
	 <%end if%> 
	 </p></b>
     
     <p><b><input type="text" name="txtR1Luogo<%=i%>" value="<%=rsTabella.Fields("Labirinto")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"> 
     <%if L1=4 then%>  
		<a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Labirinto </a>
     <%else%>
          Labirinto  
	 <%end if%> 
     </b></p>
     
  <p><b><input type="text" name="txtR1Strada<%=i%>" value="<%=rsTabella.Fields("Strada")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150">   
   <%if L1=5 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada </a>
     <%else%>
          Strada  
	 <%end if%> 
     </b></p>
 
  <p><input type="text" name="txtR1Strada_OK<%=i%>" value="<%=rsTabella.Fields("Strada_OK")%>" tabindex="<%=(7*i)+5%>" size="135">  <b> 
    <%if L1=6 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada_OK </a>
     <%else%>
          Strada_OK  
	 <%end if%> 
     </b></p>

<p><input type="text" name="txtR1Strada_KO<%=i%>" value="<%=rsTabella.Fields("Strada_KO")%>" tabindex="<%=(7*i)+6%>" size="135">
<b>
     <%if L1=7 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada_KO </a>
     <%else%>
          Strada_KO  
	 <%end if%> 
     </b></p>
	
 
    
    <p><input type="text" name="txtR1Distanza<%=i%>" value="<%=rsTabella.Fields("Distanza")%>" tabindex="<%=(7*i)+6%>" size="5"><b>
	Distanza </b></p>
                        
<%end if%>	

      <a target="_blank" title="Esegui simulazione Topolino" href="6_simula_metafora_topolino.asp?CodiceMetafora=<%=rsTabella("CodiceMetafora")%>">Simulazione</a><br> <br>


<% 
Case Cartella&"_U_3_5" ' metafora topolino METAFORA NAVIGAZIONE
	
 ID=rsTabella("CodiceMetafora")  
' devo controllare se CodiceMetafora esiste nella tabella dei linkNavigazione, se compare in  in tal caso leggo la L1 ed in quella posizione invece dell'ancora metto href
										  '0		   1		 2			3		4			5          6
				QuerySql="Select * FROM LinkNavigazione WHERE Id_n1="&ID&";"
				Set rsLink = ConnessioneDB.Execute(QuerySQL)
				
				If rsLink.BOF=True And rsLink.EOF=True Then  ' se la metafora non compare nella tabella link allora metto tutte ancore
				%>
<p><input type="text" name="txtAutista<%=i%>"  value="<%=rsTabella.Fields("Autista")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250"><b><a name="<%=ID%>_1"> Autista</a> <br>
</b></p> 
  <p><input type="text" name="txtR1Destinazione<%=i%>" value="<%=rsTabella.Fields("Destinazione")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150"><b> <a name="<%=ID%>_2">Destinazione</a></b></p> 
  <p>
	<input type="text" name="txtR1Carburante<%=i%>" value="<%=rsTabella.Fields("Carburante")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150"><b> <a name="<%=ID%>_3">Carburante </a></b></p>
  <p>
	<input type="text" name="txtR1Luogo<%=i%>" value="<%=rsTabella.Fields("Luogo")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"><b> <a name="<%=ID%>_4">Luogo </a></b></p>
  <p><input type="text" name="txtR1Strada<%=i%>" value="<%=rsTabella.Fields("Strada")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150"><b> <a name="<%=ID%>_5">Strada</a> </b></p>
  <p><input type="text" name="txtR1Strada_OK<%=i%>" value="<%=rsTabella.Fields("Strada_OK")%>" tabindex="<%=(7*i)+5%>" size="135"><b> <a name="<%=ID%>_6">Strada_OK </a></b></p>
	<p><input type="text" name="txtR1Strada_KO<%=i%>" value="<%=rsTabella.Fields("Strada_KO")%>" tabindex="<%=(7*i)+6%>" size="135"><b><a name="<%=ID%>_7"> Strada_KO</a> </b></p>
	<p><input type="text" name="txtR1Cespugli<%=i%>" value="<%=rsTabella.Fields("Cespugli")%>" tabindex="<%=(7*i)+6%>" size="135"><b> <a name="<%=ID%>_8">Cespugli</a> </b></p>
	<p><input type="text" name="txtR1Lupo<%=i%>" value="<%=rsTabella.Fields("Lupo")%>" tabindex="<%=(7*i)+6%>" size="135"><b> 
	<a name="<%=ID%>_9">Lupo </a></b></p>
	<p><input type="text" name="txtR1Cestino<%=i%>" value="<%=rsTabella.Fields("Cestino")%>" tabindex="<%=(7*i)+6%>" size="135"><b>
	<a name="<%=ID%>_10">Cestino</a> </b></p>
	<p><input type="text" name="txtR1Distanza<%=i%>" value="<%=rsTabella.Fields("Distanza")%>" tabindex="<%=(7*i)+6%>" size="5"><b>
	Distanza </b></p>
   
    <%else ' metto gli href
	L1=rsLink(2)
	Id_n1=rsLink(1)
	Id_n2=rsLink(3)
	L2=rsLink(4)
	T2=rsLink(6)
	
	 %>
	<p><input type="text" name="txtAutista<%=i%>"  value="<%=rsTabella.Fields("Autista")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250">
     <b>
      <%if L1=1 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Autista</a>
     <%else%>
     Autista   
	 <%end if%>  
     <br></b></p>
   
     
     <b><p><input type="text" name="txtR1Destinazione<%=i%>" value="<%=rsTabella.Fields("Destinazione")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150">    
      <%if L1=2 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Destinazione</a>
     <%else%>
          Destinazione
	 <%end if%> 
    </p></b> 
    <b><p><input type="text" name="txtR1Carburante<%=i%>" value="<%=rsTabella.Fields("Carburante")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150">  
    <%if L1=3 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Carburante </a>
     <%else%>
          Carburante  
	 <%end if%> 
	 </p></b>
     
     <p><b><input type="text" name="txtR1Luogo<%=i%>" value="<%=rsTabella.Fields("Luogo")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"> 
     <%if L1=4 then%>  
		<a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">Luogo </a>
     <%else%>
          Luogo  
	 <%end if%> 
     </b></p>
     
  <p><b><input type="text" name="txtR1Strada<%=i%>" value="<%=rsTabella.Fields("Strada")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150">   
   <%if L1=5 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada </a>
     <%else%>
          Strada  
	 <%end if%> 
     </b></p>
 
  <p><input type="text" name="txtR1Strada_OK<%=i%>" value="<%=rsTabella.Fields("Strada_OK")%>" tabindex="<%=(7*i)+5%>" size="135">  <b> 
    <%if L1=6 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada_OK </a>
     <%else%>
          Strada_OK  
	 <%end if%> 
     </b></p>

<p><input type="text" name="txtR1Strada_KO<%=i%>" value="<%=rsTabella.Fields("Strada_KO")%>" tabindex="<%=(7*i)+6%>" size="135">
<b>
     <%if L1=7 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Strada_KO </a>
     <%else%>
          Strada_KO  
	 <%end if%> 
     </b></p>
	
 <p><input type="text" name="txtR1Cespugli<%=i%>" value="<%=rsTabella.Fields("Cespugli")%>" tabindex="<%=(7*i)+6%>" size="135"><b> 
     <%if L1=8 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Cespugli </a>
     <%else%>
          Cespugli  
	 <%end if%> 
     </b></p>
    
    <p><input type="text" name="txtR1Lupo<%=i%>" value="<%=rsTabella.Fields("Lupo")%>" tabindex="<%=(7*i)+6%>" size="135"><b> 
	<%if L1=9 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Lupo </a>
     <%else%>
          Lupo  
	 <%end if%> 
     </b></p>
    
    
    <p><input type="text" name="txtR1Cestino<%=i%>" value="<%=rsTabella.Fields("Cestino")%>" tabindex="<%=(7*i)+6%>" size="135"><b>
    <%if L1=10 then%>  
		 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> Cestino </a>
     <%else%>
          Cestino  
	 <%end if%> 
     </b></p>
    
    <p><input type="text" name="txtR1Distanza<%=i%>" value="<%=rsTabella.Fields("Distanza")%>" tabindex="<%=(7*i)+6%>" size="5"><b>
	Distanza </b></p>
                        
<%end if%>	
    
     <a target="blank" title="Esegui simulazione Navigazione" href="6_simula_metafora_navigazione.asp?CodiceMetafora=<%=rsTabella("CodiceMetafora")%>">Simulazione</a> <br><br>
 
<%  Case Cartella&"_U_2_7" 
 
End Select
	 
	
	    Paragrafo=rsTabella(0)
		Modulo=rsTabella.fields("ID_Mod")
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&rsTabella.Fields("CodiceMetafora")&".txt"
   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url=Replace(url,"\","/")
' Open file for reading.

' Set objFSO = CreateObject("Scripting.FileSystemObject")
'			    url1="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFile.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(url)
'				 
'				objCreatedFile.Close
				
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
   sReadAll = objTextFile.ReadAll
	'sReadAll=url
'	response.write(sReadAll)
	 objTextFile.Close	%>
	<b>Spiegazione</b><p><textarea rows="8" name="S1<%=i%>" tabindex="<%=(7*i)+7%>" value="ciao" cols="116"><%=Response.write(sReadAll)%> </textarea></p>
 
     <p><input type="checkbox"  name="cbSegnalata<%=i%>"> <b> Segnalata </b><br><br>
    
<%if (session("Admin")=true) then %>
 
 <p><input type="text" name="txtVAl<%=i%>" value="<%=rsTabella.Fields("Voto")%>" size="1"  ><b> 
	Valutazione </b> </p>
	<p> <!--<input type="text" name="txtINQUIZ<%=i%>" value="<%=rsTabella.Fields("In_Quiz")%>" size="1" ><b> In Quiz </b></p>-->
  <!--Definisce i due bottoni del form -->
   
<p><!--<input type="submit" value="Invia" name="B1"> </p> -->
<% else 
   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
 <p><input type="text" disabled="disabled" name="txtVAl<%=i%>" value="<%=rsTabella.Fields("Voto")%>" size="1"><b> 
	Valutazione </b></p>
   
    <p>
	 <input type="text" visible="false" name="txtINQUIZ<%=i%>" value="<%=rsTabella.Fields("In_Quiz")%>" size="1"><b> In Quiz </b>
	</p>
   
    <a href="javascript:history.back()">	Indietro </a>
<% end if 
end if %>
<p><input type="checkbox"  name="cbStampa<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b> 
	Seleziona per la stampa </b><br>
<%    i=i+1
    rsTabella.movenext
	%></div><%
loop
%><br>
<img src="../../img/printer.jpg" alt="Stampa schede selezionate" onClick="stampa();">
&nbsp;<br>
<b>Stampa <input type="text" name="txtNUMREC" value="<%=i%>" size="1" style="border:none">Frasi</b></p>

<input type="button" value="Seleziona tutti" onClick="checkTutti()">
<input type="button" value="Deseleziona tutti" onClick="uncheckTutti()"><br><hr>

<%if Session("Admin")=true then%>
<b>Voto</b><input type="text"   name="txtVoto" size="1">
<input type="button" value="Valuta tutti" onClick="valutaTutti()" >
<input type="text" name="txtNUMVAL" value="<%=i%>" size="1" style="border:none"><br><hr>
 <input type="text" name="txtNUMREC" value="<%=i%>" size="1"> <b>Totale Metafore</b>
 <p><input type="submit" value="Invia" name="B1"> </p> <!--Definisce i due bottoni del form -->
   <br>
   <%end if%>
</form> <!-- Chiude l'interfaccia -->
<!--#include file="../include/tornaquaderno.html" --> 
<!-- </div>-->
<!-- Deseleziono tutto e poi seleziono solo i segnalati-->
<script language="javascript" type="text/javascript" >
    deselezionatutticheckbox();
 </script>
 <%' rimetto da capo il recordset per selezionare le cb segnalate dopo averle resettate tutte
   Set rsTabella = ConnessioneDB.Execute(QuerySQL1)	
  ' response.write(QuerySQL)
   rsTabella.movefirst 
	i=0 
	do while not rsTabella.eof  
	  if rsTabella("Segnalata")="1" then
	  %>
	    <script language="javascript" type="text/javascript" > 
  		  with (document.dati) {
		  for (var i=1; i <= elements.length; i++) { 
		 if (elements[i].name == 'cbSegnalata'+<%=i%>)  
		    {
		    elements[i].checked = true; 	 
			}
	 }
    }
 </script>
	  <%end if
	  i=i+1
	 
	  rsTabella.movenext
	loop
 
 %>
</body>
<% else%> 
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if %>
</html>