<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 

d=request.querystring("cla")
cod=request.querystring("cod")
cod=Request.QueryString("cod")

 'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query
 Dim ConnessioneDB,ConnessioneDB1, rsTabella,rsTabella1, QuerySQL,QuerySQL1 ,CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione
 'StringaConnessione= Response.Cookies("Dati")("StrConn")

   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
	'ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
    '          "DBQ=" & Server.MapPath("../database/Copiaditestonline.mdb")
    'ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
    '              "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2008-2009/ECDL/database/Copiaditestonline.mdb"    
  '

   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   </body>
<center>
<p>
 
</p>
<div class="citazioni" ><div> <span style="font-style: normal">

<b><font size="3">GESTIONE</font>&nbsp;</b> </span></div>
<hr>
<%If d="" and cod="" then%>
					<p><br><br>&nbsp;<table border=1 align=center bordercolor=pink >
	<tr><td colspan=2><b>Scegli la classe</b></td></tr>
	
	<tr><td><a href="../studente_domande_gruppi.asp?cla=3">3 Pc</a></td></tr>
	<tr><td><a href="../studente_domande_gruppi.asp?cla=4">4 Pa</a></td></tr>
	
	<!--<tr><td><a href="studente_domande_gruppi.asp?cla=6">Admin</a></td></tr>
	 -->
	</table>
<%else
	if cod="" then
		QuerySQL="SELECT Url, Data, Descrizione FROM VERIFICHE Where Classe='"& d &"'"
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)%> 
		<br><br>&nbsp
		<table border=1  align=center bordercolor=pink >
		<tr><td colspan=2>
			<p align="center"><b>Scegli  </b></td></tr>
	    <%
		do while not rsTabella.eof%> 
			<tr><td><a href="<%=rsTabella(0)%>"><%=rsTabella(1) & " " & rsTabella(2)%></a></td></tr>
			<%
			rsTabella.movenext
		loop%>
		</table>


	<br>
	 <h5 style="text-align: center" > <span style="font-style: normal"> Hai scelto la classe <%=d%> ora scegli lo studente </h5></span>
	<br><b>
	<% response.Write("Classifica al " &day(date())&"/"&month(date())&"/"&year(date()))%> <br><br></b>
	
	<form method="POST" form action="aggiorna_punteggio.asp?cla=<%=d%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
	 
	 <table align=center border=1 bordercolor=pink>
	<%if (session("Admin")=true) then %>
	<tr><td><b>N.</b></td><td><center><b>Cognome Nome</b><td><b>TOT</b></td><td><b>PD</b></td><td><b>PN</b></td><td><b>Crediti</b></td><td><b>Voto</b></td><td><b>+/-</b></td></tr>
	<%else%>
	<tr><td><b>N.</b></td><td><center><b>Cognome Nome</b><td><b>TOT</b></td><td><b>PD</b></td><td><b>PN</b></td><td><b>Crediti</b></td><td><b>Voto</b></td> </tr>
	
	<%end if %>
	</div>
	
	<%
	QuerySQL="SELECT MAX([TOT]) AS [MAX] FROM PUNTEGGI_STUDENTI WHERE PUNTEGGI_STUDENTI.Classe='" & d & "'"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	max=rsTabella(0) 
	if max=0 then
	  max=1
	end if
	 'QuerySQL="SELECT Cognome, Nome, CodiceAllievo FROM Allievi WHERE Classe='" & d & "' ORDER BY Allievi.Cognome" 
								'			0						1						2							3						4				5							6
	 QuerySQL="SELECT PUNTEGGI_STUDENTI.PUNTI, PUNTEGGI_STUDENTI.Cognome, PUNTEGGI_STUDENTI.Nome, PUNTEGGI_STUDENTI.CodiceAllievo,PUNTEGGI_STUDENTI.Crediti,PUNTEGGI_STUDENTI.TOT,PUNTEGGI_STUDENTI.PN" &_
	" FROM PUNTEGGI_STUDENTI " &_
	" WHERE PUNTEGGI_STUDENTI.Classe='" & d & "'" &_
	" ORDER BY PUNTEGGI_STUDENTI.TOT DESC, PUNTEGGI_STUDENTI.PUNTI DESC"
	
	
	
	
	 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	 i=0
	 do while not rsTabella.eof 
	
	   
	if (session("Admin")=true) then 
	
			 if (fix((rsTabella(5)*8/max) * 10) / 10) <6 then
			  'response.write("ciao") %>
				
				<tr><td><%=i+1%></td><td><a href="../studente_domande.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>">   <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%>  </a></td><td><%=rsTabella(5)%>  </td><td><%=rsTabella(0)%></td><td><%=rsTabella(6)%></td><td><%=rsTabella(4)%></td><font color="#FF0000"><td bordercolor="#FF0000" ><%=fix((rsTabella(5)*8/max) * 10) / 10 %></td></font><td><input type="text" NAME="<%=i%>" value="0" size="1"></td></tr>
			  
			 <%else%>
			  <tr><td><%=i+1%></td><td><a href="../studente_domande.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>">   <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%>  </a></td><td><%=rsTabella(5)%>  </td><td><%=rsTabella(0)%></td><td><%=rsTabella(6)%></td><td><%=rsTabella(4)%></td><td><%=fix((rsTabella(5)*8/max) * 10) / 10 %></td><td><input type="text" NAME="<%=i%>" value="0" size="1"></td></tr>
			 <%end if%><u></u>
         
      <%else 
        'response.write("ciao")
           if (fix((rsTabella(5)*8/max) * 10) / 10) <6 then%>
			<tr><td><%=i+1%></td><td><a href="../studente_domande.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>"><%=rsTabella("Cognome")%><%=rsTabella("Nome")%>  </a></td><td><%=rsTabella(5)%>  </td><td><%=rsTabella(0)%></td><td><%=rsTabella(6)%></td><td><%=rsTabella(4)%></td><td bordercolor="#FF0000"><%=fix((rsTabella(5)*8/max) * 10) / 10 %></td></tr>
<%         else%>
				<tr><td><%=i+1%></td><td><a href="../studente_domande.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>">   <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%>  </a></td><td><%=rsTabella(5)%>  </td><td><%=rsTabella(0)%></td><td><%=rsTabella(6)%></td><td><%=rsTabella(4)%></td><td><%=fix((rsTabella(5)*8/max) * 10) / 10 %></td></tr>
           <%end if 
		end if 
rsTabella.movenext
i=i+1
loop
%>
</table> 
<%
else

'QuerySQL="SELECT Allievi.Cognome,Allievi.Nome, Domande.Id_Arg, Domande.Quesito,Domande.Data,Domande.Risposta1,Domande.Risposta2,Domande.Risposta3,Domande.Risposta4,Domande.RispostaEsatta,Domande.CodiceDomanda " &_
'" FROM Allievi INNER JOIN Domande ON Allievi.CodiceAllievo=Domande.Id_Stud" &_
'" WHERE Allievi.CodiceAllievo='" & cod & "'"
' prelevo l'elenco delle domande dello studente
'                            0                1             2               3            4             5                   6                    7                 8                 9                  10                11                  12             13               14          15              16               17			18             19
QuerySQL="SELECT  Allievi.Cognome, Allievi.Nome,Domande.Id_Arg,Domande.Quesito,Domande.Data,Domande.Risposta1, Domande.Risposta2, Domande.Risposta3, Domande.Risposta4, Domande.RispostaEsatta,Moduli.Titolo, Paragrafi.Titolo, Domande.CodiceDomanda,Moduli.ID_Mod,Domande.Id_Arg,Domande.Voto,Allievi.CodiceAllievo,Domande.Voto,Domande.URL_Teoria,Domande.Cartella,Domande.Tipo,Domande.In_Quiz" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo=Domande.Id_Stud) ON Moduli.ID_Mod=Domande.Id_Mod) ON Paragrafi.ID_Paragrafo=Domande.Id_Arg" &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
" GROUP BY  Allievi.Cognome, Allievi.Nome,Domande.Id_Arg,Domande.Quesito,Domande.Data,Domande.Risposta1, Domande.Risposta2, Domande.Risposta3, Domande.Risposta4, Domande.RispostaEsatta,Moduli.Titolo, Paragrafi.Titolo,Domande.CodiceDomanda,Moduli.ID_Mod,Domande.Voto,Allievi.CodiceAllievo,Domande.Voto,Domande.URL_Teoria,Domande.Cartella,Domande.Tipo,Domande.In_Quiz" &_
" ORDER BY Domande.Id_Arg asc"
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<table align=center border=1 bordercolor=pink>
<tr><td colspan=3><center>Domande di <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%></center> </td></tr>

	<%if (session("Admin")=true) then %>
		<tr><td><b><center>Paragrafo</center></b></td><td><b>Quesito</b></td><td><b>Data</b></td><td><b>Punti</b></td><td><b>Cancella</b></td></tr>
	<%else %>
		<tr><td><b><center>Paragrafo</center></b></td><td><b>Quesito</b></td><td><b>Data</b></td><td><b>Punti</b></td></tr>
	<%end if %>

<%do while not rsTabella.EOF 
'url=Server.MapPath("/anno_2008-2009/ECDL/ECDL/"&rsTabella(13)&"_Spiegazioni/"&rsTabella(13)&"_"&rsTabella(11)&"_"&rsTabella(12)&".txt")

%>
 

	<%if (session("Admin")=true) then %>
		<tr><td><%=rsTabella(2)%></td><td><a href="../cDomande/inserisci_valutazione.asp?Tipodomanda=<%=rsTabella(20)%>&Cartella=<%=rsTabella(19)%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella(14)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>&Quesito=<%=rsTabella(3)%>&R1=<%=rsTabella(5)%> &R2=<%=rsTabella(6)%>&R3=<%=rsTabella(7)%>&R4=<%=rsTabella(8)%>&RE=<%=rsTabella(9)%>&MO=<%=rsTabella(13)%>&VAL=<%=rsTabella(15)%>&URL=<%=rsTabella(18)%>&INQUIZ=<%=rsTabella(21)%>&VALINQUIZ=<%=rsTabella(21)%> "><%=rsTabella(3)%></a></td><td><%=rsTabella(4)%></td><td><%=rsTabella(17)%></td>
		
		<td><a href= "../cDomande/cancella_domanda.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>&Cartella=<%=rsTabella(19)%>&Modulo=<%=rsTabella(13)%>&CodiceTest=<%=rsTabella(14)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>">x</a></td>
	 
	<%else %>
		<tr><td><%=rsTabella(2)%></td><td><a href="../cDomande/inserisci_valutazione.asp?Tipodomanda=<%=rsTabella(20)%>&Cartella=<%=rsTabella(19)%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella(14)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>&Quesito=<%=rsTabella(3)%>&R1=<%=rsTabella(5)%> &R2=<%=rsTabella(6)%>&R3=<%=rsTabella(7)%>&R4=<%=rsTabella(8)%>&RE=<%=rsTabella(9)%>&MO=<%=rsTabella(13)%>&VAL=<%=rsTabella(15)%>&URL=<%=rsTabella(18)%>&INQUIZ=<%=rsTabella(21)%>&VALINQUIZ=<%=rsTabella(21)%>"><%=rsTabella(3)%></a></td><td><%=rsTabella(4)%></td><td><%=rsTabella(17)%></td>
	 
	<%end if %>

 
   
<%
 
rsTabella.movenext
loop
%></table><%
' RIPETO LA STESSA LOGICA PER L?ELENVCO NODI

' prelevo l'elenco delle domande dello studente
'                            0                1             2               3            4             5                   6                    7                 8                 9                   		'  0				1			2			3		4			5			6			7			8			9				10					11			12	
'		  13			14			15				16				17                18
QuerySQL="SELECT Allievi.Cognome, Allievi.Nome, Nodi.Id_Arg, Nodi.Chi, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi, Moduli.Titolo, Paragrafi.Titolo, Nodi.CodiceNodo," &_ 
" Moduli.ID_Mod, Nodi.Voto, Nodi.Data, Allievi.CodiceAllievo, Nodi.URL_Teoria, Nodi.Cartella " &_
" FROM Allievi INNER JOIN (Paragrafi INNER JOIN (Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg) ON Allievi.CodiceAllievo = Nodi.Id_Stud " &_
" WHERE Allievi.CodiceAllievo='" & cod & "';" 

 


 
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<table align=center border=1 bordercolor=pink>
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><td>Nodi della rete non ancora inseriti!</td></tr>
			  
<% Else%>
	
		<tr><td colspan=3><center>Nodi di <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%></center> </td></tr>
		
			<%if (session("Admin")=true) then %>
				<tr><td><b><center>Paragrafo</center></b></td><td><b>Nodo</b></td><td><b>Data</b></td><td><b>Punti</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Paragrafo</center></b></td><td><b>Nodo</b></td><td><b>Data</b></td><td><b>Punti</b></td></tr>
			<%end if %>
		
		<%do while not rsTabella.EOF 
		'url=Server.MapPath("/anno_2008-2009/ECDL/ECDL/"&rsTabella(13)&"_Spiegazioni/"&rsTabella(13)&"_"&rsTabella(11)&"_"&rsTabella(12)&".txt")
 
		%>
	 
	  
			<% 
			    if (session("Admin")=true) then %>
									<tr><td><%=rsTabella(2)%></td><td><a href="../cNodi/inserisci_valutazione_nodi.asp?Cartella=<%=rsTabella(18)%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella(2)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>&Chi=<%=rsTabella(3)%>&Cosa=<%=rsTabella(4)%> &Dove=<%=rsTabella(5)%>&Quando=<%=rsTabella(6)%>&Come=<%=rsTabella(7)%>&Perche=<%=rsTabella(8)%>&Quindi=<%=rsTabella(9)%>&MO=<%=rsTabella(13)%>&VAL=<%=rsTabella(14)%>&URL=<%=rsTabella(17)%> "><%=rsTabella(3)%></a></td><td><%=rsTabella(15)%></td><td><%=rsTabella(14)%></td> 
<td><a href="../cDomande/cancella_domanda.asp?cla=<%=d%>&cod=<%=rsTabella("CodiceAllievo")%>&Cartella=<%=rsTabella(18)%>&Modulo=<%=rsTabella(13)%>&CodiceTest=<%=rsTabella(2)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>">x </a>
</td>
			 
			<%else %>
								<tr><td><%=rsTabella(2)%></td><td><a href="../cNodi/inserisci_valutazione_nodi.asp?Cartella=<%=rsTabella(18)%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella(2)%>&CodiceDomanda=<%=rsTabella(12)%>&Capitolo=<%=rsTabella(10)%>&Paragrafo=<%=rsTabella(11)%>&Chi=<%=rsTabella(3)%>&Cosa=<%=rsTabella(4)%> &Dove=<%=rsTabella(5)%>&Quando=<%=rsTabella(6)%>&Come=<%=rsTabella(7)%>&Perche=<%=rsTabella(8)%>&Quindi=<%=rsTabella(9)%>&MO=<%=rsTabella(13)%>&VAL=<%=rsTabella(14)%>&URL=<%=rsTabella(17)%> "><%=rsTabella(3)%></a></td><td><%=rsTabella(15)%></td><td><%=rsTabella(14)%></td> 
			<%end if %>   
		<% 
		rsTabella.movenext
		loop%>
		
		<%
end if%> 
</table>


<% ' logica per mostrare i risultati nei quiz dello studente relativi ai singoli paragrafi
       '                     0               1            2                 3                4                  5                   6                    7            8
QuerySQL=" SELECT Allievi.Cognome, Allievi.Nome, Moduli.Titolo, Paragrafi.Titolo, Risultati.Data, Risultati.Risultato, Allievi.CodiceAllievo, Risultati.ID_R,Risultati.Ora " &_
" FROM (Allievi INNER JOIN (Risultati INNER JOIN Paragrafi ON Risultati.CodiceTest = Paragrafi.ID_Paragrafo) ON Allievi.CodiceAllievo = Risultati.CodiceAllievo) " &_
" INNER JOIN (Moduli INNER JOIN Classi_Moduli_Paragrafi ON Moduli.ID_Mod = Classi_Moduli_Paragrafi.Id_Modulo) ON Paragrafi.ID_Paragrafo = Classi_Moduli_Paragrafi.Id_Paragrafo " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'  ORDER By Risultati.Data asc;" 

 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<table align=center border=1 bordercolor=pink>
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><td>Non ci sono quiz svolti nei paragrafi!</td></tr>
			  
<% Else%>
		<tr><td colspan=3><center>Risultati nei quiz di <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%></center> </td></tr>
			<%if (session("Admin")=true) then %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td><b>Data</b><td><b>Ora</b></td><td><b>Risultato</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td><b>Data</b></td><td><b>Ora</b><td><b>Risultato</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (session("Admin")=true) then %>
				<tr><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td><td><%=rsTabella(5)%></td><td><a href="../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&IdR=<%=rsTabella(7)%>">x</a></td>
			<%else %>
								<tr><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td> <td><%=rsTabella(5)%></td>
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>

<% ' logica per mostrare le visualizzazioni dello studente relativi ai singoli paragrafi
' faccio una query per prelevare l'elenco dei moduli e paragrafi
' per ogni modulo e paragrafo faccio una query per contare le visualizzazioni e per ogni record aggregato metto il link al dettaglio delle visualizzazioni 
	   '                     0               1            2                 3                4                  5                   6                    7            8

QuerySQL="SELECT DISTINCT (Moduli.Titolo) AS Modulo, Paragrafi.Titolo as Paragrafo, Moduli.ID_Mod, Paragrafi.ID_Paragrafo " &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN Visualizzazioni ON Moduli.ID_Mod = Visualizzazioni.ID_Mod) ON " &_
" Paragrafi.ID_Paragrafo = Visualizzazioni.ID_Paragrafo " &_
" WHERE Visualizzazioni.CodiceAllievo='" & cod & "' ORDER BY Moduli.ID_Mod, Paragrafi.ID_Paragrafo;"
'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni 
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<table align=center border=1 bordercolor=pink>
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><td>Non ci sono visualizzazioni</td></tr>
			  
<% Else%>
	
			 <% 'conto le visualizzazioni totali 
			  QuerySQL1="SELECT Count(*) AS Numero_visualizzazioni "&_
" FROM Allievi INNER JOIN Visualizzazioni ON Allievi.CodiceAllievo = Visualizzazioni.CodiceAllievo" &_
" WHERE Allievi.CodiceAllievo='"&cod &"';"
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
	num_visualizzazioni_totali=rsTabella1(0) %>
	          	<tr><td colspan=3><center>Visualizzazioni procedure in <b>Totale = <%=num_visualizzazioni_totali%></b></center> </td></tr>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td><b>Totale</b></td></tr>
		 
		<% 'adesso per ogni recordset conto le visualizzazioni e aggiungo link per il dettaglio
		   do while not rsTabella.EOF 
		     QuerySQL1="SELECT Count(*) AS Numero_visualizzazioni "&_
" FROM Allievi INNER JOIN Visualizzazioni ON Allievi.CodiceAllievo = Visualizzazioni.CodiceAllievo" &_
" WHERE (Allievi.CodiceAllievo='"&cod &"' and Visualizzazioni.ID_Mod='" &rsTabella(2) &"' and Visualizzazioni.ID_Paragrafo='" &rsTabella(3) &"')"&_
" GROUP BY Visualizzazioni.ID_Mod, Visualizzazioni.ID_Paragrafo;"
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
	num_visualizzazioni=rsTabella1(0) 
			 %>
								<tr><td><a href="../cDomande/dettagli_visualizzazioni.asp?tipo=0&ID_Mod=<%=rsTabella(2)%>&ID_Paragrafo=<%=rsTabella(3)%>&cod=<%=cod%>"><%=rsTabella(0)%></a></td><td><a href="../cDomande/dettagli_visualizzazioni.asp?tipo=1&ID_Mod=<%=rsTabella(2)%>&ID_Paragrafo=<%=rsTabella(3)%>&cod=<%=cod%>"><%=rsTabella(1)%></a></td><td><%=num_visualizzazioni%></td> 
			 
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>


<%End if

End if
'rsTabella.Close()
'Set rsTabella = nothing
if (session("Admin")=true) and d <> "" then %>

<!--</table>-->
 <p><input type="submit" value="Aggiorna" name="B1"><input type="reset" value="Azzera" name="B2"></p> <!--Definisce i due bottoni del form -->
</form> <!-- Chiude l'interfaccia -->
		
<%end if%></i>
<a href="../cGrafici/genera_grafico.asp?cla=<%=d%>">Visualizza grafico</a><br></p>


<h4 style="text-align: center"><i><a href="../../home.asp" >Vai all'HomePage</a> </h4>
</center>
</html>




