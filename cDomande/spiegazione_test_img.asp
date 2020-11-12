<!-- esegui_test_MODBC3.asp -->

<%@ Language=VBScript %>
<%Function url_img(cartella,nome_img)
	 
	 url_img="../img_quiz" & "/" & cartella &"/" & nome_img&".jpg"
 	 'url_img=replace(url_img,"/","\")
End Function %>
<%Function url_video(cartella,nome_video)
	 
	' url_video="../video_quiz" & "/" & cartella &"/" & nome_video&".htm"
	  url_video="../video_quiz"&"/"&cartella &"/" &nome_video
 	 'url_img=replace(url_img,"/","\")
End Function %>

<%Function url_video1(cartella,nome_video)
	 

	  url_video1="../video_quiz"&"/"&cartella &"/" &nome_video & "/"&nome_video
 	 
End Function %>

<%Function url_tutorial(cartella,nome_spiegazione)
	 
	' url_video="../video_quiz" & "/" & cartella &"/" & nome_video&".htm"
	  url_tutorial="../tutorial_quiz"&"/"&cartella &"/" &nome_spiegazione&".htm"
 	 'url_img=replace(url_img,"/","\")
End Function %>
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

<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
<TITLE>IMMAGINI DELLA RETE</TITLE>
</HEAD>
<%Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

 Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato

    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
  
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  CodiceDomanda=Request.QueryString("CodiceDomanda")
  davisualizzazioni=Request.QueryString("davisualizzazioni")' per distinguere il caslo in cui la pag è chiamata da visualizzazioni e in tal caso devo fare la query per prelevare solo una specifica domanda
  
   
%>
 <%'response.write("Stati :  " & stato & " " & stato0)%>
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color="#FF0000">SPIEGAZIONE IMMAGINI:</font></b> 
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
 
 
 <table border="1" align=center width="60%">
		 
		<tr>
		<td width="12%"><font color="#0022FF"><b>Codice Domanda</b></font></td>
			<td width="88%">
			<p align="center"><font color="#0022FF"><b>Azione</b></font>
		  </td>
			 
		</tr>
		<tr>
			<td colspan=2>
			<p align="center"><font color="#0022FF"><b>Procedure</b></font></td>
		</tr>
		<tr>
			<td colspan=2>
			<p align="center"><font color="#0022FF"><b>Video</b></font></td>
		</tr>
		<tr>
			<td colspan=2>
			<p align="center"><font color="#0022FF"><b>Spiegazione</b></font></td>
		</tr>
	</table>
	<br>
<%   
  

if davisualizzazioni=1 then  ' se viene invocata da visualizzazioni cambio query
   QuerySQL="SELECT Domande1.CodiceDomanda, Domande1.NumeroDomanda, Domande1.Quesito, Domande1.Risposta1, Domande1.Risposta2, Domande1.Risposta3, Domande1.Risposta4, Domande1.RispostaEsatta, Moduli.Titolo, Paragrafi.Titolo,Domande1.Video,Domande1.Spiegazione,Domande1.Tipo,Moduli.ID_Mod,Paragrafi.ID_Paragrafo" &_
		" FROM Paragrafi INNER JOIN (Moduli INNER JOIN Domande1 ON Moduli.ID_Mod = Domande1.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande1.Id_Arg " &_   
		" WHERE Paragrafi.ID_Paragrafo='" & Codice_Test & "'" &_
		" and Domande1.CodiceDomanda="&CodiceDomanda&";"
else

		if (clng(Stato)=0) or (clng(Stato0)=0) then 
		' 'Definzione codice SQl della query per ricercare le domande del paragrafo 
					   
		QuerySQL="SELECT Domande1.CodiceDomanda, Domande1.NumeroDomanda, Domande1.Quesito, Domande1.Risposta1, Domande1.Risposta2, Domande1.Risposta3, Domande1.Risposta4, Domande1.RispostaEsatta, Moduli.Titolo, Paragrafi.Titolo,Domande1.Video,Domande1.Spiegazione,Domande1.Tipo,Moduli.ID_Mod,Paragrafi.ID_Paragrafo,Domande.Cartella" &_
		" FROM Paragrafi INNER JOIN (Moduli INNER JOIN Domande1 ON Moduli.ID_Mod = Domande1.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande1.Id_Arg " &_   
		" WHERE Paragrafi.ID_Paragrafo='" & Codice_Test & "'" &_
		" ORDER BY Paragrafi.ID_Paragrafo, Domande1.NumeroDomanda;"
		
		else 
								'	1						2					3					4					5					6					7					8					9					10
		QuerySQL=" SELECT Domande1.CodiceDomanda, Domande1.NumeroDomanda, Domande1.Quesito, Domande1.Risposta1, Domande1.Risposta2, Domande1.Risposta3, Domande1.Risposta4, Domande1.RispostaEsatta, Moduli.Titolo, Paragrafi.Titolo,Domande1.Video,Domande1.Spiegazione,Domande1.Tipo,Moduli.ID_Mod,Paragrafi.ID_Paragrafo,Domande.Cartella" &_
		" FROM Paragrafi INNER JOIN (Moduli INNER JOIN Domande1 ON Moduli.ID_Mod = Domande1.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande1.Id_Arg" &_   
		" WHERE  Moduli.ID_Mod='" & Modulo & "'" &_
		" ORDER BY Paragrafi.ID_Paragrafo, Domande1.NumeroDomanda;"
							
		end if    
  
end if  
				'dim objFSO,objCreatedFile
				'Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'Dim sRead, sReadLine, sReadAll, objTextFile
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="C:\Inetpub\umanetroot\anno_2009-2010\ECDL\database\logim3.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(QuerySQL)
				'objCreatedFile.Close 
	
	    
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 
      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%><center>
  <H4>Procedure non ancora disponibili!</H4></center>
  
<% Else
  
  i=1 'inizializza la variabile i (contatore delle domande)
  Do until rsTabella.EOF
  		 

  
    							'	0						1				2					3						4					5					6					7					8				9
 					
   %>
  
  <table border="1"  align=center width="60%">
		<tr>
			<td><font><b><%=rsTabella(1)%></b></td>
			<td><font><b><%=rsTabella(2)%></b></font></td> 
		</tr>
		<tr>
			<td colspan=2 align="center">
			<%
					Select Case rsTabella(7)
					Case 1 
						%><img src="<%=url_img(Modulo,rsTabella.Fields("Risposta1"))%>" style="border: 1px dotted #4F6A98;"><BR><%
						Case 2 
						%><img src="<%=url_img(Modulo,rsTabella.Fields("Risposta2"))%>" style="border: 1px dotted #4F6A98;"><BR><%
						Case 3 
						%><img src="<%=url_img(Modulo,rsTabella.Fields("Risposta3"))%>" style="border: 1px dotted #4F6A98;"><BR><%
						Case 4 
						%><img src="<%=url_img(Modulo,rsTabella.Fields("Risposta4"))%>" style="border: 1px dotted #4F6A98;"><BR><%	
					End Select
					%> 
		     </td>
		</tr>
		 <% if rsTabella(10)="si" then %>
		     	<% if rsTabella(12)=1 then ' devo definire un url diverso per visualizzare i video lunghi creati con camtasia come adesso%>
				 
			 	     <tr><td align="center" colspan="2">  <a href="<%=url_video1(Modulo,rsTabella.Fields("NumeroDomanda"))%>.html" target="_blank"> video </a>               </td></tr>	
				<%else%>
					 
		      	<tr><td align="center" colspan="2">  <a href="spiegazione_video.asp?davisualizzazioni=<%=davisualizzazioni%>&video=<%=url_video(Modulo,rsTabella.Fields("NumeroDomanda"))%>&ID_Mod=<%=rsTabella(13)%>&ID_Paragrafo=<%=rsTabella(14)%>&CodiceDomanda=<%=rsTabella(0)%> " target="_blanck"> video </a>               </td></tr>	
		      	<%end if%>
			<%end if%>
			 <% if rsTabella(11)="si" then %>
		 <tr><td align="center" colspan="2">  <a href="<%=url_tutorial(Modulo,rsTabella.Fields("NumeroDomanda"))%>" target="_self"> spiegazione </a></td></tr>
			<%end if%>
	</table>
	<br>
<%
       i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
   <div class="citazioni">
  <a href="../cClasse/scegli_azione_app.asp?Cartella=<%=Cartella%>&Stato=1&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>">	Indietro </a>
  </div>


  </div>
  </div>
</BODY>
</HTML>