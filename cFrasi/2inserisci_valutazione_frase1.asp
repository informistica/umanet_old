<%@ Language=VBScript %>
<html>
<head>
	<link rel="stylesheet" type="text/css" href="../stile.css">

<script language="javascript" type="text/javascript">
function showText2() {window.alert("La sessione � scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>

</head>

<%
  Response.Buffer = true
  'On Error Resume Next
    ' per il controllo della validit� della sessione, se � scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">



   <%
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,url4,RecSet

     URLprov = Request.ServerVariables("HTTP_REFERER")


   CodiceTest = Request.QueryString("CodiceTest")
   Modulo=Request.QueryString("Modulo")
   Cartella=Request.QueryString("Cartella")
   CodiceAllievo=Request.QueryString("CodiceAllievo")
   Motivazione=Request.Form("txtSegnalazione")
    umanet=Request.QueryString("umanet")
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")

  %>
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  <%

	QuerySQL1="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1)
	Valutato=rsTabella.fields("Valutato")
	DVAbilitato=rsTabella.fields("DVAbilitato")
	'limiteBonus=rsTabella("limiteBonus")
	limiteBonus=1
	rsTabella.close

		   tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")

 tFra=request.querystring("tFra")



   Paragrafo=Request.QueryString("Paragrafo")
   cla=Request.QueryString("cla")
   'Domanda = Request.Form("txtDomanda")
   'Domanda = Request.QueryString("Chi")
   Domanda = server.htmlencode(Request.QueryString("Chi"))
  ' response.write("Domanda:" &Domanda)

   'ID=Request.Form("txtCodiceDomanda")
	ID=Request.QueryString("CodiceFrase")
	'response.write("<br>Codice Domanda:" &ID)

	   Spiegazione=Request.Form("txtSpiegazione")
	   'response.write("<br>Spiegazione:" &Spiegazione)

		'VAL2=Request.Querystring("VAL")
	   RE = clng(Request.Form("txtRE"))
	  ' VAL=clng(Request.Querystring("VAL"))
	   VAL=clng(Request.Form("txtVAL"))
	   INQUIZ=clng(Request.Form("txtINQUIZ"))
	   DATA=Request.Form("txtDataDomanda")

		Segnalata=Request.Form("txtSegnalata")
		 
   		Segno=Request.Form("txtSegno")
		Motivazione=Request.Form("txtSegnalazione")
		 'response.write("<br>1="&Segnalata)
	   if Segnalata="" then
   		  Segnalata=0
   		elseif Segnalata=1 and Segno=1 Then
        	  Segnalata=2
            elseif Segnalata=1 and Segno=0 Then
                 Segnalata=1
            else
              Segnalata=0
   		end if


	   'response.write("<br>2="&Segnalata)
voto=VAL
if VAL=0 and strcomp(segnalata,"0")=0 then  ' la analizzo solo se ha voto = 0 e non � segnalata altrimenti metto val=0
	    if Valutato=1 then
	     voto=1 ' aggiunto per evitare errore

				 if len(trim(Spiegazione))<40 then

						voto=0
						Segnalata=1
						Sintesi2="TROPPO CORTA!"
					 else
					    voto=1
						Segnalata=0
					 end if


	   end if
  VAL=voto
'else

end if

	' if Segnalata = 1 then
		' voto = 0
		' VAL = 0 ' imposto automaticamente il voto a 0 se la domanda � segnalata

	' end if

   'response.write("VAL2="&VAL2)
   ' per la spiegazione della domanda
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"& Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
   url=Replace(url,"\","/")

    url_feedback=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"& Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&"_feedback.txt"
     url_feedback=Replace(url_feedback,"\","/")


		'tolgo l'aggiornameto della data anche da parte dell'amministratore: non funziona!!

    if session("Admin")=true then
      'QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "', Voto = " & VAL & ", In_Quiz = " & INQUIZ &" ,Data= '" & cdate(DATA) &"', Segnalata='" &Segnalata&"' WHERE CodiceFrase =" &ID&";"
	  QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "', Voto = " & VAL & ", In_Quiz = " & INQUIZ &" ,Segnalata='" &Segnalata&"' WHERE CodiceFrase =" &ID&";"
	 else

	'
	' tolgo l'aggiornamento della data da parte dello studente
	' modifica di una frase segnalata: imposto il voto a 1 e segnalata a 0
	' QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "', Data= '" & cdate(DATASTUD) &"', Segnalata='" &Segnalata&"' WHERE CodiceFrase =" &ID&";"
	    if clng(Request.Form("txtVAL"))>0 then ' se il voto non � zero lo lascio cos� com'�, per segnalazioni di miglioramanento
		 QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "',Segnalata='" &Segnalata&"' WHERE CodiceFrase =" &ID&";"
		else
		QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "',Segnalata='" &Segnalata&"' WHERE CodiceFrase =" &ID&";"
     ' commento perchè quando modificano si annulla il punteggio
	   '' QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "',Segnalata='" &Segnalata&"',Voto ='"&voto&"' WHERE CodiceFrase =" &ID&";"
		end if

	end if

	' response.write(QuerySQL & "segnalata="&segnalata)
	' response.write(QuerySQL)
	 ConnessioneDB.Execute(QuerySQL)
' response.write(url_feedback &"<br>")
	 if (Segnalata=1) or (Segnalata=2) then
	 ' faccio la notifica

	  if motivazione<>"" then

		set objFSO=Server.CreateObject("Scripting.FileSystemObject")
		if objFSO.FileExists(url_feedback) then
			objFSO.DeleteFile url_feedback
		end if
       response.write(url_feedback)
		Set objCreatedFile = objFSO.CreateTextFile(url_feedback, True)
		
		'  response.Write("<br><br>Creo file feedback : " &url_feedback)
		' response.Write("<br>Contenuto feedback : "& Motivazione)
		objCreatedFile.WriteLine(Motivazione)
		objCreatedFile.Close
		 if objFSO.FileExists(url_feedback) then
				'response.write("<br>Creato file di feedback:"&url_feedback)
		 else
				response.write("<br>Impossibile creare file :"&url_feedback)
				response.write ("<br>"&Err.Description)

		 end if

		 set objFSO = Nothing
		 set objCreatedFile = Nothing


	end if




	 Capitolo=Request.QueryString("Capitolo")
	 'response.write(Capitolo)




	'parametriurlnotifica = "Cartella="&Cartella&"&id_classe="&session("id_classe")&"&cod="&CodiceAllievo&"&CodiceTest="&CodiceTest&"&CodiceFrase="&ID&"&Paragrafo="&Paragrafo&"&Capitolo="&Capitolo&"&MO="&Modulo&"&VAL="&VAL&"&tCap="&tCap&"&tSot="&tSot&"&tFra="&tFra
	 'response.write(parametriurlnotifica)

	 ' aggiunto ../cFrasi/ -> la notifica viene aperta in una pagina della cartella cMessaggi
	'  Azione="<a  target=blank href=../cFrasi/2inserisci_valutazione_frase.asp?"&parametriurlnotifica&">Ho segnalato una tua frase !</a>"
	 'Testo=Motivazione
	 'Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

	  Testo=Motivazione
	 Azione="<a  target=blank href=../cFrasi/2inserisci_valutazione_frase.asp?cla="&Session("Id_Classe")&"&cod="&CodiceAllievo&"&CodiceFrase="&ID&"&cartella="&cartella&"&classe="&Session("Id_Classe")&"&CodiceTest="&CodiceTest&"&Paragrafo="&Paragrafo&"&MO="&Modulo&"&Capitolo="&Capitolo&">Ho segnalato una tua frase !</a>"
	 Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

Testo = Replace(Testo, Chr(34), "`")
Testo = Replace(Testo, "'", "`")

	 QuerySQL="select Azione from Avvisi where Azione like '%CodiceFrase="  & ID &"%';"
	 response.write(QuerySQL)
	 set rsTabNotifica=ConnessioneDB.Execute(QuerySQL)
	
	 if rsTabNotifica.eof then
		QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Testo & "','" & Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
		'response.write(QuerySQL)
		if strcomp(CodiceAllievo,Session("CodiceAllievo"))<>0 then ' evito di notificare a me stesso
			ConnessioneDB.Execute(QuerySQL)
		end if
	  end if

	 end if
' %> <br> <%
'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create the FSO.
Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(Cartella)
'response.write(url)
objFSO.DeleteFile url
Set objCreatedFile = objFSO.CreateTextFile(url, True)
' Write a line with a newline character.
objCreatedFile.WriteLine(Spiegazione)
'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
objCreatedFile.Close


' modifico immagini (se necessario!!)
img = Request("Img")

 if img = 1 then

 url1=Request.Form("txtImg1")
 url2=Request.Form("txtImg2")
 url3=Request.Form("txtImg3")

'response.write(url1&"<br>")
'response.write(url2&"<br>")
'response.write(url3&"<br>")
 QuerySQL = "DELETE FROM Frasi_Img WHERE Id_Frase = '"&ID&"';"
 ConnessioneDB.Execute(QuerySQL)

 'dopo aver eliminato le immagini presenti in precedenza inserisco quelle nuove

	if url1<>"" then
	imgname="Img1"
	 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url1 & "','" & imgname & "';"
	 ConnessioneDB.Execute(QuerySQL)
''	 response.write(QuerySQL&"<br>")
	end if

	if url2<>"" then
	imgname="Img2"
	 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url2 & "','" & imgname & "';"
	 ConnessioneDB.Execute(QuerySQL)
	'' response.write(QuerySQL&"<br>")
	end if

	if url3<>"" then
	imgname="Img3"
	 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url3 & "','" & imgname & "';"
	 ConnessioneDB.Execute(QuerySQL)
	 'response.write(QuerySQL&"<br>")
	end if

 end if
  'response.write(QuerySQL&"<br>")

On Error Resume Next
If Err.Number = 0 Then

'Session("Modificata")=true -> tolto perch� inutile: faccio comparire alert e torno alla pagina da cui sono arrivato


	urlRedirect = "../cClasse/quaderno.asp?umanet="&umanet&"&stile="&session("stile")&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&DataClaq2="&Session("DataClaq2")&"&DataClaq="& Session("DataClaq")&"&tCap="&tCap&"&tSot="& tSot&"&tFra="& tFra
	
	response.write("<script>alert('Modifica effettuata correttamente'); window.location.href = '"&urlRedirect&"'</script>")

Else
Response.Write Err.Description
Err.Number = 0
End If



urlRedirect=""





   %>
	</font>


   <!--   <h4><a href="studente_domande.asp?cod=<%=CodiceAllievo%>&CodiceAllievo=<%=CodiceAllievo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua
		a valutare o modificare le domande...</a></h4>-->
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
<%



			   %><h3><a href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Torna alla pagina Studenti </a></h3>

                <!--#include file="../include/tornaquaderno.html" -->


			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->


	</body>
	</html>
