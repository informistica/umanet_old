<%@ Language=VBScript %>

  <% 
   Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
 
	idConta=request.querystring("ID")
	iThreadParent=request.querystring("TParent")
 
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
%>
   	<!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->
	<!--#include file = "../service/controllo_sessione.asp"-->

<%

 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	'response.write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	VotoPalese=rsTabella("VotoPalese")
	 

  QuerySQL="select count(*) from Voti WHERE ThreadQuote="& idConta &";"


		 set oRs1=ConnessioneDB.Execute(QuerySQL)
		 if oRs1.eof then
	       voti=0
		   else
		   voti=oRs1(0)
		 end if
'response.write(QuerySQL)

	  	   QuerySQL=" SELECT Sum(Voto) AS SommaDiVoto, CodiceAllievo,Cognome,Nome from Voti WHERE ThreadQuote="& idConta &" GROUP BY CodiceAllievo,Cognome,Nome ;"

		 set oRs1=ConnessioneDB.Execute(QuerySQL)
		 if oRs1.eof then
	       voti=0
		   else
		   voti=oRs1(0)
		 end if
 
	mipiace=0
	nonpiace=0
	numVotanti=0
	voto=0
	titolo=""
	titolo1="Piace a "
	titolo2="Non piace a "

	while not  oRs1.eof
	   if oRs1("SommaDiVoto")>0 then
			 if (VotoPalese=1) or (Session("Admin")=true) then
			     titolo=titolo&" "&titolo1&" "&oRs1("Cognome") &" " & left(oRs1("Nome"),1) &"." & "(Voto=" & 5 + oRs1("SommaDiVoto")&")"
			 else
				 'sAns1=sAns1&"<img src='img/icon_star_red.gif' width='13' height='12' ><br>"
			 end if
	      mipiace=mipiace+oRs1("SommaDiVoto")
	   else
	         if (VotoPalese=1) or (Session("Admin")=true) then
			    titolo=titolo&" "&titolo2&" "&oRs1("Cognome") &" " & left(oRs1("Nome"),1) &"." & "(Voto=" & 6 + oRs1("SommaDiVoto")&")"

	         else
			 			    ' sAns1=sAns1&"<img src='img/icon_star_black.gif' width='13' height='12' ><br>"
			 end if
	     nonpiace=nonpiace+oRs1("SommaDiVoto")
	   end if
'next
	if oRs1("SommaDiVoto")>0 then
		voto=voto + (5 + oRs1("SommaDiVoto"))
	else
	   voto=voto + (6 + oRs1("SommaDiVoto"))
	end if
	numVotanti=numVotanti+1
	oRs1.movenext
	wend
    'titolo="ciao"
    media= fix ((voto/numVotanti)*10)/10 
    response.write(media&"£"&titolo)
	set orS1=nothing
 %>