<%@ Language=VBScript %>


        <%


 function ReplaceCar(sInput)
 dim sAns
 sAns=sInput
   sAns = Replace(sAns, "  ", " ") 'sostituizione doppio spazio con uno singolo
   sAns = Replace(sAns, "	", " ") 'sostituzione spazi per evitare errori
   sAns = Replace(sAns, " ?", "?") ' rimozione spazio prima del punto di domanda
   sAns = Replace(sAns, "’", "'") ' sostituzione di un'apice con quello classico
   sAns = Replace(sAns, "…", "...") 'sostituzione tre puntini
   sAns = Replace(sAns, Chr(25), "'") 'sostituizione apice
   sAns = Replace(sAns, VBCrLf, "") 'sostituizione ritorno a capo
  sAns = Replace(sAns,chr(96),chr(39)) ' sostituizione finale dell'apice storto con il classico apice

 ReplaceCar = sAns
 end function




		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->






    <%
	  id = Request.QueryString("id")
	 		 que="Select topic,comments  from FORUM_MESSAGES where ID="&id
		  set rsTesto=ConnessioneDB.Execute(que)
		  if not rsTesto.eof then
		  titolo=rsTesto(0)
		  testo=rsTesto(1)
		   'testo = Replace(testo, VBCrLf, "<br>")
		   testo = Replace(testo, VBCrLf, "")
       testo = Replace(testo,"""", "")
      '' testo = Replace(testo,"'", chr(96))
		   'testo = Replace(testo, """"", "'")
		  else
		   testo="????"
		  end if
		  set rsTesto=nothing
	   'response.write(testo)
'

	'	response.write("{")
	'	 response.write("""topic"": """&rtrim(ltrim(titolo))&""","  &_
	'	 """comments"": """&rtrim(ltrim(testo))&"""")
	'	 response.write("}")
	response.write(titolo&"£££"&rtrim(ltrim(testo)))
'


%>
