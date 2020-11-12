<%@ Language=VBScript %>


<%
Response.charset="utf-8"
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
  sAns = Replace(sAns,chr(34),chr(39)) ' sostituizione " con il classico apice

 ReplaceCar = sAns
 end function
 Dim sRead, sReadLine, sReadAll
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
    <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  	<!-- #include file = "../service/controllo_sessione.asp" -->
    <%
	  Quesito = Request.QueryString("Quesito")
	  id_classe=Request.QueryString("id_classe") ' 0 Topolino, 1 Navigazione, 2 Sdesideri

    querySql="select count(*)  from Allievi where Id_Classe='"&id_classe&"' and Attivo=1"
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    num=rsTabella(0)

    querySql="select CodiceAllievo,Cognome,Nome,Classe from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 order by Cognome,Nome"
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)

    i=1
    codice="Non ha risposto"
    json="{"
    json=json&"""num"": """&num&""","
    do while not rsTabella.EOF
       querySql="select CodiceFrase,Id_Mod,Cartella,Id_Arg from Frasi where Id_Stud='"&rsTabella("CodiceAllievo") &"' and Chi='"&Quesito&"'"
       Set rsTabella1 = ConnessioneDB.Execute(querySql)
       if Paragrafo="" Then
         query="Select Titolo from Paragrafi where ID_Paragrafo='"&rsTabella1("Id_Arg")&"'"
          Set rsTabellaTitolo = ConnessioneDB.Execute(query)
          paragrafo=rsTabellaTitolo(0)
       end if
         if not rsTabella1.eof then
         url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & rsTabella1("Cartella") &"/" &rsTabella1("Id_Mod")&"_Frasi/"&rsTabella1("Id_Mod")&"_"&Paragrafo&"_"&rsTabella1("CodiceFrase")&".txt"
         url=Replace(url,"\","/")
         	if objFSO.FileExists(url) then
         		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
         		sReadAll="" 'pulisco sReadAll -> altrimenti rimane la vecchia spiegazione
         		sReadAll = objTextFile.ReadAll
         		objTextFile.Close
         	else
         	  sReadAll="Il file non esiste"
         	end if
         json=json&""""&i&""": """&ReplaceCar(sReadAll)&""","
         Else
           json=json&""""&i&""": """&codice&""","
         end if
       rsTabella.MoveNext
       i=i+1
    loop
    json=left(json,len(json)-1)
    json=json&"}"
    response.write(json)



 %>
