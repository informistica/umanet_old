<%@ Language=VBScript %>
<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>
    <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  	<!-- #include file = "../service/controllo_sessione.asp" -->
    <%






	id_prefrase=request.querystring("ID_Prefrase")
	Quesito=request.querystring("Quesito")
	Spiegazione=request.querystring("chi")
	CodiceAllievo=Session("CodiceAllievo")
	CodiceTest=request.querystring("CodiceTest")
	Modulo=request.querystring("Modulo")
	Paragrafo=request.querystring("Paragrafo")
	Cartella=request.querystring("Cartella")
	img = request.querystring("Img")
	cFile = request.querystring("cFile")
	sintesi = request.querystring("sintesi")
	Segnalata = 0
	voto=1
	CodiceSottopar=request.querystring("CodiceSottopar")
  Sottoparagrafo=request.querystring("Sottoparagrafo")

 url1=Request.querystring("Img1")
 url2=Request.querystring("Img2")
 url3=Request.querystring("Img3")



' vedo se l'utente ha già inserito la frase cper quella prefrase,
' se è presente cancello il file spiegazione e lo aggiorno come in ins_valutaz_frase1
' altrimenti lo creo comein ins_frase1'

QuerySQL = "SELECT CodiceFrase FROM Frasi WHERE Id_Stud='"&CodiceAllievo &"' and Id_Prefrase="&id_prefrase
'response.write("<br>"&QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
if  rsTabella.eof Then  ' non esiste creo per la prima volta'


	QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Id_Prefrase,Img,Segnalata,Id_Sottoparagrafo) SELECT '" & Quesito & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & FormatDateTime(now,2) & "','" & Voto & "','" & Cartella & "','" & FormatDateTime(now, 4) & "'," & id_prefrase & "," & img & "," & Segnalata & ",'" & CodiceSottopar &"';"
	'response.write("<br>"&QuerySQL)
	ConnessioneDB.Execute QuerySQL
	QuerySQL = "SELECT CodiceFrase FROM Frasi WHERE CodiceFrase=(Select Max(CodiceFrase) FROM Frasi);"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	ID=rsTabella(0)
	CARTA=rsTabella(1)
   ' inserimento inserito il 24/05/19 ma non testato, ho sospeso il salvataggio come bozza per capire il problema degli url persi per alcune risposte
		   'qua inserisco   le immagini (o le pagine html) linkate cpon url anzichè uploadate
			if url1<>"" then
			imgname="Img1"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url1 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			' response.write(QuerySQL&"<br>")
			end if
			if url2<>"" then
			imgname="Img2"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url2 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			' response.write(QuerySQL&"<br>")
			end if
			if url3<>"" then
			imgname="Img3"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url3 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			'' response.write(QuerySQL&"<br>")
			end if

Else  ' esiste già aggiorno gli url delle immagini'
  ID=rsTabella(0)
  ' non posso aggiornare le immagini perchè la tabella frasi_img non ha la chiave primari e quindi non posso distinguere i 3 url

End if

 ' response.write(QuerySQL)

'	prelava ID dell'ultimo record inserito


	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	url=Replace(url,"\","/")
	'response.write("<br>"&url)
		'response.write("<br>chi="&chi)

'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
set objFSO=Server.CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(url) then
	objFSO.DeleteFile url
end if

Set objCreatedFile = objFSO.CreateTextFile(url, True)
objCreatedFile.WriteLine(ltrim(Spiegazione))
objCreatedFile.Close




On Error Resume Next
'response.write("eccezione="& session("eccezione")&"& rand="&rand&" proba="&probabilita&"<br>")
If Err.Number = 0 Then
  Response.Write "Salvata come bozza!"
Else
  Response.Write Err.Description
Err.Number = 0
End If
%>
