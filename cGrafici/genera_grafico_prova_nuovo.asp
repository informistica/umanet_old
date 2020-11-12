<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 'on error resume next
PS=request.querystring("PS")
byGrafico=Request.QueryString("byGrafico")
'on error resume next

id_classe=request.querystring("id_classe")
DataCla=request.QueryString("DataCla")
DataCla2=request.QueryString("DataCla2")
indice_periodo= Request.QueryString("indice_periodo")
divid=Session("divid")
' data del nuovo periodo quando sono chiamato
Al=Request.QueryString("Al")

  Dim ConnessioneDB,ConnessioneDB1, rsTabella,rsTabella1, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione
  Dim periodi(), indice_periodo

   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
	 Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection")
	Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario


	 'Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection")

	%>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" --><!-- Da implementare -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->
    <!-- #include file = "../var_globali.inc" -->




   </body>
<center>
<p>

</p>
<div class="citazioni" ><div> <span style="font-style: normal">

<b><font size="3">GESTIONE</font>&nbsp;</b> </span></div>
<hr>
<%
 dim objFSO,objCreatedFile
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim sRead, sReadLine, sReadAll, objTextFile



   QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato=rsTabellaCI("CIAbilitato")
	ScalaValutaz=rsTabellaCI("ScalaValutaz")
	Runner=rsTabellaCI("Runner")
	rsTabellaCI.close



 QuerySQL="SELECT count(*) FROM [3PERIODI] Where ID_Classe='"& Session("Id_Classe") &"';"
 response.write(QuerySQL&"<br>")
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
numPeriodi=rsTabella(0)+1' +1 per Oggi
response.write("numPeriodi="&numPeriodi&"<br>")
redim periodi(numPeriodi)
' faccio la query per prelevare i periodi di valutazione per questa classe
QuerySQL="SELECT * FROM [3PERIODI] Where Id_Classe='"& Session("Id_Classe") &"';"
response.write(QuerySQL&"<br>")
		'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logPeriodi.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   ' carico il vettore delle date di valutazione
response.write("riga 85")
    periodi(0)=inizio_anno
response.write("riga 87")

' 1_classifica_new necessita dei valori per datacla,datacla2,datafine.id_classe'
	%>
         <!-- #include file = "../cClasse/studente_domande_include/1_classifica_new.asp" -->

    <%

'CodiceAllievo=id
response.write("riga 95")
if byGrafico<>"" then
response.write("riga 97")
		Dim FileObject,riga
		Set FileObject=CreateObject("Scripting.FileSystemObject")
	'	esiste=FileObject.FileExists("data.xml")
		'Response.Write("<p>il file esiste? "&esiste&"</p>")
		Set FileObject=Nothing



		Set objFSO = CreateObject("Scripting.FileSystemObject")

		'url=Server.MapPath(homesito & "/Grafici")& "/data.xml"

		'url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("Cartella")& "/data.xml"

		url=Server.MapPath(homesito & "/script/cGrafici/Grafici")& "/data.xml"
		url=Replace(url,"\","/")
		Set objCreatedFile = objFSO.CreateTextFile(url, True)
		' Write a line with a newline character.
		gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())

		riga="<chart caption='PUNTEGGI DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Punti' numberPrefix='P.'>"
		objCreatedFile.WriteLine(riga)
		 do while not rsTabella.EOF
			cognome=rsTabella.fields("Cognome")
			cognome1=Replace(cognome," ","")
			'Replace(url,"\","/")

			riga="<set label='" &cognome1 &"' value='" &rsTabella.fields("TOT") & "'/>"
			objCreatedFile.WriteLine(riga)
			rsTabella.movenext
		   loop

		riga="</chart>"
		objCreatedFile.WriteLine(riga)

		' objCreatedFile.WriteLine(rsTabella(0))

		'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
		objCreatedFile.Close
		rsTabella.close()

		Response.redirect "Grafici/genera.html"
		'If Err.Number = 0 Then
'
'		Response.redirect "../Grafici/genera.html"
'		Else
'		Response.Write Err.Description
'		Err.Number = 0
'		End If

else
response.write("riga 147")
' genero la cronologia della classifica per la classe
     ' ottengo dal al
	  QuerySQL="SELECT MAX(ID_Periodo) AS [MAX] FROM [3PERIODI] WHERE Id_Classe='" & id_classe & "'"
    response.write(QuerySQL&"<br>")
	  Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
	  maxP=rsTabella1(0)

	  QuerySQL="SELECT Data FROM [3PERIODI] WHERE ID_Periodo=" & maxP & ";"
    response.write(QuerySQL&"<br>")
	  Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
	  Al=rsTabella1("Data")
	  Dal=DataCla
	  rsTabella1.close
	  PuntiForum=0
	 Posizione=1
     do while not rsTabella.EOF 'and Posizione<20
	   VotoVirtuale=(rsTabella("TOT")/max)*ScalaValutaz
	   VotoVirtuale=round(VotoVirtuale*10)/10
	   QuerySQL="insert into [4PERIODI_CLASSIFICA] (CodiceAllievo,Dal,Al,Posizione,Tot,Pd,Pn,Pf,Pm,Pc,Ps,Vv)" &_
	   " select '" & rsTabella("CodiceAllievo") &"','" & Dal &"','" & Al &"','" & Posizione&"','" & rsTabella("TOT")&"','" & rsTabella("PD")&"','" & rsTabella("PN")&"','" & rsTabella("PF")&"','" & rsTabella("PM")&"','" & rsTabella("Crediti")&"','" & rsTabella("PuntiForum") &"','" & VotoVirtuale & "';"
     response.write(QuerySQL&"<br>")
	  ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'url="C:\Inetpub\umanetroot\Anno_2012-2013\logDataCla"&Posizione&".txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close

			ConnessioneDB.Execute QuerySQL
			rsTabella.movenext
	        Posizione= Posizione+1
	 loop
	 rsTabella.close()
     response.write("Inserita")
''	 response.Redirect "../cAdmin/admin.asp?id_classe="&id_classe&"&divid="&divid


end if

'On Error Resume Next


%>
<h4 style="text-align: center"><i><a href="../../home.asp" >Vai all'HomePage</a> </h4>
</center>
</html>
