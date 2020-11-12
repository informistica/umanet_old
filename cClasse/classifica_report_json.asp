
 <!-- #include file = "../cUtenti/adovbs.inc" -->
<%
'Response.charset="utf-8"  
Response.charset="iso-8859-1"
Call Response.AddHeader("Access-Control-Allow-Origin", "*") 

 		Set ConnessioneDB0 = Server.CreateObject("ADODB.Connection")  ' per il DBClassifica
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario


		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

      <%

set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
dim oRs,oRs1, oCmd,oCmd1, sSQL, sAns
dim oParam

function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sInput,"è","&egrave;")

  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")

ReplaceCar = sAns
 

end function






function classereport(idp,DataCla,DataCla2)

	set oCmd = Server.CreateObject("ADODB.Command")
    set oCmd.ActiveConnection = conn
	
	oCmd.CommandText = "sp_CLASSIFICA"
	oCmd.CommandType = adCmdStoredProc
	set oParam = cmd.CreateParameter("data1", 202, 1,10)
	oCmd.parameters.append oParam
	oParam.value = DataCla
	set oParam1 = cmd.CreateParameter("data2", 202, 1,10)
	oCmd.parameters.append oParam1
	oParam1.value =  DataCla2

	set oParam2 = cmd.CreateParameter("datafine", 202, 1,10)
	oCmd.parameters.append oParam2
	oParam2.value = datafine
	set oParam3 = cmd.CreateParameter("id_classe", 200, 1,20)
	oCmd.parameters.append oParam3
	oParam3.value = id_classe
	oCmd.execute ' creo le tabelle d'appoggio
	set oParam = nothing
	set oParam1 = nothing
	set oParam2 = nothing
	set oParam3 = nothing
	' Eseguo la seconda procedura per montare la classifica
	set oCmd1 = Server.CreateObject("ADODB.Command")
	set oCmd1.ActiveConnection = conn
	'response.Write("id_classe"&id_classe)
	'if cint(PS)=0 then ' se devo escludere i punti social
	'	oCmd1.CommandText = "sp_mount_CLASSIFICA_SINT"
	'else
		oCmd1.CommandText = "sp_mount_CLASSIFICA"
	'end if
	oCmd1.CommandType = adCmdStoredProc
	set rsTabella = oCmd1.execute
	'response.write("<br>")
	riga="Report &"&classe&"&"&datacla&"&"&datacla2

	i=0
	rsTabella.movefirst
	max=rsTabella("TOT")

	do while not rsTabella.eof

		punti=rsTabella(7)
		CodiceAllievo=rsTabella("CodiceAllievo")

		i=i+1
		 QuerySQL="  INSERT INTO Report (Id_Periodo, Id_Stud,Id_Classe,N,VV,PT)  SELECT " & idp  & ",'" & CodiceAllievo & "','" & id_classe & "'," & i & ",'" & fix((rsTabella("TOT")*ScalaValutaz/max) * 10) / 10 &"',"&punti

		ConnessioneDB.Execute(QuerySQL)

		 rsTabella.movenext
	loop
	
	classereport=1
end function
  
 	id_classe=request.querystring("Id_classe")
 	classe=request.querystring("classe")
	'response.write("id_classe-classe:"&id_classe&"-"&classe)
 	QuerySQL="Select count(*) from Allievi where Id_Classe='" & id_classe&"' and Attivo=1;"
	Set rsTabellaT = ConnessioneDB.Execute(QuerySQL)
	numstud=rsTabellaT(0)
	 'response.write(numstud)
	rsTabellaT.close
	QuerySQL="Select * from Setting where Id_Classe='" & id_classe&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	ScalaValutaz=rsTabellaCI("ScalaValutaz")
	rsTabellaCI.close
	QuerySQL="Select count(*) from [3PERIODI] where Id_Classe='" & id_classe&"';"
	Set rsTabellaPeriodi = ConnessioneDB.Execute(QuerySQL)
	numperiodi=rsTabellaPeriodi(0)
	QuerySQL="Select * from [3PERIODI] where Id_Classe='" & id_classe&"' order by Data;"
	Set rsTabellaPeriodi = ConnessioneDB.Execute(QuerySQL)
	
	' cancello tutti i dati precedenti 
	QuerySQL="delete from Report where Id_Classe='" & id_classe&"';"
	ConnessioneDB.Execute(QuerySQL)
 
	data_from=inizio_anno
	nome_file="report&"&classe&".json"
	intestazione=nome_file
	do while not rsTabellaPeriodi.eof
		id_periodo=rsTabellaPeriodi("ID_Periodo")
		data_to=left(rsTabellaPeriodi("Data"),11)
		intestazione=intestazione&"&"&replace(data_to,"/","_")
	 
		a=classereport(id_periodo,data_from,data_to)
		data_from=data_to
		rsTabellaPeriodi.movenext
		i=i+1
	loop
	intestazione=intestazione&"&"
	'ultimo periodo
	oggi = Right("0" & Day(Date()),2) &"/"& Right("0" & Month(Date),2) &"/"&Year(Date)  
 
	rsTabellaPeriodi.close()
	
	'response.write("Intestazione="&intestazione)


	
Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

 

url=Server.MapPath(homesito & "/grafici/as_")& anno_scolastico& "/recupero_"&nome_file
url=Replace(url,"\","/")
Set objCreatedFile = objFSO.CreateTextFile(url, True)
response.write("<br><br>Creato il file:"&url&"<br>") 

'objCreatedFile.WriteLine(intestazione)

'****DA SISTEMARE
numcolonne=10


riga="{""intestazione"":"""&intestazione&""", ""risultati"": ["


	'dopo aver inserito la tabella Report accedo ad essa per creare il file
	QuerySQL="Select * from [Allievi] where Id_Classe='" & id_classe&"' and Attivo=1 order by Cognome, Nome;"
	Set rsTabellaAllievi = ConnessioneDB.Execute(QuerySQL)
	 
	'for j=0 to numperiodi-1
	'	riga=riga+"<th><b>N.</b></th><th><center><b>VV</b></center></th><th><center><b>PT</b></center></th>"
	'next
	'riga=riga&"</tr></thead><tbody>"
	'response.write(riga&"<br>")
	'objCreatedFile.WriteLine(riga)

	rsTabellaAllievi.movefirst
	
	do while not rsTabellaAllievi.eof
        j=0
        cognome=replace(trim(rsTabellaAllievi("Cognome"))," ","")
        nome=replace(trim(rsTabellaAllievi("Nome"))," ","")
        media=0
		if j=0 then
		  riga=riga&"["""&cognome&nome&""","
		  j=1
		end if
		QuerySQL="select N,VV,PT from Report where Id_Stud='"&rsTabellaAllievi("CodiceAllievo")&"' order by ID"
		Set rsTabellaReportAllievo = ConnessioneDB.Execute(QuerySQL)
		do while not rsTabellaReportAllievo.eof
			N=rsTabellaReportAllievo("N")
			VV=rsTabellaReportAllievo("VV")
            media=media+round(VV,2)
            VV=replace(VV,",",".")
			PT=rsTabellaReportAllievo("PT")
			riga=riga&""""&N&""","""&VV&""","""&PT&""","
            riga=trim(riga)
			rsTabellaReportAllievo.movenext
            j=j+1
		loop
        media=media/j
        media=round(media,1)
        media=replace(media,",",".")
        riga=trim(riga&""""&media&""",")
        riga=left(riga,len(riga)-1)
		riga=riga&"],"
	'	response.write(riga)
		'objCreatedFile.WriteLine(riga)
		rsTabellaAllievi.movenext
		j=0
	loop
    riga=left(riga,len(riga)-1)
    riga=riga&"]}"
    response.write(riga)
    objCreatedFile.WriteLine(riga)

rsTabellaAllievi.close()
rsTabellaReportAllievo.close()
objCreatedFile.Close


%>

 








 