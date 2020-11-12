

<%



	'sConnString = "Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
	 '" Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"

set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
dim oRs,oRs1, oCmd,oCmd1, sSQL, sAns

'Parametri per la costruzione delle tabelle T_Punteggi_....
'id_classe="1COM"
'DataCla="09/09/2013"
'DataCla2="15/06/2015"
'datafine="12/12/2112" ' devo lasciarla  altrimenti problemi perchè ha valore vuoto, perche ?
'DataClax= 1+cdate(DataCla2)
'DataCla2=DataClax
'response.write(DataCla & " - " &DataClax &" - " & datafine)

	dim oParam
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
	'response.write DataCla2
	'oParam1.value =  DataClax

'	oParam1.value = cdate(DataCla2)
	'oParam1.value ="15/10/2014"
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
	if cint(PS)=0 then ' se devo escludere i punti social
		oCmd1.CommandText = "sp_mount_CLASSIFICA_SINT"
	else
		oCmd1.CommandText = "sp_mount_CLASSIFICA"
	end if
	oCmd1.CommandType = adCmdStoredProc
	set rsTabella = oCmd1.execute







   %>

</body>
</html>
