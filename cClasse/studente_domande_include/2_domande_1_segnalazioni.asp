<%
'DataClaq="10/01/2012"
'DataClaq2="01/01/2015"
  
  set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")

'conn.mode = 3

set cmd1.activeconnection = conn  
set cmd2.activeconnection = conn 
set cmd3.activeconnection = conn 

idpar=rsTabellaParagrafi("ID_Paragrafo")

	cmd1.CommandText = "sp_DOMANDE1_segnalate"
	cmd1.CommandType = adCmdStoredProc
	set oParam1 = cmd1.CreateParameter("data1", 202, 1,10)
	cmd1.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd1.CreateParameter("data2", 202, 1,10)
	cmd1.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam4 = cmd1.CreateParameter("id_paragrafo", 200, 1,20)
	cmd1.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabellaDomande=cmd1.execute 'si blocca qui
	
	
		    '   dim objFSO,objCreatedFile
			'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
			'	Dim sRead, sReadLine, sReadAll, objTextFile
			'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logDom_"&p&"_1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(idpar)
'				objCreatedFile.Close
				
	
	
	' QuerySQL="SELECT * FROM  [MODULO_PARAGRAFO_DOMANDE1] WHERE CodiceAllievo='LordGourmet';"
	'QuerySQL="SELECT * FROM  [MODULO_PARAGRAFO_DOMANDE1] WHERE CodiceAllievo='informistica';"
	'QuerySQL="SELECT * FROM  [MODULO_PARAGRAFO_DOMANDE1] WHERE CodiceAllievo='"&cod&"' and ID_Paragrafo='"& idpar &"';"
    ' Set rsTabellaDomande = ConnessioneDB.Execute(QuerySQL)

	' ho aggiunto nella vista url_teoria
    
	 
	 
	cmd2.CommandText = "sp_count_DOMANDE1_segnalate"
	cmd2.CommandType = adCmdStoredProc
	set oParam1 = cmd2.CreateParameter("data1", 202, 1,10)
	cmd2.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd2.CreateParameter("data2", 202, 1,10)
	cmd2.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam4 = cmd2.CreateParameter("id_paragrafo", 200, 1,20)
	cmd2.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabella1=cmd2.execute ' creo le tabelle d'appoggio 
	
	' QuerySQL="SELECT count(*) FROM [MODULO_PARAGRAFO_DOMANDE1] WHERE CodiceAllievo='informistica';"
     'Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

	
	
	
	numrsDomande=rsTabella1(0)
 
 
 'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logDom_"&p&"_2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(numrsDomande)
'				objCreatedFile.Close
'				
	
 
 
 	 
	cmd3.CommandText = "sp_sum_DOMANDE1_segnalate"
	cmd3.CommandType = adCmdStoredProc
	set oParam1 = cmd3.CreateParameter("data1", 202, 1,10)
	cmd3.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd3.CreateParameter("data2", 202, 1,10)
	cmd3.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam4 = cmd3.CreateParameter("id_paragrafo", 200, 1,20)
	cmd3.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabella2=cmd3.execute ' creo le tabelle d'appoggio 
	 'SELECT sum(*) FROM  MODULO_PARAGRAFO_DOMANDE1 WHERE CodiceAllievo='informistica';"
	'QuerySQL=" SELECT sum(Voto) FROM [MODULO_PARAGRAFO_DOMANDE1]  WHERE CodiceAllievo='informistica';"
     'Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)






	
	 numrsDomande2=rsTabella2(0)
	 if rsTabella2(0)&"" =""  then
	   numrsDomande2=0
	 end if 
	
	' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logDom_"&p&"_3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(numrsDomande2)
'				objCreatedFile.Close

	
	
	
  
	 set cmd1=nothing
	set cmd2=nothing
	set cmd3=nothing
  
   set oParam1 = nothing
	set oParam2 = nothing
	set oParam4 = nothing


  

 %>