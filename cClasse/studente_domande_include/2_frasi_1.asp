<% 

idpar=rsTabellaParagrafi("ID_Paragrafo")
'response.write(cdate(DataClaq)&"----"&cdate(DataClaq2))
 
'DataClaq= trim(DataClaq)
'DataClaq2=trim(1+cdate(DataClaq2))
  'response.write("<br>"&DataClaq)  
'response.write("<br>"&DataClaq2) 
'cod="informistica"
'idpar="Expo_1_1"
	'DataClaq= mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  
	 'DataClaq2= mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)		
'response.write("<br>"&DataClaq)  
'response.write(DataClaq2)
'response.write("<br>"&idpar)
 		
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")

'conn.mode = 3

set cmd1.activeconnection = conn  
set cmd2.activeconnection = conn 
set cmd3.activeconnection = conn 

	cmd1.CommandText = "sp_FRASI1"
	cmd1.CommandType = adCmdStoredProc
	set oParam1 = cmd1.CreateParameter("data1", 202, 1,10)
	cmd1.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd1.CreateParameter("data2", 202, 1,10)
	cmd1.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam3 = cmd1.CreateParameter("codiceallievo", 200, 1,20)
	cmd1.parameters.append oParam3
	oParam3.value = cod
	set oParam4 = cmd1.CreateParameter("id_paragrafo", 200, 1,20)
	cmd1.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabellaFrasi=cmd1.execute ' creo le tabelle d'appoggio
	
	
    
	 
	 
	cmd2.CommandText = "sp_count_FRASI1"
	cmd2.CommandType = adCmdStoredProc
	set oParam1 = cmd2.CreateParameter("data1", 202, 1,10)
	cmd2.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd2.CreateParameter("data2", 202, 1,10)
	cmd2.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam3 = cmd2.CreateParameter("codiceallievo", 200, 1,20)
	cmd2.parameters.append oParam3
	oParam3.value = cod
	set oParam4 = cmd2.CreateParameter("id_paragrafo", 200, 1,20)
	cmd2.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabella1=cmd2.execute ' creo le tabelle d'appoggio 
	numrsFrasi=rsTabella1(0)
 
	 
	cmd3.CommandText = "sp_sum_FRASI1"
	cmd3.CommandType = adCmdStoredProc
	set oParam1 = cmd3.CreateParameter("data1", 202, 1,10)
	cmd3.parameters.append oParam1
	oParam1.value = DataClaq	
	set oParam2 = cmd3.CreateParameter("data2", 202, 1,10)
	cmd3.parameters.append oParam2
	oParam2.value = DataClaq2
	set oParam3 = cmd3.CreateParameter("codiceallievo", 200, 1,20)
	cmd3.parameters.append oParam3
	
	oParam3.value = cod
	set oParam4 = cmd3.CreateParameter("id_paragrafo", 200, 1,20)
	cmd3.parameters.append oParam4
	oParam4.value = idpar
	Set rsTabella2=cmd3.execute ' creo le tabelle d'appoggio 
	 numrsFrasi2=rsTabella2(0)
	 if rsTabella2(0)&"" =""  then
	   numrsFrasi2=0
	 end if 
	
	
    set cmd1=nothing
	set cmd2=nothing
	set cmd3=nothing
  
    set oParam1 = nothing
	set oParam2 = nothing
	set oParam3 = nothing
	set oParam4 = nothing
	
	
	
   %>