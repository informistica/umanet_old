<%
	'REWRITE THIS:  DON'T PUT THE DATABASE IN THE
	'SAME DIRECTORY AS THE ASP FILES.  THIS WAS JUST
	'DONE FOR DEMO PURPOSES, SO YOU CAN GET THE DEMO
	'UP AN RUNNING RIGHT AWAY.  YOU CAN
    	'HARD CODE YOUR DATABASE PATH, RATHER THAN
	'RESOLVING IT DYNAMICCLY
	sFile = request.ServerVariables("PATH_TRANSLATED")
	sSplit = split(sFile, "\")
	for iCtr = 0 to uBound(sSplit) - 1
		sDir = sDir & sSplit(ictr) & "\"
	next
	
	Session.Timeout=60
 if Session("DBDiario")="" then
	  response.redirect "../redirect.asp"
 end if
	
	'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDir & "forum.mdb"
	
	sConnString =  "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../../database/" & Session("DBDiario"))
	'sConnString =  "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
	'"DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBDiario")
	
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn



%>
