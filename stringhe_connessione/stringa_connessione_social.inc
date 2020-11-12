<%
	'REWRITE THIS:  DON'T PUT THE DATABASE IN THE
	'SAME DIRECTORY AS THE ASP FILES.  THIS WAS JUST
	'DONE FOR DEMO PURPOSES, SO YOU CAN GET THE DEMO
	'UP AN RUNNING RIGHT AWAY.  YOU CAN
    	'HARD CODE YOUR DATABASE PATH, RATHER THAN
	'RESOLVING IT DYNAMICCLY
	'sFile = request.ServerVariables("PATH_TRANSLATED")
	'sSplit = split(sFile, "\")
	'for iCtr = 0 to uBound(sSplit) - 1
'		sDir = sDir & sSplit(ictr) & "\"
'	next
	
'	Session.Timeout=60
 if Session("DBCopiatestonline")=""  and by_email="" then
	  response.redirect "../service/redirect.asp"
 end if
	
 
	
'	sConnString = "Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
' " Initial Catalog=Copiaditestonline; User Id=sa; Password=;"
 
		if session("DB")=1 then
  
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
else
 
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
end if



 
 

set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn



%>
