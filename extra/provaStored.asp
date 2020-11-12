<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>

<body>


<%
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
 		  
			   sConnString = "Driver={SQL Server};Server=MAUROSHODE6E;" &_
				  "User ID=sa;Password=;Database=Copiaditestonline"
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn  
set oCmd = Server.CreateObject("ADODB.Command")
	set oCmd.ActiveConnection = conn
	oCmd.CommandText = "FORUM_MESSAGE"
	oCmd.CommandType = 4
	set oParam = cmd.CreateParameter("MESSAGEID", 3, 1)
	oCmd.parameters.append oParam
	oParam.value = cint(ID)
	
'	set oParam1 = cmd.CreateParameter("Bacheca", 3, 1)
'	oCmd.parameters.append oParam1
'	oParam1.value = "informistica"
'	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
'	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_
	
	set oRs = oCmd.execute
	set oParam = nothing

%>
</body>
</html>
