<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>

<body>

<%
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
	Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		  
	ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../database/" & Session("DBCopiatestonline"))
		
		  ConnessioneDB1.Open "Driver={SQL Server};Server=MAUROSHODE6E;" &_
				  "User ID=sa;Password=;Database=Copiaditestonline"
 
		QuerySQL="SELECT * " &_
" FROM preFrasi;"  
 
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
   
    do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO preFrasi  (Id_Mod, Id_Paragrafo,CodiceFrase,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo) SELECT '" &  rsTabella(1) & "','" & rsTabella(2) & "','" & rsTabella(3) & "','" & rsTabella(4) & "'," & rsTabella(5) & "," & rsTabella(6) & ",'" & rsTabella(7) & "'," & rsTabella(8) & "," & rsTabella(9) & ",'" & rsTabella(10) &"';"
		response.write("<br>"&QuerySQL2)
		ConnessioneDB1.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set ConnessioneDB=nothing
	   set ConnessioneDB1=nothing
	  set rsTabella=nothing
%>
</body>
</html>
