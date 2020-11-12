<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>

<body>

<% 

	QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and Frasi.Id_Stud='"& Session("CodiceAllievo") & "'" &_
	 " and Frasi.Id_Mod='" & Modulo &"' and Frasi.Id_Arg='" & CodiceTest &"';"	 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	 numFrasi=rsTabella1(0)
	 if rsTabella1(0)&"" =""  then
	   numFrasi=0
	 end if 
	 
%>
</body>

</html>
