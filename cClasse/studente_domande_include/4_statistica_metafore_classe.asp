<%  

 'response.write("cioa")

if strcomp(rsTabellaModuli("Titolo"),"Interfaccia UWWW")=0 then
 '  QuerySQL1="SELECT Count(*) FROM PreMetafore1 where Id_Mod='"& rsTabellaModuli("ID_Mod")&"';"	 
'  '  response.write(QuerySQL1 & " " &numrsPreMetafore)
'	 Set rsTabella1_1 = ConnessioneDB.Execute(QuerySQL1)
'	
'	 numrsPreMetafore=rsTabella1_1(0)

	' response.write(QuerySQL1 & " " &numrsPreMetafore)
	  
  
  
  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));" 
	 'conta il numero di metafore topolino 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  numrsTabellaT=rsTabella1(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella1(0)&"" =""  then
	   numrsTabellaT=0
	 end if 
 'response.write("<br>numrsTabellaT="&numrsTabellaT)
	' response.write(QuerySQL1)
	 QuerySQLTOPO=QuerySQL1
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));"
	 'conta il numero di punti ottenuti nelle metafore topolino
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	 ' numrsTabella2=rsTabella2(0)
	    numrsTabellaPT=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabellaPT=0
	 end if 
	 
	'  QuerySQL1="SELECT Count(*) AS Num FROM PreMetafore where Id_Mod='"& rsTabellaModuli("ID_Mod")&"';"	 
	 'Set rsTabella1_1 = ConnessioneDB.Execute(QuerySQL1)
	' numrsPreMetafore=rsTabella1_1(0)
	 
	 
	   QuerySQL2="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));" 
	 'conta il numero di metafore navigazione
	
	' response.write("???????"&QuerySQL2)
	 Set rsTabella3 = ConnessioneDB.Execute(QuerySQL2)
	  numrsTabellaN=rsTabella3(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella3(0)&"" =""  then
	   numrsTabellaN=0
	 end if  
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));" 
	  'conta il numero di punti ottenuti nelle metafore navigazione
	
	 Set rsTabella4 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabellaPN=rsTabella4(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella4(0)&"" =""  then
	   numrsTabellaPN=0
	 end if  
	 
	 	 
	   'per riepilogare tutti le metafore desideri e relativi punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));"
	 'conta il numero di metafore navigazione
' response.write(QuerySQL1&"<br>")
	 Set rsTabella5 = ConnessioneDB.Execute(QuerySQL1)
	  numrsTabellaD=rsTabella5(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella5(0)&"" =""  then
	   numrsTabellaD=0
	 end if  
	 
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' " &_
 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));" 
	  'conta il numero di punti ottenuti nelle metafore navigazione
	' response.write(QuerySQL2&"<br>")
	 Set rsTabella6 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabellaPD=rsTabella6(0)
	 ' se non restituisce nulla serve per dargli un valore
	if rsTabella6(0)&"" =""  then
	   numrsTabellaPD=0
	 end if  
	 
 
	 
 numrsMetafore=numrsTabellaN+numrsTabellaT+numrsTabellaD
 ' per ora lo metto qui anzichè in cima al file perchè per  calcolarlo   devo fare la stored procedure con la Union premetafore1 e non so se servirà
 numrsPreMetafore=numrsMetafore
 
 numrsMetafore2=numrsTabellaPN+numrsTabellaPT+numrsTabellaPD
 else
 numrsMetafore=0
 numrsMetafore2=0
 numrsPreMetafore=0
 end if
 
%>
 
 



 