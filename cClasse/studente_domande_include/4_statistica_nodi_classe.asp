<%  
 'conto le Nodi svolte dallo studente nel periodo e nel modulo	
   QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_ 
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"	 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
     numrsNodi=rsTabella1(0)
	'response.write("numrsNodi="&numrsNodi)
	' conto le prefrasi totali serve per il calcolo della % svolta
	 QuerySQL1="SELECT Count(*) AS Num FROM Prenodi where Id_Mod='"& rsTabellaModuli("ID_Mod")&"';"	 
	 Set rsTabella1_1 = ConnessioneDB.Execute(QuerySQL1)
	 numrsPreNodi=rsTabella1_1(0)
	'response.write(QuerySQL1)
	  
	 ' calcolo la somma dei punteggi serve per statistica e punti bonus
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_ 
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"
	' response.write(QuerySQL2)
	 ' response.write("<br>" & 1+ CDATE(formattaDataCla(DataClaq2)))
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsNodi2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	    numrsNodi2=0
	 end if 
 
%>
 
 



 