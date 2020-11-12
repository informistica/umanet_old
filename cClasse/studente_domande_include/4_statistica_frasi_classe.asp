<%  
 'conto le frasi svolte dallo studente nel periodo e nel modulo	
  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_  
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"	 
	 
	' WHERE     (Data >= CONVERT(DATETIME, '2014-05-01 00:00:00', 102)) AND (Data <= CONVERT(DATETIME, '2014-05-31 00:00:00', 102))
	 
	' QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
'	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "' and Frasi.Id_Stud='"& cod & "'" &_
'	 " and  Data>=#" &formattaDataCla(DataClaq)  &"#" &_
'	 " AND Data<=#" & 1+ CDATE(formattaDataCla(DataClaq2)) &"#"&_  
'	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"	 
	 
	 ' (Data < CONVERT(DATETIME, '2014-05-31 00:00:00', 102)) AND (Data > CONVERT(DATETIME, '2014-05-01 00:00:00', 102))
	 
	 
	' response.write(QuerySQL1&"<br>")
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
     numrsFrasi=rsTabella1(0)
	' conto le prefrasi totali serve per il calcolo della % svolta
	 QuerySQL1="SELECT Count(*) AS Num FROM Prefrasi where Id_Mod='"& rsTabellaModuli("ID_Mod")&"';"	 
	 Set rsTabella1_1 = ConnessioneDB.Execute(QuerySQL1)
	 numrsPreFrasi=rsTabella1_1(0)
	' response.write(QuerySQL1)
	  
	 ' calcolo la somma dei punteggi serve per statistica e punti bonus
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_  
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"
	 
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsFrasi2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	    numrsFrasi2=0
	 end if 
 
%>
 
 



 