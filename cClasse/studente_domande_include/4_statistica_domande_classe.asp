<%  
 'conto le Domande svolte dallo studente nel periodo e nel modulo	
   QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_ 
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"	 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
     numrsDomande=rsTabella1(0)
	 'response.write("<br>numrsDomande="&numrsDomande)
	' conto le prefrasi totali serve per il calcolo della % svolta
	 QuerySQL1="SELECT Count(*) AS Num FROM Predomande where Id_Mod='"& rsTabellaModuli("ID_Mod")&"';"	 
	 Set rsTabella1_1 = ConnessioneDB.Execute(QuerySQL1)
	 numrsPreDomande=rsTabella1_1(0)
	' response.write(QuerySQL1)
	 if  numrsPreDomande=0 then
	    numrsPreDomande=numrsDomande
	 end if
	 ' calcolo la somma dei punteggi serve per statistica e punti bonus
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where  Moduli.ID_Mod='" & rsTabellaModuli("ID_Mod") & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_ 
	 " AND Moduli.ID_Mod='" &rsTabellaModuli("ID_Mod") &"';"
	 
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsDomande2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	    numrsDomande2=0
	 end if 
 
%>
 
 



 