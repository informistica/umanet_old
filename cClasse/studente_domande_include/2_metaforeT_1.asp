<%
' prelevo l'elenco delle metafore Topolino dello studente

'seleziono la metafora topolino
QuerySQL="SELECT * " &_
" FROM Elenco_Metafore_Topolino " &_
" WHERE CodiceAllievo='" & cod & "' and Pi=0 " &_
	  " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 " order by CodiceMetafora asc;"
 
 'response.write(QuerySQL)
 Set rsTabellaMetafore = ConnessioneDB.Execute(QuerySQL)
 	
 
 
' per riepilogare tutte le metafore topolino, navigazione,..., e  relativi punti da indicare nel titolo Interfaccia UWWW N(..)
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Topolino.Id_Stud='"& cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 ";"
	 'conta il numero di metafore topolino 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  numrsTabellaT=rsTabella1(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella1(0)&"" =""  then
	   numrsTabellaT=0
	 end if 
	 
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Topolino.Id_Stud='"& cod & "'" &_
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
	  %>