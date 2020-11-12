<%
' prelevo l'elenco delle metafore Topolino dello studente
 
 	
 ' seleziono le metafore navigazione
 QuerySQL="SELECT * " &_
" FROM Elenco_Metafore_Navigazione " &_
" WHERE CodiceAllievo='" & cod & "' and Pi=0 " &_
 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 " order by CodiceMetafora asc;"
	 
 Set rsTabellaMetafore = ConnessioneDB.Execute(QuerySQL)
 'rsTabella0.movefirst

' response.write("?<br>"&QuerySQL)
	  'per riepilogare tutti le metafore navigazione e relativi punti 
	   
	  QuerySQL2="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Navigazione.Id_Stud='"& cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 ";"
	 'conta il numero di metafore navigazione
	
	' response.write("???????"&QuerySQL2)
	 Set rsTabella3 = ConnessioneDB.Execute(QuerySQL2)
	  numrsTabellaN=rsTabella3(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella3(0)&"" =""  then
	   numrsTabellaN=0
	 end if  
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Navigazione.Id_Stud='"&cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 ";"
	  'conta il numero di punti ottenuti nelle metafore navigazione
	
	 Set rsTabella4 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabellaPN=rsTabella4(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella4(0)&"" =""  then
	   numrsTabellaPN=0
	 end if  
	 
	   %>