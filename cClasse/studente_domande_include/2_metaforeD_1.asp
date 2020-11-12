<%
' prelevo l'elenco delle metafore Topolino dello studente

'seleziono la metafora topolino
 
QuerySQL="SELECT * " &_
" FROM Elenco_Metafore_Desideri " &_
" WHERE CodiceAllievo='" & cod & "' and Pi=0 " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 " order by CodiceMetafora asc;"
'	response.write("<br>"&QuerySQL)
 Set rsTabellaMetafore = ConnessioneDB.Execute(QuerySQL)
 
    
 	 		
  
	 
	   'per riepilogare tutti le metafore desideri e relativi punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Desideri.Id_Stud='"& cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 ";"
	 'conta il numero di metafore navigazione
' response.write(QuerySQL1&"<br>")
	 Set rsTabella5 = ConnessioneDB.Execute(QuerySQL1)
	  numrsTabellaD=rsTabella5(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella5(0)&"" =""  then
	   numrsTabellaD=0
	 end if  
	 
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Desideri.Id_Stud='"& cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 ";"
	  'conta il numero di punti ottenuti nelle metafore navigazione
	' response.write(QuerySQL2&"<br>")
	 Set rsTabella6 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabellaPD=rsTabella6(0)
	 ' se non restituisce nulla serve per dargli un valore
	if rsTabella6(0)&"" =""  then
	   numrsTabellaPD=0
	 end if  
 
	 
		%> 