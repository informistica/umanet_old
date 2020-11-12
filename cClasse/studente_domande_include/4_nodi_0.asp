<% '                      
'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1" &_
'" WHERE  CodiceAllievo='" & cod & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
'	  " AND ID_Mod='"&rsTabellaModuli("ID_Mod")&"'" &_ 
'	 " order by CodiceNodo asc;"
' 
 QuerySQL="SELECT * FROM [MODULO_PARAGRAFO_NODI1]" &_
" WHERE (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_ 
	  " AND ID_Mod='"&rsTabellaModuli("ID_Mod")&"'" &_ 
	 " order by CodiceNodo asc;"
 
    
 
 
   'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  ' response.write(QuerySQL)
 Set rsTabellaNodi = ConnessioneDB.Execute(QuerySQL)


	' ' per riepilogare tutti i nodi e punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM [MODULO_PARAGRAFO_NODI1]" &_
	 " where   ID_Mod<>'6C' " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_ 
	  " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo") & "';"
	     'response.write(QuerySQL)
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  numrsNodi= rsTabella1(0)
'response.write(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM MODULO_PARAGRAFO_NODI1" &_
	 " where  ID_Mod<>'6C' " &_
     " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_ 
	 " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo") & "';"
'	 
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   'numrsTabella2=rsTabella2(0)
	   numrsNodi2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsNodi2=0
	 end if 
 
	
	 
	 
		%>
   
   
    