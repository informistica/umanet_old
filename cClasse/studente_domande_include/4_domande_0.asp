<% 
 
QuerySQL="SELECT * FROM  [MODULO_PARAGRAFO_DOMANDE1] " &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	  " AND ID_Mod='"&rsTabellaModuli("ID_Mod")&"';" 

 'QuerySQLDomande=QuerySQL
' Set rsTabella3 = ConnessioneDB.Execute(QuerySQL)
' response.write("SASA"&QuerySQL)
 'response.write(DataClaq2)
' QuerySQL="SELECT * FROM  [MODULO_PARAGRAFO_DOMANDE1] WHERE CodiceAllievo='informistica';"
	  
 
 Set rsTabellaDomande = ConnessioneDB.Execute(QuerySQL)


 %>
 
 <% 'QuerySQL1="SELECT Count(*) AS Num FROM MODULO_PARAGRAFO_DOMANDE1" &_
'	 " where  CodiceAllievo='"& cod & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
'	 " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo") & "';" 
'	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
'	 numrsDomande= rsTabella1(0)
'	 ' response.write(QuerySQL1)
'	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM MODULO_PARAGRAFO_DOMANDE1" &_
'		 " where  CodiceAllievo='"& cod & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
'	 " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo") & "';" 
'	 
'	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
'	   'numrsTabella2=rsTabella2(0)
'	   numrsDomande2=rsTabella2(0)
'	 ' se non restituisce nulla serve per dargli un valore
'	 if rsTabella2(0)&"" =""  then
'	   numrsDomande2=0
'	 end if 
' 
  

 %>