<% ' prelevo e conto le frasi del paragrafo
	 QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 " &_
	 " WHERE CodiceAllievo='" & cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
	 " AND ID_MOD='"& rsTabellaModuli("ID_Mod")&"'"&_
 " order by In_Umanet,Posizione,Data,Ora;"
   Set rsTabellaFrasi = ConnessioneDB.Execute(QuerySQL)
  ' response.write("??"&QuerySQL)
   
	
   ' conto frasi e punti del paragrafo
   ' per riepilogare il totale di tutte le frasi 
	' QuerySQL1="SELECT Count(*) AS Num FROM  MODULO_PARAGRAFO_FRASI1  " &_
'	 " WHERE CodiceAllievo='" & cod & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
'	  " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo")&"';"	 
'	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
'	  numrsFrasi=rsTabella1(0)
'	 
'	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM  MODULO_PARAGRAFO_FRASI1  " &_
'	 " WHERE CodiceAllievo='" & cod & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
'	  " AND ID_Paragrafo='"&rsTabellaParagrafi("ID_Paragrafo")&"';"
'	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
'	   numrsFrasi2=rsTabella2(0)
'	 ' se non restituisce nulla serve per dargli un valore
'	 if rsTabella2(0)&"" =""  then
'	   numrsFrasi2=0
'	 end if 
   
   %>