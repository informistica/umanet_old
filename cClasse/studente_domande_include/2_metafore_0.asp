<% 
 
  metafore=1 ' se non faccio query diventa 
  Select Case rsTabellaParagrafi("ID_Paragrafo")%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
                                
 <%QuerySQL="SELECT * FROM  Elenco_Metafore_Topolino WHERE CodiceAllievo='" & cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_
	  " AND ID_Paragrafo='"&Cartella&"_U_2_3"&"';" 
	  %>
           
								<% Case Cartella&"_U_2_5" 'Navigazione%>
  <%QuerySQL="SELECT * FROM  Elenco_Metafore_Mavigazione WHERE CodiceAllievo='" & cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_
	  " AND ID_Paragrafo='"&Cartella&"_U_2_5"&"';" 
	   %>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
                                     <%QuerySQL="SELECT * FROM  Elenco_Metafore_Desideri WHERE CodiceAllievo='" & cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_
	  " AND ID_Paragrafo='"&Cartella&"_U_2_8"&"';" 
	%>
 								<%case else %>
                                <%
								metafore=0 ' serve dopo per evitare di accedere ad rstabMetafore che è nulla se 
								' faccio query ad ok nulla per avere rsTabellametafore.eof utilizzabile più avnati
								QuerySQL="SELECT * FROM  Elenco_Metafore_Topolino WHERE CodiceAllievo='" & cod & "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104))"&_
	  " AND ID_Paragrafo='pippo';" 
								%>
							<%End Select%>
 
<%
'response.write(rsTabellaParagrafi("ID_Paragrafo"))
  Set rsTabellaMetafore = ConnessioneDB.Execute(QuerySQL)
%>
 