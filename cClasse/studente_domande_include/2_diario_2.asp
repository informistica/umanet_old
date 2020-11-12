<%
	  'conto i post totali 
			  ' QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
' " FROM FORUM_MESSAGES " &_
' " WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage=0 and comments<>'InizializzaDB' and Id_Social=2 " &_
' " and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 ' " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataCla2)) &"', 104))"


	  QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage=0 and comments<>'InizializzaDB' and Id_Social=2 " &_
" ;"
	 
	 

 'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close
				
   Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	num_post_totali=rsTabella1(0)
	'num_post_totali_punti=0
	num_post_totali_punti=rsTabella1(1)
	if isnull(num_post_totali_punti)  then
	   num_post_totali_punti=0
	end if
	
	 
				
	 ' QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
' " FROM FORUM_MESSAGES " &_
' " WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage <>0 and comments<>'InizializzaDB' and Id_Social=2 " &_ 
' " and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 ' " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq)) &"', 104))"

	 
	 QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage <>0 and comments<>'InizializzaDB' and Id_Social=2 " &_ 
" ;"


   Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	num_messaggi=rsTabella1(0)
	'num_messaggi_punti=0
	num_messaggi_punti=rsTabella1(1)
	if isnull(num_messaggi_punti) then
	   num_messaggi_punti=0
	end if
%>

 
<%' carico i messaggi del forum
' QuerySQL="SELECT * " &_
' " FROM FORUM_MESSAGES " &_
' " WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and comments<>'InizializzaDB' and Id_Social=2" &_ 
' " and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 ' " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataCla2)) &"', 104))"&_
' " ORDER BY ID desc;"

QuerySQL="SELECT * " &_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and comments<>'InizializzaDB' and Id_Social=2" &_ 
" ORDER BY ID desc;"
 
 
 
 Set rsTabellaDiario = ConnessioneDB3.Execute(QuerySQL)
 appQuery=QuerySQL
 'response.write(QuerySQL)
 'response.write( DataCla  )
 
 
 
 
 
 %>
 
  
 

