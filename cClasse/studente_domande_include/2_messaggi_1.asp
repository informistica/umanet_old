<% 

'if strcomp(formattaDataCla(DataClaq2),"12/32/2013")=0 then
'response.write("ciaklklo")
'DataClaq2="12/31/2013"
'DataCla2="12/31/2013"
'response.write("<br>1"&formattaDataCla(DataClaq2))
'response.write("<br>2"& DataClaq2)
'response.write("<br>3"& cint(01))

'end if
if strcomp(request.cookies("Dati")("Admin"),"True")=0 and strcomp(cod,request.cookies("Dati")("CodAdmin"))=0 then
'QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
'" WHERE Avvisi.CodiceAllievo2='" & cod & "'" &_ 
'" and  Data>=#" &formattaDataCla(DataClaq)  &"#" &_
'	 " AND Data<=#" &  cdate(formattaDataCla2(DataClaq2)) &"#"&_ 
'" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod & "'" &_ 
" and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"
else

'response.write(formattaDataCla2(DataClaq2))
'response.write("<br>"& DataClaq2)
'QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
'" WHERE Avvisi.CodiceAllievo='" & cod & "'" &_ 
'" and  Data>=#" &formattaDataCla(DataClaq)  &"#" &_
'	 " AND Data<=#" & cdate(formattaDataCla2(DataClaq2)) &"#"&_ 
'" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "'" &_ 
" and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_
" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"

 
end if

QuerySQL2=QuerySQL


' " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 

'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni
appAvvisi= QuerySQL
 Set rsTabellaAvvisi = ConnessioneDB.Execute(QuerySQL) ' messaggi personali
 ' response.write("rsTabellaAvvisi:"&QuerySQL)
' carico i messaggi della lavagna da lavagna.mdb per gli avvisi sui compiti da svolgere
 QuerySQL="SELECT  *  " &_
" FROM FORUM_MESSAGES_CLASSI " &_
" WHERE Id_Classe ='" & id_classe & "' and comments<>'InizializzaDB' " &_
" and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataCla2)) &"', 104))"&_
" and ParentMessage=0 and Id_Social=1 and Visibile=1 ORDER BY DatePosted desc ;"

' non metto le date tanto li cancello i compiti vecchi
' " and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND DatePosted<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 

'" and  DatePosted>=#" &DataClaq &"#" &_
'" AND DatePosted<=#" & DataClaq2  &"#" &_ 
'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni 
 Set rsTabellaAvvisi2 = ConnessioneDB3.Execute(QuerySQL)  ' messaggi a tutta la classe
 'response.write(DataClaq2&"---rsTabellaAvvisi2:"&QuerySQL)
 appAvvisi2= QuerySQL
 
 
 	'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\log_messaggi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close
' 
 
 QuerySQL="SELECT  *  " &_
" FROM FORUM_MESSAGES_CLASSI " &_
" WHERE Id_Classe ='" & id_classe & "' and comments<>'InizializzaDB' " &_
" and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataCla2)) &"', 104))"&_
" and ParentMessage=0 and Id_Social=2  and Visibile=1  and Descrizione<>'Interrogazioni' and Descrizione<>'Feedback' ORDER BY DatePosted desc ;"
 Set rsTabellaDiario = ConnessioneDB3.Execute(QuerySQL)  ' messag
 
 'response.write(QuerySQL)
 
  QuerySQL="SELECT  *  " &_
" FROM FORUM_MESSAGES_CLASSI " &_
" WHERE Id_Classe ='" & id_classe & "' and comments<>'InizializzaDB' " &_
" and (DatePosted>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataCla2)) &"', 104))"&_
" and ParentMessage=0 and Id_Social=0   and Visibile=1 ORDER BY DatePosted desc ;"
 Set rsTabellaForum = ConnessioneDB3.Execute(QuerySQL)  ' messag
 
 
 
%>
 
 



 