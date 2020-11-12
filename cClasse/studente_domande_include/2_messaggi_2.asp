<% 
if session("Admin")=True and cod=Session("CodAdmin") and 1=0 then


' tolgo la query diversificata tra admin e utente normale ( non capisco il perchè!! ) -> mostro le notifiche come destinatario non come mittente


'QuerySQL0="SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
'" WHERE Avvisi.CodiceAllievo2='" & cod  & "' and Visto=0" &_ 
'" and  Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
'	 " AND Data<=#" & 1+ cdate(mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_ 
'";"


'QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
'" WHERE Avvisi.CodiceAllievo2='" & cod  & "' and Visto=0" &_ 
'" and  Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
'	 " AND Data<=#" & 1+ cdate(mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_ 
'" ORDER BY Data desc ;"

'response.write(QuerySQL)


QuerySQL0 = "SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod  & "' and Visto=0 and Testo <> Azione" &_
";"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod  & "' and Visto=0 and Testo <> Azione" &_ 
" ORDER BY Data desc ;"


else


' QuerySQL0="SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
' " WHERE Avvisi.CodiceAllievo='" & cod  & "' and Visto=0" &_ 
' " and  Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
	 ' " AND Data<=#" & 1+ cdate( mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_ 
' ";"
' QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
' " WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0" &_ 
' " and  Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
	 ' " AND Data<=#" & 1+ cdate( mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_ 
' "  ORDER BY Data desc ;"


QuerySQL0="SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod  & "' and Visto=0 and ((CAST(Testo as ntext) NOT LIKE Azione) OR Testo IS NULL)" &_  
";"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0 and ((CAST(Testo as ntext) NOT LIKE Azione) OR Testo IS NULL)" &_ 
"  ORDER BY Data desc ;"


QuerySQL0_Letti="SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod  & "' and Visto=1 and ((CAST(Testo as ntext) NOT LIKE Azione) OR Testo IS NULL)" &_  
";"

QuerySQL_Letti="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=1 and ((CAST(Testo as ntext) NOT LIKE Azione) OR Testo IS NULL)" &_ 
"  ORDER BY Data desc ;"

end if

 Set rsTabellaAvvisiP = ConnessioneDB.Execute(QuerySQL0)
 numMessaggi=rsTabellaAvvisiP(0)
 Set rsTabellaAvvisiP = ConnessioneDB.Execute(QuerySQL)
				
 'response.write(QuerySQL0&"<br>"&QuerySQL)
 
 
 
 Set rsTabellaAvvisiLetti = ConnessioneDB.Execute(QuerySQL0_Letti)
 numMessaggiArchivio=rsTabellaAvvisiLetti(0)
 Set rsTabellaAvvisiLetti = ConnessioneDB.Execute(QuerySQL_Letti)
  
 
%>
 
 



 