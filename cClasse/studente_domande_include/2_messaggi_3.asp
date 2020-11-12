<% ' questo e per le notifiche in navigation pesca solo quewlli no
'response.write("DataCla="&DataCla & "SessionDataClaq="&Session("DataClaq"))
'if DataCla="" and Session("DataClaq")<>"" then
' DataCla=Session("DataClaq")
' DataCla2=Session("DataClaq2")
'else
'    DataCla=DataClaq
'    DataCla2=DataClaq2
'end if
function formatta_data_LO(Data)
    DataTest=formatDateTime(Data,2)
    'gira_data=Day(DataTest)&"/"&Month(DataTest)&"/"&Year(DataTest)
  
    if day(DataTest) < 10 then
    giorno="0" & day(DataTest) 
	else
	giorno=day(DataTest)
    end if
	
	if len(year(DataTest) ) = 2 then
	anno="20"& year(DataTest)
	elseif len(year(DataTest) ) =  3 then
	anno="2"& year(DataTest)
	else
	anno=year(DataTest)
	end if
    if month(DataTest) < 10 then
    mese="0" & month(DataTest) 
	else
	mese=month(DataTest)
    end if
	' pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
'	  if (left(pathEnd1,10)<>"c:\inetpub") then
'		 locale=1
'	  else
'		 locale=0
'	  end if 	
'	 ' response.write(left(pathEnd1,10))
'	'response.write("locale="&locale)
'     if locale=1 then
'	     formatta_data_LO = giorno & "/" & mese& "/" & anno   
'	 else				 
'		 formatta_data_LO = mese & "/" & giorno& "/" & anno  
'     end if 
	formatta_data_LO = giorno & "/" & mese& "/" & anno    
end function
' aggiunto il 8/10/2014
if DataClaq="" then
  DataClaq=Session("DataCla")
  DataCla=Session("DataCla")
end if
if DataClaq2="" then
  DataClaq2=Session("DataCla2")
  DataCla2=Session("DataCla2")
end if

'aggiunto in data 18/09/2015
DataClaq=formatta_data_LO(DataClaq)


	'DataClaq2=formatdatetime(now(),2)
'	DataClaq2="19/03/2015"
'	
'	d=CDate(DataClaq2)
'response.write(FormatDateTime(d) & "<br />")
'response.write(FormatDateTime(d,1) & "<br />")
'response.write(FormatDateTime(d,2) & "<br />")
'response.write(FormatDateTime(d,3) & "<br />")
'response.write(FormatDateTime(d,4) & "<br />")
'
'response.write("inizio"& formatta_data_LO(DataClaq))
'response.write("fine"& formatta_data_LO(1+cdate(DataClaq2)))

'	
	
'DataClaq="12/12/2015"
if session("Admin")=True and cod=Session("CodAdmin") and 1=0 then


' anche qui tolgo la query diversificata tra admin e utente normale!!

'on error resume next
' lo devo mettere perchè da un errore misterioso sulla query : parametri previsti 1, non ne vengo a capo perchè è la stessa identica query che in altre pagine non da errore, va in errore solo in home_uecdl_app e quindi in questa pagina non funziona il numero di notifiche presenti nel centro messaggi


'QuerySQL0="SELECT count(*) FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo WHERE Avvisi.CodiceAllievo='informistica';"
'response.write("DataClaq="&DataClaq & " " & "DataClaq2="&DataClaq2)
QuerySQL0="SELECT count(*) as NUM FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod & "' and Avvisi.Visto=0" &_ 
   " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &formatta_data_LO(1+CDATE(DataClaq2)) &"', 104))"&_
";"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod & "' and Visto=0" &_ 
  " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &formatta_data_LO(1+CDATE(DataClaq2)) &"', 104))"&_
" ORDER BY Data desc ;"

else

' QuerySQL0="SELECT count(*) as NUM FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
' " WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0" &_ 
   ' " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 ' " AND (Data<= CONVERT(DATETIME,'" &formatta_data_LO(1+CDATE(DataClaq2)) &"', 104))"&_
' ";"
' QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
' " WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0" &_ 
   ' " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 ' " AND (Data<= CONVERT(DATETIME,'" &formatta_data_LO(1+CDATE(DataClaq2)) &"', 104))"&_
' "  ORDER BY Data desc ;"


QuerySQL0="SELECT count(*) as NUM FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0" &_ 
";"

QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "' and Visto=0" &_ 
"  ORDER BY Data desc ;"



end if

' per farla funzionare via web ho cambiato 01/13/214 in 13/01/2104 MISTERO!! perchè direttamete sul server funziona così com'è
'QuerySQL0="SELECT count(*) as NUM FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo"&_ 
'" WHERE Avvisi.CodiceAllievo='informistica' and Visto=0 and (Data>='13/01/2014') AND (Data<= '10/11/2014');"
'QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo"&_ 
'" WHERE Avvisi.CodiceAllievo='informistica' and Visto=0 and (Data>='13/01/2014') AND (Data<= '10/11/2014') ORDER BY Data desc ;" 
'response.write("2messaggi:"&QuerySQL0)
'response.write("<br>2messaggi:"&QuerySQL)

'response.write("2messaggi:"&QuerySQL0)
 Set rsTabellaAvvisiP = ConnessioneDB.Execute(QuerySQL0)
 numMessaggi=rsTabellaAvvisiP(0)
 Set rsTabellaAvvisiP = ConnessioneDB.Execute(QuerySQL) ' messaggi personali

 

 'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\anno_2013-2014\log0.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL0)
'				objCreatedFile.Close
'				
'					url="C:\Inetpub\umanetroot\anno_2013-2014\log1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				
 
 
%>
 
 



 