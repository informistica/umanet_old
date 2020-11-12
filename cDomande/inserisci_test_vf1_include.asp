 

 <%

 
   
   
  '  if strcomp(Multiple,"1")=0  then
'     'se non devo inserire domanda multipla pongo a 0 il campo 
'		Multiple=1
'	 else
'		Multiple=0
' 	 end if 
'	response.write("<br>strcomp(RE,1)="&strcomp(RE,"1"))
	'response.write("<br>strcomp(RE,2)="&strcomp(RE,"2"))
	
	' if Multiple=0 then
'		   if  (strcomp(RE,"1")<> 0) and (strcomp(RE,"2")<> 0) and (strcomp(RE,"3")<> 0) and (strcomp(RE,"4")<> 0)then 
'			  errore=1
'		   end if 
'	 else
'			  errore=0
'	 end if 
'  ' if Multiple<>"0" then
'    'response.Write("Multiple=1 ? 0=si "& strcomp(Multiple,"1") & "=" & Multiple)
'	if (Multiple=1) then
'       ' controllo validità numero che indica la risposta esatta deve appartenere alla tabella di corrispondenza
'       esiste=controlla(RE) 'NB DA SISTEMARE !!!!!!!!!!!!!
'       if esiste = 0 then
'          errore = 3
'	   end if 	   
'	   
'  	 else
'	   if ((RE<1) or (RE>4)) then 
'		  errore=1
'	   end if
'  	 end if 
   
   if ( (len(Domanda)=0) or (len(Spiegazione)=0)  ) then 
       errore=2
   end if
  ' response.write("<br>Domanda"&Domanda)
'    response.write("<br>Spiegazione"&Spiegazione)
'	 response.write("RE<br>"&RE)
' 
' 
 
   'Domanda1=Domanda
   'response.write("Domanda="&Domanda1)
if strcomp(preDomanda,1)=0 then
   ID_Predomanda=clng(ID_Predomanda)
else
   ID_Predomanda=0
end if
  
 ' response.write("Errore="&errore)




'errore=0	



	 Segnalata=0
 if (errore=0) then
		 ' inserisci la domanda
			 ' devo vedere se il setting è tale da richiedere voto=1 come default oppure no
			QuerySQL1="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
			Valutato=rsTabella.fields("Valutato") 
			DVAbilitato=rsTabella.fields("DVAbilitato")
			rsTabella.close
			if Valutato=1 then
			     
				 if len(Spiegazione)<100 then
							voto=0
							Segnalata=1
							
						 else
						     Segnalata=0
							 Voto=1 ' valore di default
						 end if
				 
				 if (DVAbilitato=1)and (len(Spiegazione)>350)   then
					    voto=2
						Segnalata=0
				 end if 
	 
			else
			    Voto=0
			end if
			
	   ' if (strcomp(Tipo,"1")=0) then ' se la domanda è plus metto nullo il campo che poi aggiorno con l'url del file di testo che contiene la domanda dopo
			 if Img="" then ' se sono di tipo 1 testo plus e senza immagine
				Img=0
			 end if
			 if ID_Predomanda="" then ' se sono di tipo 1 testo plus e senza immagine
				ID_Predomanda=0
			 end if
			 'se sono domanda normale il quesito è contenuto in DOmanda, quindi assegno a Titolo che finisce nella query unica
			 if strcomp(Tipo,"0")=0 then
				Titolo= Domanda 
			 end if
			 
			 if RE="" then
			  RE=0
			 end if
			 'FORSE possa fare una query unica 
			  QuerySQL="INSERT INTO Domande (Quesito, Risposta1, Risposta2,Risposta3,Risposta4,RispostaEsatta,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Tipo,In_Quiz,Multiple,ID_Predomanda,Img,VF,Ora,Id_Sottoparagrafo,Segnalata,Lingua) SELECT '"& ReplaceCar(Titolo) & "','"& R1 & "', '" & R2 & "','" & R3 & "','" & R4 & "','" & RE & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "'," & Voto & ",'" & Cartella & "','" & Tipo & "','" & In_Quiz_Stud &"','" & Multiple &"'," & ID_Predomanda &"," & Img  &"," & 1  &",'" & FormatDateTime(now, 4)  & "','" & CodiceSottopar  & "','" & Segnalata & "','" & lingua & "';"
		 
		 
		
	   ConnessioneDB.Execute QuerySQL 
	  ' response.write(QuerySQL)
	'	prelava ID dell'ultimo record inserito
	
		QuerySQL = "SELECT CodiceDomanda,Cartella FROM Domande WHERE CodiceDomanda=(Select Max(CodiceDomanda) FROM Domande);" 
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		ID=rsTabella(0)
		CARTA=rsTabella(1)
	       
		url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"  
		  %><br><%
	 
		url1= "../Materie/"&Session("ID_Materia")& "/" & CARTA & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	   
		'if (strcomp(Tipo,"1")=0)  then ' se la domanda è plus aggiorno con l'url del file di testo che contiene la domanda 
		  url4=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
		 ' url5=  CARTA & "/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
			  QuerySQL ="UPDATE Domande SET Quesito='" & Titolo & "' , URL_Teoria = '" & url1 & "' WHERE CodiceDomanda =" &ID&";"
	   ' else
		  '    QuerySQL ="UPDATE Domande SET URL_Teoria = '" & url1 & "' WHERE CodiceDomanda =" &ID&";"
		'end if 
		' Set objFSO = CreateObject("Scripting.FileSystemObject")  
	'  	url2="C:\Inetpub\umanetroot\anno_2012-2013\logInserisciDOmanda2.txt"
	'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
	'				objCreatedFile.WriteLine(QuerySQL)
	'				objCreatedFile.Close 
		ConnessioneDB.Execute(QuerySQL)
	'response.write(QuerySQL & "<br>")
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA
	
	Dim objFSO,objCreatedFile
	'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	'Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	url=Replace(url,"\","/")
	 
	 
	'response.write("?"&url)
	 
					
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	objCreatedFile.WriteLine(ReplaceCar(Spiegazione))
	objCreatedFile.Close
	
	if strcomp(Tipo,"1")=0 then 'CREAZIONE FILE DI TESTO PER INSERIRE LA DOMANDA
		url4=Replace(url4,"\","/")
		Set objCreatedFile = objFSO.CreateTextFile(url4, True)
		objCreatedFile.WriteLine(Domanda)
		objCreatedFile.Close
	end if 
	'response.write("<br>" & url4)
	
	 
		'If Err.Number = 0 Then
'			Response.Write "<span class='alert-success'>Inserimento avvenuto! </span>"
'		Else
'			Response.Write Err.Description 
'			Err.Number = 0
'		End If

 %>
