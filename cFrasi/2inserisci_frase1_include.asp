



<% 
' controllo coerenza di utente che inserisce in classe

QuerySQL1="Select Classe from Allievi where CodiceAllievo='" & CodiceAllievo&"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL1)
cla=rsTabella(0) ' se trovo Cs e Ct segnalo errore perchè si sono confusi gli utenti delle due classi
if (strcomp(Mid(cla, 2,2),Mid(Modulo, 2,2))<>0) and (session("Admin")=false) then
	erroreUtente=1
else
 	erroreUtente=0
end if

'response.write("<br>errore:"&Mid(cla, 2,2)&"-"&Mid(Modulo, 2,2)& "-erroreUtente="&erroreUtente)

if (erroreUtente = 0) then


	QuerySQL1="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1)
	Valutato=rsTabella.fields("Valutato")
	DVAbilitato=rsTabella.fields("DVAbilitato")
	'limiteBonus=rsTabella("limiteBonus")
	limiteBonus=1
	rsTabella.close



	bonusCompletato=1
	Valutato=1 ' lo metto a 1 uno per usare il campo valutato solo per le domande quiz
		if Valutato=1 then
		voto=1 ' aggiunto per evitare errore
					
					if len(trim(sintesi))<40 then
					if url1<>"" or url2<>"" or url3<>""  then ' se carico l'immagine va bene anche senza commento
							voto=1
							Segnalata=0
							Sintesi2=""
						else
							voto=0
							Segnalata=1
						end if

					else
						Voto=1 ' valore di default
						Segnalata=0
					end if
		end if
	'response.write("voto;"&voto)


			if DVAbilitato=1 and (bonusCompletato=0) then
				randomize()
				rand=rnd()
				if len(trim(sintesi))>400 then
				     if (clng(left(rand*100,1)) mod 2)= 0 then ' se il numero casuale è pari (testa o croce)
					 	Voto=2
						Sintesi2= "HAI OTTENUTO 1 BONUS!"
					 else
					    Voto=1
						Sintesi2= "POTEVI OTTENERE 1 BONUS! MA NON SEI STATO FORTUNATO"
					end if
					Segnalata=0
				else
				    Voto=1
					Segnalata=0
				end if
			end if
	else
	    Segnalata=1
	    Voto=0
	end if '  if (errore=0) then nella pagina 2inserisci_fras1.asp
  end if
if (erroreUtente = 0) then
	' se contiene dei link a pagine .html e php non controllo la lunghezza della risposta
'	if url1<>"" or url2<>"" or url3<>"" then
'	Voto=1
'	Segnalata=0
'	end if

	if  session("eccezione")=1 then
			randomize()
			rand=rnd()*100
		''	probabilita=70  ' memorizzato nel profilo dell'utente, più è alto e più è penalizzante, di default vale 50
		if session("recupero")=0 then
			if rand < session("Probabilita") then
			 Voto=0
			 else
			 Voto=1
			end if
		end if
			yesterday = Year(Date)&Right("0" & Month(Date),2)&  Right("0" & Day(Date() -1),2) 
			giorno= Right("0" & Day(Date() -1),2)
			mese=Right("0" & Month(Date),2)
			if (giorno=31) and (mese>1) then
			 yesterday = Year(Date)&Right("0" & Month(Date)-1,2)&  Right("0" & Day(Date() -1),2)
			end if

			QuerySQL = "Update Eccezioni_Frasi Set Scadenza='"&yesterday&"' WHERE Id_Stud='"&CodiceAllievo &"' and Id_Prefrase="&ID_Prefrase
			''elimino l'eccezione utilizzata
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)

	end if




	QuerySQL = "SELECT Scadenza FROM preFrasi WHERE ID_Prefrase="&ID_Prefrase
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	dataScadenza=rsTabella(0)

	QuerySQL = "SELECT CodiceFrase,Cartella FROM Frasi WHERE Id_Stud='"&CodiceAllievo &"' and Id_Prefrase="&ID_Prefrase
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	if  rsTabella.eof Then ' non esiste la inserisco '

				   if strcomp(preFrase,1)=0 then ' provengo da preFrase aggiungo il ID
				        id_pref=clng(ID_Prefrase)
							if AggRisFrase<>"" or url1<>"" then ' se sono in upload img
								img=1
							else
								img=0
							end if
							'se è inserita con eccezione metto la data di scadenza così posso usare la vista frasiOltreScadenza e non ho il problema di spostare le date dei compiti inseririti nel recupero nel periodo giusto
							if  session("eccezione")=1 or session("recupero")=1 then
								
								QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Id_Prefrase,Img,Segnalata,Id_Sottoparagrafo) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & dataScadenza & "','" & Voto & "','" & Cartella & "','" & FormatDateTime(now, 4) & "'," & id_pref & "," & img & "," & 1 & ",'" & CodiceSottopar &"';"
							else
								QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Id_Prefrase,Img,Segnalata,Id_Sottoparagrafo) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & Cartella & "','" & FormatDateTime(now, 4) & "'," & id_pref & "," & img & "," & Segnalata & ",'" & CodiceSottopar &"';"
							end if
					else
						 	if  session("eccezione")=1  or session("recupero")=1 then
								QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Img,Segnalata,Id_Sottoparagrafo) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & dataScadenza & "','" & Voto & "','" & Cartella & "','" & FormatDateTime(now, 4) & "," & img& "," & 1  & ",'" & CodiceSottopar &"';"
							else
								QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Img,Segnalata,Id_Sottoparagrafo) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & Cartella & "','" & FormatDateTime(now, 4) & "," & img& "," & Segnalata  & ",'" & CodiceSottopar &"';"
							end if
					end if
		     ConnessioneDB.Execute QuerySQL
			' response.write("query="&QuerySQL)
		'	prelava ID dell'ultimo record inserito

		    QuerySQL = "SELECT CodiceFrase,Cartella FROM Frasi WHERE CodiceFrase=(Select Max(CodiceFrase) FROM Frasi);"
		    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		    ID=rsTabella(0)
		    CARTA=rsTabella(1)

		
  else
   ' DEVO FIXARE QUI IL SALVATAGGIO COME BOZZA, SE ESISTE LA BOZZA NON VENGONO SALVATI GLI URL
	 ' semplicemte cancello eventuali url già presenti, non posso aggiornare perchè no chiave primaria, e le reinserisco.
   ID=rsTabella(0)
	 CARTA=rsTabella(1)
	 'QuerySQL="UPDATE Frasi SET Voto= "&Voto&", Segnalata="&Segnalata&" where CodiceFrase="&ID
	 
	 QuerySQL="DELETE FROM [Frasi_Img] WHERE Id_Frase="&ID
	 ConnessioneDB.Execute(QuerySQL)

  end if ' if  rsTabella.eof Then ' non esiste la inserisco '
 'response.write(QuerySQL&"<br>")


   'qua inserisco   le immagini (o le pagine html) linkate cpon url anzichè uploadate
			if url1<>"" then
			imgname="Img1"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url1 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			' response.write(QuerySQL&"<br>")
			end if
			if url2<>"" then
			imgname="Img2"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url2 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			' response.write(QuerySQL&"<br>")
			end if
			if url3<>"" then
			imgname="Img3"
			 QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & url3 & "','" & imgname & "';"
			 ConnessioneDB.Execute(QuerySQL)
			'' response.write(QuerySQL&"<br>")
			end if


	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"

	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DEL NODO

	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Create the FSO.

	url=Replace(url,"\","/")
	url_feedback=left(url,instr(url,".")-1)
	url_feedback=url_feedback&"_feedback.txt"
	url_feedback=Replace(url_feedback,"\","/")
	 
	

	
	'response.write("ulr spiegazio="&url)
	'response.write("<br>sintesi="&ltrim(Sintesi))
	set objFSO=Server.CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(url) then
		objFSO.DeleteFile url
	end if
	
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	objCreatedFile.WriteLine(ltrim(Sintesi))
	objCreatedFile.WriteLine(ltrim(Sintesi2))
	objCreatedFile.Close

	if Segnalata=1 then 'aggiungo notifica
		
		Motivazione="Risposta incompleta"
		Set objCreatedFile = objFSO.CreateTextFile(url_feedback, True)
		objCreatedFile.WriteLine(Motivazione)
		objCreatedFile.Close

		parametriurlnotifica = "Cartella="&Cartella&"&id_classe="&session("id_classe")&"&cod="&CodiceAllievo&"&CodiceTest="&CodiceTest&"&CodiceFrase="&ID&"&Paragrafo="&Paragrafo&"&Capitolo="&Capitolo&"&MO="&Modulo&"&VAL="&Voto&"&tCap="&tCap&"&tSot="&tSot&"&tFra="&tFra
		'response.write(parametriurlnotifica)

		' aggiunto ../cFrasi/ -> la notifica viene aperta in una pagina della cartella cMessaggi
		Azione="<a  target=blank href=../cFrasi/2inserisci_valutazione_frase.asp?"&parametriurlnotifica&">Ho segnalato una tua frase !</a>"
		Testo="Ricontrolla il compito"
		Commentatore="Admin A."
		QuerySQLNot="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Testo & "','" & Azione & "','" & now() & "','" & "informistica" & "','" & Commentatore & "';"
		'response.write(QuerySQLNot)
		ConnessioneDB.Execute(QuerySQLNot)

	end if

	'On Error Resume Next
	'response.write("eccezione="& session("eccezione")&"& rand="&rand&" proba="&probabilita&"<br>")
	If Err.Number = 0 Then
	Response.Write "Inserimento avvenuto! "
	Else
	Response.Write Err.Description
	Err.Number = 0
	End If

 else %>

  <div class="alert alert-error">
     <b><%=response.write("Attenzione stai tentando di inserire con l'utente dell'altra classe, effettua il logout e rientra.<br> Ricordati di non tenere aperte schede delle due classi contemporaneamente")%></b>
   </div>

 <%end if %>
