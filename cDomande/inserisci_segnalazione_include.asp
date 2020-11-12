 <% 
 
   
	query="Select Segnalata, Id_Stud from Domande where CodiceDomanda="&rsTabella("CodiceDomanda")
	set record=ConnessioneDB.Execute(query)
	if record(0)=0 then
	
	 ' aggiunto ../cFrasi/ -> la notifica viene aperta in una pagina della cartella cMessaggi
	 Azione="<a  target=blank href=../cDomande/inserisci_valutazione.asp?daQuaderno=1&CodiceDomanda="&rsTabella("CodiceDomanda")&">Ho segnalato una tua domanda !</a>"
	 Motivazione="Ricontrolla la coerenza della domande, delle risposte e della spiegazione"
	 Testo=Motivazione
	 Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."
	 QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & rsTabella("Id_Stud") & "','" & Testo & "','" & Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
	'response.write(QuerySQL)
	 ConnessioneDB.Execute(QuerySQL)
	 
	
	QuerySQL ="UPDATE Domande SET Domande.Segnalata = 1 WHERE Domande.CodiceDomanda= "&rsTabella.fields("CodiceDomanda") & ";"
	ConnessioneDB.Execute(QuerySQL)
	
	end if
 
	
	 
	 %>