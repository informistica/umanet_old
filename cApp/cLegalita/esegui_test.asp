<%@ Language=VBScript %>

<%
Response.charset="utf-8" 'codifica caratteri speciali funzionante!! 
Call Response.AddHeader("Access-Control-Allow-Origin", "*") 

'paragrafo = Request.QueryString("paragrafo")

%>


<%
 Dim Num_Quiz,rand,Quiz,orderby
 Dim objFSO, objTextFile
 Dim sRead, sReadLine, sReadAll
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 
 function ReplaceCar(sInput)
 dim sAns
 sAns=sInput
  
  sAns = Replace(sAns, Chr(34), "'") 
   sAns = Replace(sAns, "  ", " ") 'sostituizione doppio spazio con uno singolo
   sAns = Replace(sAns, "	", " ") 'sostituzione spazi per evitare errori
   sAns = Replace(sAns, " ?", "?") ' rimozione spazio prima del punto di domanda
   sAns = Replace(sAns, "’", "'") ' sostituzione di un'apice con quello classico
   sAns = Replace(sAns, "…", "...") 'sostituzione tre puntini
   sAns = Replace(sAns, Chr(25), "'") 'sostituizione apice
 
  sAns = Replace(sAns,chr(96),chr(39)) ' sostituizione finale dell'apice storto con il classico apice
 
  
 ReplaceCar = sAns
 end function

 Sub setInQuizOrderBy()
' genera un numero casuale per scegliere quale quiz e quale ordinamento per le domande   
             'Num_Quiz=rsTabella(0) 
			Num_Quiz=4
			if Num_Quiz=-1 then
			   Quiz=-1
			   randomize()
			    do 
					rand=rnd()
				loop until (cint(left((rand*5),1))>0) and (cint(left((rand*5),1))<=7)
				orderby=left((rand*5),1)
			   
			else
			 
				'response.write("NUM_QUIZ="&Num_Quiz)
				randomize()
				do 
					rand=rnd()
				loop until (cint(left((rand*5),1))>0) and (cint(left((rand*5),1))<=Num_Quiz)
				Quiz=left((rand*5),1)
				' Response.write("QUIZ="&Quiz)
				 do 
					rand=rnd()
				loop until (cint(left((rand*5),1))>0) and (cint(left((rand*5),1))<=7)
				orderby=left((rand*5),1)
			end if
end sub %>
<% Response.Buffer=True %>
 

<%  
  'On Error Resume Next  
    
		 
 ' per generare un ordinamento casuale delle domande in base ad uno dei seguenti campi
 Dim order(8)
 
 
order(0)="" ' non lo uso 
order(1)="CodiceDomanda" 
order(2)="Quesito" 
order(3)="Risposta1"
order(4)="Risposta2"
order(5)="Risposta3"  
order(6)="Risposta4" 
order(7)="Data" 
 
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
		 
 
	 
		%> 
        <!-- #include file = "../../var_globali.inc" --> 
		<!-- #include file = "../include/stringa_connessione.inc" --> 
 	     
	 
                 
<%  

classe="Expo"	
TestAbilitato=1
call setInQuizOrderBy()
	'orderby=1
	'Quiz=1
	
	'NumeroDomande = 4 ' usare numero PARI!!!
	if request.querystring("ndomande")="" then
	  NumeroDomande=10
	else
	   NumeroDomande=cint(request.querystring("ndomande"))
	   if (NumeroDomande mod 2)<>0 then
		NumeroDomande=NumeroDomande-1 ' se non è pari lo rendo pari
	   end if
	end if
	
	Tempo = 30
	
	
	' conto il numero di domande disponibili per il quiz estratto
	 
	QuerySQL = "SELECT count(*)  FROM Leg_Domande WHERE In_Quiz = '"&Quiz&"'"
	' QuerySQL = QuerySQL & "order by "&order(orderby)&";"
	'response.write(QuerySQL)
	set rsDomande = ConnessioneDB.Execute(QuerySQL)
	ndomande=rsDomande(0)
	'response.write(ndomande)
	
	
	
	dim idxd()
	redim idxd(ndomande)
	dim used()
	redim used(ndomande)
	for i=0 to ndomande
		used(i)="false"
	next 
	Limite = ndomande-1	
	' genero il vettore idxd() di numeri casuali senza ripetizioni
	for i=0 to Limite 
		randomize()
		do 
			rand=CInt(Limite*Rnd() + 1 )
			if rand=0 then 
			  rand=1
			end if  
		loop until (used(rand)="false") 
		idxd(i)=rand  ' genero il vettore con gli indici random senza ripetizioni compresi tra 0 e ndomande-1
		' lo utilizzo per accedere tramite for i ai vettori paralleli che svolgono il ruolo del recordsetTabella
		used(rand)="true"
	next 
	
	 ' for i=1 to Limite
	    ' response.write(i&")"&idxd(i)&"<br>")
	 ' next
	
	
	
	dim quesito(),ra(),rb(),rc(),rd(),re(), spiega(),tipo()
	redim quesito(ndomande),ra(ndomande),rb(ndomande),rc(ndomande),rd(ndomande),re(ndomande), spiega(ndomande),tipo(ndomande)
	QuerySQL = "SELECT * FROM Leg_Domande WHERE In_Quiz = '"&Quiz&"'"
	QuerySQL = QuerySQL & " order by "&order(orderby)&";"
	set rsD = ConnessioneDB.Execute(QuerySQL)
	'response.write(QuerySQL)
	i=0
	do while not rsD.EOF
		tipo(i)=rsD("VF")
		'response.write("<br>t"&tipo(i))
		quesito(i)=rsD("Quesito")
		'response.write("<br>q"&quesito(i))
		ra(i)=rsD("Risposta1")
		rb(i)=rsD("Risposta2")
		rc(i)=rsD("Risposta3")
		rd(i)=rsD("Risposta4")
		re(i)=rsD("RispostaEsatta")
		'response.write(quesito(i)&"<br>"&ra(i)&"<br>"&rb(i)&"<br>"&rc(i))
		url = Server.MapPath(homesito)&"\DB1"&Mid(Replace(rsD("URL_Teoria"),"/","\"),3)
		'response.write("<br>"&url)
		If objFSO.FileExists(url) then
			Set objTextFile = objFSO.OpenTextFile(url, ForReading)
			sReadAll = ltrim(objTextFile.ReadAll)
			sReadAll = Replace(sReadAll, vbNewLine, " ")
			sReadAll = ReplaceCar(sReadAll)
			spiega(i) = sReadAll		
		else	
		   spiega(i)="Spiegazione mancante"
		end if	
		i=i+1
		rsD.movenext
	loop
	' inizio a stampare json
	response.write "{""domande"": """&NumeroDomande&""", ""tempo"": """&Tempo&""","
	
	i=0
	ndate=0  ' per saltare se un indice risulta vuoto
	while ndate<numerodomande
	'for i=0 to numerodomande
		
		'response.write("i="&x&"<br>")
		indice=idxd(i)
		'indice=10
		jsoncompleto="false"
		if tipo(indice)=1 then ' vero falso
			if quesito(indice)<>"" and re(indice)<>""  and spiega(indice)<>"" then
				jsoncompleto="true"
			end if
		else 
			 if quesito(indice)<>"" and ra(indice)<>"" and rb(indice)<>"" and rc(indice)<>"" and rd(indice)<>"" and re(indice)<>"" and spiega(indice)<>"" then
				jsoncompleto="true"
			end if
		end if
	
	
	if jsoncompleto="true" then
	
		 ndate=ndate+1
		 response.write("""VF"&i&""": """&tipo(indice)&""", ""domanda"&i&""": """&ReplaceCar(quesito(indice))&""","  &_
	 """risposta"&i&".1"": """&ReplaceCar(ra(indice))&""", ""risposta"&i&".2"": """&ReplaceCar(rb(indice))&""", ""risposta"&i&".3"": """&ReplaceCar(rc(indice))&""", ""risposta"&i&".4"": """&ReplaceCar(rd(indice))&""", ""rispostaesatta"&i&""": """&ReplaceCar(re(indice))&"""," &_  
	  """spiegazione"&i&""": """&ReplaceCar(spiega(indice))&"""")
		 
		   if ndate < (NumeroDomande) then
			response.write ","
			end if
	 
	 end if
 
 
	
	 
		
	    i=i+1
	 wend
	'next 
	
	response.write "}"
 

ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		  
         
                     
                      
%>
  
   



                

