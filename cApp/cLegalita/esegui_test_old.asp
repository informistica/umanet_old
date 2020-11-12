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

 ' 'sostituzioni inutilizzate: non cancellare e non utilizzare fino a prova contraria
 ' 'sAns=  Replace(sInput,"'",Chr(96)) 'sostituisco l'apice ' con quello storto per non dist. la sintassi
  ' sAns=  Replace(sAns,Chr(39),Chr(96)) 
  ' sAns=  Replace(sAns,Chr(44),Chr(96)) 
  ' sAns = Replace(sAns,chr(146),Chr(96))
  ' sAns = Replace(sAns,chr(147),Chr(96))
  ' sAns = Replace(sAns,chr(148),Chr(96))
  ' sAns = Replace(sAns,chr(239),chr(96))
  
  ' sAns = Replace(sAns,"’",chr(96))
  ' sAns = Replace(sAns,"gradi",chr(248))
 ' 'sAns = Replace(sAns, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto
  
   'sAns=  Replace(sInput,Chr(34),"") ' sostituisco " con niente per non disturbare la sintassi 
   'sAns=  Replace(sAns,Chr(36),"")  ' rimuovo il simbolo $
  
  sAns = Replace(sAns, Chr(34), "'") 
   sAns = Replace(sAns, "  ", " ") 'sostituizione doppio spazio con uno singolo
   sAns = Replace(sAns, "	", " ") 'sostituzione spazi per evitare errori
   sAns = Replace(sAns, " ?", "?") ' rimozione spazio prima del punto di domanda
   sAns = Replace(sAns, "’", "'") ' sostituzione di un'apice con quello classico
   sAns = Replace(sAns, "…", "...") 'sostituzione tre puntini
   sAns = Replace(sAns, Chr(25), "'") 'sostituizione apice
  
  ' 'sostituzione caratteri vari
  ' sAns = Replace(sAns,Chr(224),"a'") 'à
  ' sAns = Replace(sAns,Chr(232),"e'") 'è
  ' sAns = Replace(sAns,Chr(233),"e'") 'è
  ' sAns = Replace(sAns,chr(236),"i'") 'ì
  ' sAns = Replace(sAns,chr(237),"i'") 'ì
  ' sAns = Replace(sAns,chr(242),"o'") 'ò
  ' sAns = Replace(sAns,chr(243),"o'") 'ò
  ' sAns = Replace(sAns,chr(249),"u'") 'ù
  ' sAns = Replace(sAns,chr(250),"u'") 'ù
  ' sAns = Replace(sAns, "&#8230;", "...")
  ' sAns = Replace(sAns, "&#224;","a'") 'à
  ' sAns = Replace(sAns, "&#225;", "à") 'à
  ' sAns = Replace(sAns, "&#249;","u'") 'ù
  ' sAns = Replace(sAns, "&#8217;", "'")
  ' sAns = Replace(sAns, "&#8211;", "-")
  ' sAns = Replace(sAns, "&#232;","e'") 'è
  ' sAns = Replace(sAns, "&#233;","e'") 'è
  ' sAns = Replace(sAns, "&#242;","o'") 'ò
  ' sAns = Replace(sAns, "&#171;","'")
  ' sAns = Replace(sAns, "&#187;","'")
  ' sAns = Replace(sAns, "&#8220;","'")
  ' sAns = Replace(sAns, "&#8221;","'")
  ' sAns = Replace(sAns, "&#236;","i'") 'ì
  ' sAns = Replace(sAns, "&#250;","u'") 'ù
  ' sAns = Replace(sAns, "&#176;",chr(248)) 'gradi
  ' sAns = Replace(sAns, "'", "'")
  ' sAns = Replace(sAns, "&quot;", "") 'sostituzione delle virgolette alte con niente per evitare errori json
  
  sAns = Replace(sAns,chr(96),chr(39)) ' sostituizione finale dell'apice storto con il classico apice
  ' sAns1 = ucase(left(sAns,1)) ' maiuscola frasi
  ' sAns2 = right(sAns, len(sAns)-1)
  
  ' sAns = sAns1&sAns2
  
  
 ReplaceCar = sAns
 end function
%>

<% Sub setInQuizOrderBy()
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
'  On Error Resume Next  
    
		 
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
   'CodiceTest = Request.QueryString("CodiceTest") ' se svolgo tutto il modulo (stato=1) contiene l'Id del modulo e non del paragrafo
 'Tipo=Request.QueryString("Tipo")
 'If Tipo="" then 
 'Tipo=1
 'end if
 'Tipo=cint(Tipo)
	
 TestAbilitato=1

 'Sottoparagrafo=Request.QueryString("Sottoparagrafo")
' CodiceSottopar = Request.QueryString("CodiceSottopar") 
'  

 
  
			call setInQuizOrderBy()
			'orderby=1
			'Quiz=1
			
			 
		   ' Response.write("QUIZ="&Quiz)
		 %>
		 
		 <%
	 
	 
	 
	'if paragrafo <> 1 then	' qui bisogna aggiungere il controllo sulla stato (=1:paragrafo) (=0:modulo)
		'argomento = "Id_Mod"
		'QueryInQuiz = " AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1)"
	'else
		'argomento = "Id_Arg"
		'QueryInQuiz = " and In_Quiz=-1" 'questa sostituisce la riga sotto, ciò si rende necessario perchè le domande inserite da admin
		'devono comparire in tutte le batterie, credo non serva altro, hai parametrizzato il tutto molto bene, quindi basta modificare solo questa riga.  
		'QueryInQuiz = "" 'secondo me serve solo nel caso in cui paragrafo diverso da 1 -> altrimenti non facciamo neanche il controllo In_Quiz: prendiamo tutto admin e non
	'end if	
	
	 
	  ' if tipo=0 then ' domande vero/falso
	  ' ' if stato=1 then query sui con codice paragrafo else query con codice modulo
    ' QuerySQL="SELECT count(*)" &_
		   ' " FROM Domande " &_
		   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=1"&QueryInQuiz&";"
        
		' ' response.write(QuerySQL)
		' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' NumDom=0
		' if not rsTabella.eof then
		    ' NumDom=rsTabella(0)
	    ' end if
		
		' if NumDom>0 then
			 ' QuerySQL="SELECT Domande.*" &_
			   ' " FROM Domande " &_
			   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=1"&QueryInQuiz&" order by Domande." & order(orderby)& " asc;"
			
			' ' response.write(QuerySQL)
			' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' end if
	 ' end if
	 
	 
	 ' if tipo=1 then ' risposta singola 
	 ' ' qui bisogna aggiungere il controllo sulla stato (=1:paragrafo) (=0:modulo)
	  ' ' if stato=1 then query sui con codice paragrafo else query con codice modulo
    ' QuerySQL="SELECT count(*)" &_
		   ' " FROM Domande " &_
		   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=0"&QueryInQuiz&";"
        
		' ' response.write(QuerySQL)
		' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' NumDom=0
		' if not rsTabella.eof then
		    ' NumDom=rsTabella(0)
	    ' end if
		' if NumDom>0 then
			 ' QuerySQL="SELECT Domande.*" &_
			   ' " FROM Domande " &_
			   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=0"&QueryInQuiz&" order by Domande." & order(orderby)& " asc;"
			
			' ' response.write(QuerySQL)
			' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' end if
	 ' end if
	 
	  ' if tipo=2 then ' risposta multipla
	  ' ' qui bisogna aggiungere il controllo sulla stato (=1:paragrafo) (=0:modulo)
	  ' ' if stato=1 then query sui con codice paragrafo else query con codice modulo
    ' QuerySQL="SELECT count(*)" &_
		   ' " FROM Domande " &_
		   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=1  and Domande.VF=0"&QueryInQuiz&";"
        
		' ' response.write(QuerySQL)
		' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' NumDom=0
		' if not rsTabella.eof then
		    ' NumDom=rsTabella(0)
	    ' end if
		' if NumDom>0 then
			 ' QuerySQL="SELECT Domande.*" &_
			   ' " FROM Domande " &_
			   ' " WHERE Domande."&argomento&"='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=1  and Domande.VF=0"&QueryInQuiz&"order by Domande." & order(orderby)& " asc;"
			
			' ' response.write(QuerySQL)
			' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' end if
	 ' end if	 
	 	
		
		
		
		' if paragrafo <> 1 then
		 ' QuerySQL="SELECT Moduli.Titolo" &_
		   ' " FROM Moduli " &_
		   ' " WHERE Moduli.ID_Mod='" & CodiceTest & "';" 
		' else
		' QuerySQL="SELECT Paragrafi.Titolo" &_
		   ' " FROM Paragrafi " &_
		   ' " WHERE Paragrafi.ID_Paragrafo='" & CodiceTest & "';" 
		' end if
		
		' 'response.write(QuerySQL)
		' Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
		' TitoloModulo=rsTabella1("Titolo")
	
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
	
	QuerySQL = "SELECT TOP("&NumeroDomande/2&") * FROM Leg_Domande WHERE In_Quiz = '"&Quiz&"' and VF = 1 UNION SELECT TOP("&NumeroDomande/2&") * FROM Leg_Domande WHERE In_Quiz = '"&Quiz&"' and VF = 0"
	QuerySQL = QuerySQL & "order by "&order(orderby)&";"
	'response.write QuerySQL
	
	set rsDomande = ConnessioneDB.Execute(QuerySQL)
	'response.write(QuerySQL)
	
	response.write "{""domande"": """&NumeroDomande&""", ""tempo"": """&Tempo&""","
	
	i=0
	do while not rsDomande.EOF
		
		url = Server.MapPath(homesito)&"\DB1"&Mid(Replace(rsDomande("URL_Teoria"),"/","\"),3)
		
		If objFSO.FileExists(url) then
		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
  		sReadAll = ltrim(objTextFile.ReadAll	)
		sReadAll = Replace(sReadAll, vbNewLine, " ")
		sReadAll = ReplaceCar(sReadAll)
		sReadAll = Server.HTMLEncode(sReadAll)
		
		' if len(sReadAll)>450 then
			' sReadAll = left(sReadAll, 450)&"..."
		' end if
		
		sAns = Replace(sAns, left(sAns,1), ucase(left(sAns,1)))
		
	else	
	   sReadAll="Spiegazione mancante"
	end if	
		
	' response.write("""VF"&i&": """&rsDomande("VF")&""", ""domanda"&i&""": """&ReplaceCar(rsDomande("Quesito"))&""","  &_
 ' """risposta"&i&".1"": """&rsDomande("Risposta1")&""", ""risposta"&i&".2"": """&rsDomande("Risposta2")&""", ""risposta"&i&".3"": """&rsDomande("Risposta3")&""", ""risposta"&i&".4"": """&rsDomande("Risposta4")&""", ""rispostaesatta"&i&""": """&ReplaceCar(rsDomande("RispostaEsatta"))&"""," &_ 
 ' """url"&i&""": """&url&"""," &_ 
 ' """spiegazione"&i&""": """&ReplaceCar(sReadAll)&"""")
 
 response.write("""VF"&i&""": """&rsDomande("VF")&""", ""domanda"&i&""": """&Server.HTMLEncode(ReplaceCar(rsDomande("Quesito")))&""","  &_
 """risposta"&i&".1"": """&Server.HTMLEncode(ReplaceCar(rsDomande("Risposta1")))&""", ""risposta"&i&".2"": """&Server.HTMLEncode(ReplaceCar(rsDomande("Risposta2")))&""", ""risposta"&i&".3"": """&Server.HTMLEncode(ReplaceCar(rsDomande("Risposta3")))&""", ""risposta"&i&".4"": """&Server.HTMLEncode(ReplaceCar(rsDomande("Risposta4")))&""", ""rispostaesatta"&i&""": """&Server.HTMLEncode(ReplaceCar(rsDomande("RispostaEsatta")))&"""," &_  
 """spiegazione"&i&""": """&ReplaceCar(sReadAll)&"""")
	
		if i < (NumeroDomande-1) then
		response.write ","
		end if
 
		i=i+1
	rsDomande.movenext
	loop
	
	response.write "}"
	
' if NumDom>10 then
	' NumDom = 10
 ' end if

' i=1 'inizializza la variabile i (contatore delle domande)		  
' response.write(" { "  &_
 ' """classe"": """ & classe& """," &_
 ' """titolo"": """ & ReplaceCar(TitoloModulo)& ""","  &_
 ' """batteria"": """ & Quiz& ""","  &_ 
 ' """totale"": """&NumDom&"""")
 
 

 
' if NumDom>0 then

' response.write(", ")
	' While not rsTabella.EOF and i<11' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
         ' ' mi serve il titolo del paragrafo per il testo della spiegazione domanda
		  ' QuerySQL="SELECT Paragrafi.Titolo" &_
		   ' " FROM Paragrafi " &_
		   ' " WHERE Paragrafi.ID_Paragrafo='" & rsTabella("Id_Arg") & "';"  		   
	' '	response.write(QuerySQL)
		' Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)
		' TitoloParagrafo=rsTabella2("Titolo")
	    ' ID=rsTabella("CodiceDomanda")
		
		' if paragrafo <> 1 then
			' url=Server.MapPath(homesito)& "/Db1/Materie/materia_1/Expo/" &CodiceTest&"_Spiegazioni/"&CodiceTest&"_"&TitoloParagrafo&"_"&ID&".txt"
	    ' else
			' url=Server.MapPath(homesito)& "/Db1/Materie/materia_1/Expo/" &Left(CodiceTest, 6)&"_Spiegazioni/"&Left(CodiceTest, 6)&"_"&TitoloParagrafo&"_"&ID&".txt"
		' end if
		
		' url=Replace(url,"\","/") 
		
		
		
	 
		
    ' If objFSO.FileExists(url) then
		' Set objTextFile = objFSO.OpenTextFile(url, ForReading)
  		' sReadAll = ltrim(objTextFile.ReadAll	)
		' sReadAll = Replace(sReadAll, vbNewLine, " ")
		' sReadAll = Server.HTMLEncode(sReadAll)
		
		' if len(sReadAll)>450 then
			' sReadAll = left(sReadAll, 450)&"..."
		' end if
		
		' sAns = Replace(sAns, left(sAns,1), ucase(left(sAns,1)))
	' else	
	   ' sReadAll="Spiegazione mancante"
	' end if	
		
		
		' if Tipo=0 then  
	   ' if i<NumDom then 
		   ' response.write("""domanda"&i&""": """&ReplaceCar(rsTabella("Quesito"))&""","  &_
 ' """rispostaesatta"&i&""": """&ReplaceCar(rsTabella("RispostaEsatta"))&"""," &_ 
 ' """url"&i&""": """&url&"""," &_ 
 ' """spiegazione"&i&""": """&ReplaceCar(sReadAll)&""""&",")
       ' else
	      ' response.write("""domanda"&i&""": """&ReplaceCar(rsTabella("Quesito"))&""","  &_
 ' """rispostaesatta"&i&""": """&ReplaceCar(rsTabella("RispostaEsatta"))&"""," &_ 
 ' """url"&i&""": """&url&"""," &_ 
 ' """spiegazione"&i&""": """&rtrim(ReplaceCar(sReadAll))&"""")
	   ' end if
	   
    ' end if
	
	' if (Tipo=1) or (Tipo=2) then   
	    ' if i<NumDom then 
		   ' response.write("""domanda"&i&""": """&ReplaceCar(rsTabella("Quesito"))&""","  &_
 ' """risposta"&i&".1"": """&ReplaceCar(rsTabella("Risposta1"))&""","  &_
 ' """risposta"&i&".2"": """&ReplaceCar(rsTabella("Risposta2"))&""","  &_
 ' """risposta"&i&".3"": """&ReplaceCar(rsTabella("Risposta3"))&""","  &_
 ' """risposta"&i&".4"": """&ReplaceCar(rsTabella("Risposta4"))&""","  &_ 
 ' """rispostaesatta"&i&""": """&ReplaceCar(rsTabella("RispostaEsatta"))&"""," &_ 
 ' """url"&i&""": """&url&"""," &_ 
 ' """spiegazione"&i&""": """&rtrim(ReplaceCar(sReadAll))&""""&",")
        ' else
		 ' response.write("""domanda"&i&""": """&ReplaceCar(rsTabella("Quesito"))&""","  &_
 ' """risposta"&i&".1"": """&ReplaceCar(rsTabella("Risposta1"))&""","  &_
 ' """risposta"&i&".2"": """&ReplaceCar(rsTabella("Risposta2"))&""","  &_
 ' """risposta"&i&".3"": """&ReplaceCar(rsTabella("Risposta3"))&""","  &_
 ' """risposta"&i&".4"": """&ReplaceCar(rsTabella("Risposta4"))&""","  &_ 
 ' """rispostaesatta"&i&""": """&ReplaceCar(rsTabella("RispostaEsatta"))&"""," &_ 
 ' """url"&i&""": """&url&"""," &_ 
 ' """spiegazione"&i&""": """&rtrim(ReplaceCar(sReadAll))&"""")
		' end if
		
    ' end if
	
	
    ' i = i+ 1 
	
    ' rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande	
   ' Wend 
   		
' end if

  
               
 
' if NumDom>0 then
' 'sReadAll = url
  ' objTextFile.Close 
  ' Set objFSO = Nothing
 ' rsTabella2.Close : Set rsTabella2 = Nothing
 ' end if
' response.write("}")	
' rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 


ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		  
         
                     
                      
%>
  
   



                

