<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
 
<% Response.Buffer=True %>


 

<%
 
if (Session("CodiceAllievo")="" or Session("Id_Classe")="") and condivisione <> 1 then
Response.Redirect "../../home.asp"
end if

if 	Session("DB") <> Request.QueryString("DB") and (Session("CodiceAllievo") = "ospite" or Session("CodiceAllievo") = "") then
Response.Redirect "../../home.asp"
end if

 %>


<html>
<head>
<title>Stampa frasi</title>
 


</head>

<%
  
 ' On Error Resume Next  
    ' per il controllo della validit� della sessione, se � scaduta -> nuovo login
   
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO,i
  Dim ConnessioneDB,rsTabella, QuerySQL,CodiceTest,StringaConnessione
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../var_globali.inc" --> 
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<% 
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
 Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  tipo=Request.QueryString("tipo")
  Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if
 'utile solo per impostare valore di tipo2
  criterio=""
  if strcomp(tipo,"Vero/Falso")=0 then
    criterio="and VF=1"
	tipo2="vf"
  end if
   if strcomp(tipo,"risposta chiusa singola")=0 then
    criterio="and VF=0 and Multiple=0"
	tipo2="sin"
  end if
   if strcomp(tipo,"risposta chiusa multipla")=0 then
    criterio="Multiple=1"
	tipo2="mul"
  end if
  
    dim fraz(4)
	dim mess(4)
   nomefile=codice_test&"_"&capitolo&"_"&paragrafo&"_"&tipo2&"_"&Lingua&".xml"
 
  
    CodiceAllievo=Request.QueryString("CodiceAllievo")
	' inutile in quanto la query viene soveascritto sotto
if clng(Stato)=1 then	
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Mod='"&Modulo&"' and Segnalata=0 " & criterio
 else
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"&Codice_Test&"' and Segnalata=0 " & criterio
 end if
  '-----
   ' mi faccio passare la query dalla pagina precedente di spiegazione che ha pi� filtri
 
  QuerySQL=Request.QueryString("QuerySQL")
  

%>
   

<body bgcolor="#FFFFFF">
<div id="container">  
 <div id="bloc_destra_cont" class="contenuti_login" style="width:95%">
<%

 response.write(QuerySQL&"<br>")
response.write("tipo2="&tipo2) 
 
 
 
 
 function ReplaceCar(sInput)
dim sAns
  
  sAns = sInput
  'sAns1 = sInput
  
 sAns = Replace(sInput,chr(236),"i"&Chr(96))
 sAns = Replace(sAns,chr(237),"i"&Chr(96))
 sAns = Replace(sAns,chr(242),"o"&Chr(96))
 sAns = Replace(sAns,chr(243),"o"&Chr(96))
 sAns = Replace(sAns,chr(249),"u"&Chr(96))
 sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
 sAns = Replace(sAns,chr(133),"a"&Chr(96))
 sAns = Replace(sAns,chr(138),"e'")
 sAns = Replace(sAns,"�","e'")
  sAns = Replace(sAns,chr(130),"e'")
 sAns = Replace(sAns, Chr(34), "'") 'sostituisco gli apici " con l'apice singolo
 sAns=  Replace(sAns,"'",Chr(96))  'sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
 sAns=  Replace(sAns,chr(58),Chr(44))  'sostituisco : con , per non disturbare la creazione del file
 sAns=  Replace(sAns,"&","e") 
 sAns=  Replace(sAns,"/","-") 
 sAns=  Replace(sAns,"\","-") 
 sAns=  Replace(sAns,"?",".") 
 sAns=  Replace(sAns,"*","x") 
 sAns=  Replace(sAns,"<","_")
 sAns=  Replace(sAns,">","_") 
   sAns = Replace(sAns,"�","e'" )
   sAns=  Replace(sAns,"'",Chr(96))
   sAns=  Replace(sAns,"�",Chr(96))
   sAns=  Replace(sAns,"�",Chr(96))
   sAns=  Replace(sAns,"�","a'")
   sAns=  Replace(sAns,"�","o'")
   sAns=  Replace(sAns,"�","u'")
   sAns = Replace(sAns,"�","'")
   sAns = Replace(sAns,"�","'")
   sAns = Replace(sAns,"�","'")
   sAns = Replace(sAns, Chr(96), "'")
   sAns = Replace(sAns, "�", "E'")
   sAns = Replace(sAns, "�", "i'")
   sAns = Replace(sAns, "�", "-")
   'sAns = Replace(sAns,VBCrlf,"<br>")
    
ReplaceCar = sAns

end function
 
 
 if (InStr(QuerySQL,"drop")=0) and (InStr(QuerySQL,"delete")=0) then
Set rsTabella = ConnessioneDB.Execute(QuerySQL)	

end if	
 
 
  
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFSO2 = CreateObject("Scripting.FileSystemObject")
 

 
'objFSO.CharSet = "UTF-8"

'objFSO2.CharSet = "UTF-8"

url2=Server.MapPath(homesito & "/script/cDomande/esportate_xml")&"/"& nomefile  
url2=Replace(url2,"\","/")
'response.write(url2)

Set objCreatedFile = objFSO2.CreateTextFile(url2, True)


riga="<?xml version=""1.0"" encoding=""UTF-8""?>"
objCreatedFile.WriteLine(riga)
riga="<quiz>"
objCreatedFile.WriteLine(riga)


%>

<% If rsTabella.BOF=True And rsTabella.EOF=True Then 

%> 
 
<% Else 
  
i=1 'inizializza la variabile i (contatore delle domande)
Do until rsTabella.EOF
  		 
	ID=rsTabella("CodiceDomanda")
	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
	url=Replace(url,"\","/")
	Set objTextFile = objFSO.OpenTextFile(url, ForReading) 
	sReadAll = objTextFile.ReadAll
	'sReadAll = url
	objTextFile.Close  

if strcomp(tipo,"Vero/Falso")=0	then
	riga="<!-- question: "& i & " -->"
	objCreatedFile.WriteLine(riga)
	riga="	<question type=""truefalse"">"
	objCreatedFile.WriteLine(riga)
	riga="		<name>"
	objCreatedFile.WriteLine(riga)
	riga="			<text>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"_"&codice_test&"_"&tipo2&"_"&i&"</text>"
	objCreatedFile.WriteLine(riga)
	riga="		</name>"
	objCreatedFile.WriteLine(riga)
	riga="		<questiontext format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</questiontext>"
	objCreatedFile.WriteLine(riga)
	riga="		<generalfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(sReadAll))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</generalfeedback>"
	objCreatedFile.WriteLine(riga)
	riga="		<defaultgrade>1.0000000</defaultgrade>"
	objCreatedFile.WriteLine(riga)
    riga="		<penalty>1.0000000</penalty>"
	objCreatedFile.WriteLine(riga)
    riga="		<hidden>0</hidden>"
	objCreatedFile.WriteLine(riga)
	if rsTabella("RispostaEsatta")=1 then 
		fraction_true=100
		msg_true="siamo d'accordo"
		fraction_false=0
		msg_false="non siamo d'accordo"
	else 
		fraction_true=0
		msg_true="Non siamo d'accordo"
		fraction_false=100 
		msg_false="Siamo d'accordo"
	end if
	riga="		<answer fraction="""&fraction_true&""" format=""moodle_auto_format"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text>true</text>"	
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_true&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)
	
	riga="		<answer fraction="""&fraction_false&""" format=""moodle_auto_format"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text>false</text>"	
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_false&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)
	riga="	</question>"
	objCreatedFile.WriteLine(riga)
end if ' fine vero/falso

 if strcomp(tipo,"risposta chiusa singola")=0 then
 
 riga="<!-- question: "& i & " -->"
	objCreatedFile.WriteLine(riga)
	riga="	<question type=""multichoice"">"
	objCreatedFile.WriteLine(riga)
	riga="		<name>"
	objCreatedFile.WriteLine(riga)
	riga="			<text>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"_"&codice_test&"_"&tipo2&"_"&i&"</text>"
	objCreatedFile.WriteLine(riga)
	riga="		</name>"
	objCreatedFile.WriteLine(riga)
	riga="		<questiontext format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</questiontext>"
	objCreatedFile.WriteLine(riga)
	riga="		<generalfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(sReadAll))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</generalfeedback>"
	objCreatedFile.WriteLine(riga)
	riga="		<defaultgrade>1.0000000</defaultgrade>"
	objCreatedFile.WriteLine(riga)
    riga="		<penalty>0.3333333</penalty>"
	objCreatedFile.WriteLine(riga)
    riga="		<hidden>0</hidden>"
	objCreatedFile.WriteLine(riga)
	riga="		<single>true</single>"
	objCreatedFile.WriteLine(riga)
    riga="		<shuffleanswers>true</shuffleanswers>"
	objCreatedFile.WriteLine(riga)
    riga="		<answernumbering>abc</answernumbering>"
	objCreatedFile.WriteLine(riga)
	riga="		<correctfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga=" 		 <text>Risposta corretta.</text>"
	objCreatedFile.WriteLine(riga)
    riga="		</correctfeedback>"
	objCreatedFile.WriteLine(riga)
    riga="		<partiallycorrectfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga=" 			<text>Risposta parzialmente esatta.</text>"
	objCreatedFile.WriteLine(riga)
    riga="		</partiallycorrectfeedback>"
	objCreatedFile.WriteLine(riga)
    riga=		"<incorrectfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga=" 			 <text>Risposta errata.</text>"
	objCreatedFile.WriteLine(riga)
    riga="		</incorrectfeedback>"
	objCreatedFile.WriteLine(riga)
	
	riga="<shownumcorrect/>"
	objCreatedFile.WriteLine(riga)
	
   Select case rsTabella("RispostaEsatta")
   case 1
	 fraction_1="100"
	 msg_1="Siamo d'accordo"
	 fraction_2="0"
	 msg_2="Non siamo d'accordo"
	 fraction_3="0"
	 msg_3="Non siamo d'accordo"
	 fraction_4="0"
	 msg_4="Non siamo d'accordo" 
   case 2
	 fraction_1="0"
	 msg_1="Non siamo d'accordo"
	 fraction_2="100"
	 msg_2="Siamo d'accordo"
	 fraction_3="0"
	 msg_3="Non siamo d'accordo"
	 fraction_4="0"
	 msg_4="Non siamo d'accordo"
   case 3
	 fraction_1="0"
	 msg_1="Non siamo d'accordo"
	 fraction_2="0"
	 msg_2="Non siamo d'accordo"
	 fraction_3="100"
	 msg_3="Siamo d'accordo"
	 fraction_4="0"
	 msg_4="Non siamo d'accordo"
   case 4
	 fraction_1="0"
	 msg_1="Non siamo d'accordo"
	 fraction_2="0"
	 msg_2="Non siamo d'accordo"
	 fraction_3="0"
	 msg_3="Non siamo d'accordo"
	 fraction_4="100"
	 msg_4="Siamo d'accordo"
End select
	

	
	riga="		<answer fraction="""&fraction_1&""" format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Risposta1")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_1&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)
	
	riga="		<answer fraction="""&fraction_2&""" format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Risposta2")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_2&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)
	
	riga="		<answer fraction="""&fraction_3&""" format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"&Server.HTMLEncode(ReplaceCar(rsTabella("Risposta3")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_3&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)
	
	riga="		<answer fraction="""&fraction_4&""" format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Risposta4")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& msg_4&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)

	riga="	</question>"
	objCreatedFile.WriteLine(riga)
 
 end if ' fine risposta singola
 
 
  if strcomp(tipo,"risposta chiusa multipla")=0 then
  
   riga="<!-- question: "& i & " -->"
	objCreatedFile.WriteLine(riga)
	riga="	<question type=""multichoice"">"
	objCreatedFile.WriteLine(riga)
	riga="		<name>"
	objCreatedFile.WriteLine(riga)
	riga="			<text>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"_"&codice_test&"_"&tipo2&"_"&i&"</text>"
	objCreatedFile.WriteLine(riga)
	riga="		</name>"
	objCreatedFile.WriteLine(riga)
	riga="		<questiontext format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Quesito")))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</questiontext>"
	objCreatedFile.WriteLine(riga)
	riga="		<generalfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(sReadAll))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="		</generalfeedback>"
	objCreatedFile.WriteLine(riga)
	riga="		<defaultgrade>1.0000000</defaultgrade>"
	objCreatedFile.WriteLine(riga)
    riga="		<penalty>0.3333333</penalty>"
	objCreatedFile.WriteLine(riga)
    riga="		<hidden>0</hidden>"
	objCreatedFile.WriteLine(riga)
	riga="		<single>false</single>"
	objCreatedFile.WriteLine(riga)
    riga="		<shuffleanswers>true</shuffleanswers>"
	objCreatedFile.WriteLine(riga)
    riga="		<answernumbering>abc</answernumbering>"
	objCreatedFile.WriteLine(riga)
	riga=		"<correctfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga=" 			 <text>Risposta corretta.</text>"
	objCreatedFile.WriteLine(riga)
    riga=		"</correctfeedback>"
	objCreatedFile.WriteLine(riga)
    riga="		<partiallycorrectfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga="			 <text>Risposta parzialmente esatta.</text>"
	objCreatedFile.WriteLine(riga)
    riga="		</partiallycorrectfeedback>"
	objCreatedFile.WriteLine(riga)
    riga=		"<incorrectfeedback format=""html"">"
	objCreatedFile.WriteLine(riga)
    riga=" 			 <text>Risposta errata.</text>"
	objCreatedFile.WriteLine(riga)
    riga="		</incorrectfeedback>"
	objCreatedFile.WriteLine(riga)
	
	riga="	<shownumcorrect/>"
	objCreatedFile.WriteLine(riga)
	 
	
   risp=cstr(rsTabella("RispostaEsatta"))
   	
   Select case cint(len(risp))
	   case 1
			fraction="100"
	   case 2
			fraction="50"
	   case 3
			fraction="33.33333"
	   case 4
			fraction="25" 
	End select

	for j=1 to 4
	
	if (InStr(risp,cstr(j))) <> 0 then
	   fraz(j-1)=fraction
	   mess(j-1)="Siamo d'accordo"
	else
	   fraz(j-1)="0"
	   mess(j-1)="Non siamo d'accordo"
	end if
	 

	next

for j=1 to 4
 
	
	riga="		<answer fraction="""&fraz(j-1)&""" format=""html"">"
	objCreatedFile.WriteLine(riga)
	riga="			<text><![CDATA[<p>"& Server.HTMLEncode(ReplaceCar(rsTabella("Risposta"&j)))&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)	
	riga="			<feedback format=""html"">"
	objCreatedFile.WriteLine(riga)	
	riga="			<text><![CDATA[<p>"& mess(j-1)&"<br></p>]]></text>"
	objCreatedFile.WriteLine(riga)
	riga="			</feedback>"
	objCreatedFile.WriteLine(riga)
	riga="		</answer>"
	objCreatedFile.WriteLine(riga)

next
   

	riga="	</question>"
	objCreatedFile.WriteLine(riga)
 
  
  end if ' fine risposta multipla
	
	
    i = i+ 1 
    rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
   Loop
   riga="</quiz>"
   objCreatedFile.WriteLine(riga)
   
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 
 %>
<br><br>
 
File xml generato, reperibile in: <br> <%=url2%>
 
 
</div>
</body>
 


</html>