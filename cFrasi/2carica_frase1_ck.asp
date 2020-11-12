<%@ Language=VBScript %>


        <%


  


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->






    <%


ID_Prefrase=Request("ID_Prefrase")
Quesito=Request("Quesito")
Risposta=Request("testo")
Modulo=Request("Modulo")
Paragrafo=Request("Paragrafo")
Cartella=Request("Cartella")
Img=Request("Img")
CodiceSottopar = Request("CodiceSottopar")
Sottoparagrafo=Request("Sottoparagrafo")
 url1=Request("txtImg1")
 url2=Request("txtImg2")
 url3=Request("txtImg3")

ID=1
ID=79

url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"

	'CREAZIONE FILE DI TESTO PER INSERIRE LA RISPOSTA

	Dim objFSO
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	 
	'Create the FSO.

	url=Replace(url,"\","/")
	'response.write("ulr spiegazio="&url)
	'response.write("<br>risposta="&Risposta)
	'response.write("<br>sintesi="&ltrim(Sintesi))
	set objFSO=Server.CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	messaggio = objTextFile.ReadAll
	objTextFile.Close
	response.write(messaggio)

		'  RESPONSE.WRITE(querysql)
	'	  If Err.Number = 0 Then
	 '      
	'		'Response.Write "Modifica avvenuta! "
	'			stato=1
	'			messaggio="Modifica avvenuta"
	'		Else
	'			stato=0
	'			messaggio=Err.Description
	'		Err.Number = 0
	'		End If

		''response.write(QuerySQL)
'

'		response.write(" { ")
'		 response.write("""stato"": """&stato&""","  &_
'		 """messaggio"": """&messaggio&"""")
'		 response.write("}")
'


%>
