<%@ Language=VBScript %>
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	<style>
	<!--
	 li.MsoNormal
		{mso-style-parent:"";
		margin-bottom:.0001pt;
		font-size:12.0pt;
		font-family:"Times New Roman";
		margin-left:0cm; margin-right:0cm; margin-top:0cm}
	-->
	</style>
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi cancellare i dati degli altri studenti!")

location.href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>"
//location.href=window.history.back();
 }
 </script>
</head>





   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   Verifica = clng(Request.QueryString("Verifica"))
   MO=Request.QueryString("MO")
   Modulo=MO
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   On Error Resume Next
   
   tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 tDom=request.querystring("tDom")
 tFra=request.querystring("tFra")
 tNod=request.querystring("tNod")
 
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Capitolo=Request.QueryString("Capitolo")
   ID=Request.QueryString("CodiceDomanda")
    
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
d=request.querystring("cla")
'CodiceAllievo=request.querystring("cod")
   
    
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url3=Replace(url,"\","/")
url=url3
response.write(url)

QuerySQL ="select Id_Stud from Domande  WHERE Domande.CodiceDomanda=" &ID&";"
 set rsTabella=ConnessioneDB.Execute(QuerySQL)
 CodiceAllievo=rsTabella("Id_Stud")
 set rsTabella=nothing
 

 if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%
if Verifica=0 then ' se non è chiamata per verifica per per eliminazione allora la cancello
     QuerySQL ="DELETE FROM DOMANDE WHERE CodiceDomanda =" &ID&";"
else
	QuerySQL ="UPDATE Domande SET Domande.Segnalata = 0 WHERE Domande.CodiceDomanda=" &ID&";"
end if
	 ConnessioneDB.Execute(QuerySQL)


'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
'Create the FSO.
Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(url)
objFSO.DeleteFile url
'response.write(url)


If Err.Number = 0 Then
Response.Write "Cancellazione avvenuta! "

 Session("Datacla")=DataClaq
  Session("Datacla2")=DataClaq2
						 
							response.redirect "../cClasse/quaderno.asp?stile="&session("stile")&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&DataClaq2="&DataClaq2&"&DataClaq="& DataClaq&"&DataCla2="&DataClaq2&"&DataCla="& DataClaq&"&tCap="&tCap&"&tSot="& tSot&"&tDom="&tDom
		 
Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>
	</font>   
	 
	 
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
    </body>
    </html>
	