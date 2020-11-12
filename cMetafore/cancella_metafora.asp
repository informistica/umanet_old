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
<body>
    <div id="container">
	<div class="risultati_test">
	<font color=#FF0000 size="4">


   <% 
    Response.Buffer=True 
  
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
   
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
   'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Capitolo=Request.QueryString("Capitolo")
   ID=Request.QueryString("CodiceMetafora")
   id_classe=request.querystring("id_classe")

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
d=request.querystring("cla")
cod=request.querystring("cod")
if cod="" then
cod=request.querystring("CodiceAllievo")
end if
CodiceAllievo=cod  
  'homesito="/anno_2010-2011_ITC/ECDL"     
url=Server.MapPath(homesito)& "/Db"&Session("DB")& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url2=Server.MapPath(homesito)& "/Db"&Session("DB")& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_simula_"&ID&".txt"
'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url =Replace(url,"\","/")
url2 =Replace(url2,"\","/")

if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then 
 

'Select case CodiceTest
 Select case Paragrafo
   'Case  Cartella&"_U_3_3" 
     Case "Topolino ed Obiettivi"
	 
	 ' devo mantenere la lista concatenata di metafore, quindi prelevo pf della metafora che cancello e lo assegno alla precedente
	 QuerySQL0="Select Pi,Pf from M_Topolino where CodiceMetafora =" &ID&";"
	 response.write(QuerySQL0)
	 set rsTab =	ConnessioneDB.Execute(QuerySQL0)
	 
	 
	  piNew=rsTab(0)
	 pfNew=rsTab(1)
	 QuerySQL ="DELETE   FROM M_Topolino WHERE CodiceMetafora =" &ID&";"
	 QuerySQL1 ="UPDATE M_Topolino SET  Pi = "&piNew &", Pf = "&pfNew &" WHERE Pf =" &ID&";"
	 QuerySQL2 ="UPDATE M_Topolino SET  Pi = "&piNew &", Pf = "&pfNew &" WHERE Pf =" &ID&";"
   'Case  Cartella&"_U_3_5" 
   
   Case "Navigazione nella Rete della Vita"
    QuerySQL0="Select Pi,Pf from M_Navigazione where CodiceMetafora =" &ID&";"
	 set rsTab=	ConnessioneDB.Execute(QuerySQL0)
	 piNew=rsTab(0)
	 pfNew=rsTab(1)
      QuerySQL ="DELETE   FROM M_Navigazione  WHERE CodiceMetafora =" &ID&";"
	   QuerySQL1 ="UPDATE M_Navigazione SET  Pi = "&piNew &", Pf = "&pfNew &" WHERE Pf =" &ID&";"
 ' ConnessioneDB.Execute QuerySQL		
	'Case  Cartella&"_U_3_8" 
	Case "Relazione Cliente Servitore"
	  QuerySQL0="Select Pi,Pf from M_Desideri where CodiceMetafora =" &ID&";"
	 set rsTab=	ConnessioneDB.Execute(QuerySQL0)
	 piNew=rsTab(0)
	 pfNew=rsTab(1)
      QuerySQL ="DELETE   FROM M_Desideri  WHERE CodiceMetafora =" &ID&";"
	   QuerySQL1 ="UPDATE M_Desideri  SET Pi = "&piNew &", Pf = "&pfNew &" WHERE Pf =" &ID&";"
	  
end select   
	
	if ID<>0 then
	response.write(QuerySQL)
	ConnessioneDB.Execute(QuerySQL)
	end if
	if piNew<>0 then
	response.write(QuerySQL1)
	ConnessioneDB.Execute(QuerySQL1)
	end if
	if pfNew<>0 then
	response.write(QuerySQL2)
	ConnessioneDB.Execute(QuerySQL2)
	end if

'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
'Create the FSO.
Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(url)
if objFSO.FileExists(url) then
objFSO.DeleteFile url
end if
if objFSO.FileExists(url2) then
objFSO.DeleteFile url2
end if
'response.write(url)
On Error Resume Next
If Err.Number = 0 Then
	if Request.ServerVariables("HTTP_REFERER") <>"" then 
							
		 end if 

Response.Write "Cancellazione avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If




response.Redirect request.serverVariables("HTTP_REFERER") 

   %>
	</font>   
	 
		
      <h4><a href="../cClasse/studente_domande.asp?cla=<%=d%>&id_classe=<%=id_classe%>">Torna alla classifica ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
		<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 
						</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->
  
<% 
response.Redirect request.serverVariables("HTTP_REFERER") 
else 
response.write(  ucase(Session("CodiceAllievo"))& "=" & ucase(CodiceAllievo))
%>
   <BODY onLoad="showText();">
<% end if %>
	</body>
	</html>
	