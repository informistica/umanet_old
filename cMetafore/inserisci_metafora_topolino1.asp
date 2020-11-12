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
function showText2() {window.alert("La sessione � scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>

</head>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validit� della sessione, se � scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Sintesi
   Dim Topolino,Formaggio,Fame,Labirinto,Strada,Strada_KO,Strada_OK,Testata,Distanza
   
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   daSimulazione = Request.QueryString("daSimulazione")
   daSviluppa = Request.QueryString("daSviluppa")
   daTopolino = Request.QueryString("daTopolino") ' settato se sono chiamato per lo sviluppo della metafora
   CodiceTest = Request.QueryString("CodiceTest")
   Li = cint(Request.QueryString("Li"))
   ThreadParent = cint(Request.QueryString("ThreadParent"))
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   prenodo=Request.QueryString("prenodo") ' serve per capire il chiamante e quindi sapere se alla fine devo redirectare ad home_ver o home_app
   
   
    
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
  ' CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
  ' homesito="/anno_2010-2011_ITC/ECDL"
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")
	CodiceMetafora=Request.QueryString("CodiceMetafora")
	Paragrafo=Request.QueryString("Paragrafo")
	Modulo=Request.QueryString("Modulo")
	DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
 
	   Topolino = ucase(Request.Form("txtTopolino"))
	   Topolino = Replace(Topolino, Chr(34), "'")
	   Topolino=  Replace(Topolino,"'",Chr(96))
  
   Formaggio = ucase(Request.Form("txtFormaggio"))
   Formaggio = Replace(Formaggio, Chr(34), "'")
   Formaggio=  Replace(Formaggio,"'",Chr(96))


   Fame = ucase(Request.Form("txtFame"))
   Fame = Replace(Fame, Chr(34), "'")
   Fame=  Replace(Fame,"'",Chr(96))

   Labirinto = ucase(Request.Form("txtLabirinto"))
   Labirinto = Replace(Labirinto, Chr(34), "'")
   Labirinto=  Replace(Labirinto,"'",Chr(96))
   Strada = ucase(Request.Form("txtStrada"))
   Strada = Replace(Strada, Chr(34), "'")
   Strada=  Replace(Strada,"'",Chr(96))

   Strada_KO = ucase(Request.Form("txtStrada_ko"))
   Strada_KO = Replace(Strada_KO, Chr(34), "'")
   Strada_KO=  Replace(Strada_KO,"'",Chr(96))
   
   
   
   Strada_OK = ucase(Request.Form("txtStrada_ok"))
   Strada_OK = Replace(Strada_OK, Chr(34), "'")
   Strada_OK=  Replace(Strada_OK,"'",Chr(96))
   
   Testata = ucase(Request.Form("txtTestata"))
   Testata = Replace(Testata, Chr(34), "'")
   Testata=  Replace(Testata,"'",Chr(96))
   
   
   Distanza = ucase(Request.Form("txtDistanza"))
   Distanza = Replace(Distanza, Chr(34), "'")
   Distanza=  Replace(Distanza,"'",Chr(96))
   
   Sintesi=ucase(Request.Form("S1"))
   Sintesi= Replace(Sintesi, Chr(34), "'")
   Sintesi=  Replace(Sintesi,"'",Chr(96))
    Sintesi=  Replace(Sintesi,Chr(39),Chr(96))

if daSimulazione<>"" then ' aggiorno 
 
 
   Strada_KO = ucase(Request.Form("txtStradaKO"))
   Strada_KO = Replace(Strada_KO, Chr(34), "'")
   Strada_KO=  Replace(Strada_KO,"'",Chr(96))
   
   
   
   Strada_OK = ucase(Request.Form("txtStradaOK"))
   Strada_OK = Replace(Strada_OK, Chr(34), "'")
   Strada_OK=  Replace(Strada_OK,"'",Chr(96))
 end if

   
   
   
   if ( (len(Topolino)=0) or (len(Formaggio)=0) or (len(Fame)=0) or (len(Labirinto)=0) or (len(Strada)=0) or (len(Strada_KO)=0) or(len(Strada_OK)=0) or(len(Testata)=0) or(len(Distanza)=0)) then
 '  Response.Redirect("inserisci_test.asp?Cartella=Cartella&Num=0&Cognome=Cognome&Nome=Nome&CodiceTest=CodiceTest&Capitolo=Capitolo&Paragrafo=Paragrafo&Modulo=Modulo") 
   ' Response.Redirect("inserisci_test.asp") 
   errore=2
  
   end if
   
 if (errore=0) then
   
   ' devo vedere se il setting � tale da richiedere voto=1 come default oppure no  
    QuerySQL1="Select * from Setting"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
	Valutato=rsTabella.fields("Valutato") 
	rsTabella.close
	if Valutato=1 then
        Voto=1 ' valore di default 
	else
	    Voto=0
	end if
	
	if daSimulazione<>"" then ' aggiorno 
 
 
   Strada_KO = ucase(Request.Form("txtStradaKO"))
   Strada_KO = Replace(Strada_KO, Chr(34), "'")
   Strada_KO=  Replace(Strada_KO,"'",Chr(96))
   
   
   
   Strada_OK = ucase(Request.Form("txtStradaOK"))
   Strada_OK = Replace(Strada_OK, Chr(34), "'")
   Strada_OK=  Replace(Strada_OK,"'",Chr(96))
   
 
 
 QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK &  "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata &"', Distanza= '" & Distanza & "'  WHERE CodiceMetafora =" &CodiceMetafora&";"
' response.write(QuerySQL)
			   
			    randomize()
				rand=rnd()
				rand=cint(left(rand*1000,1))
				' per evitare che la pagina venga memorizzata in cache e quindi non aggiornata, il parametro dice che � una pagina diversa
  ConnessioneDB.Execute QuerySQL 
 response.Redirect "6_simula_metafora_topolino.asp?Cartella="& Session("Cartella")&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&CodiceMetafora&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo&"&Modulo="&Modulo&"&nocache="&rand
				
 else	
  QuerySQL="INSERT INTO M_Topolino (Topolino, Formaggio, Fame,Labirinto,Strada,Strada_OK,Strada_KO,Testata,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ThreadParent) SELECT '" & Topolino & "','" & Formaggio & "', '" & Fame & "','" & Labirinto & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Testata & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4)& "',"& ThreadParent &";" 
 '  end if 
  ConnessioneDB.Execute QuerySQL 
 end if 
   

 if (daSimulazione<>1) or  (daSviluppa=1)  then
  
    QuerySQL = "SELECT CodiceMetafora,Cartella FROM M_Topolino WHERE CodiceMetafora=(Select Max(CodiceMetafora) FROM M_Topolino);" 
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    ID=rsTabella(0)
	
	 
     Session("CodiceMetafora")=ID
	 Session("Capitolo")=Capitolo
	 Session("Paragrafo")=Paragrafo
	 Session("CodiceTest")=CodiceTest

	 
    CARTA=rsTabella(1)
	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" 'per il server on line
     
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DELLA METAFORA
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
	 
	url=Replace(url,"\","/")
	  
		
					'url2="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logTopo.txt"
	'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
	'				objCreatedFile.WriteLine(url)
	'				objCreatedFile.Close
	
	'response.write(url)

  if instr(Sintesi,"<script>")<>0 then
	   Sintesi=Replace(Sintesi,"<script>","")
	   Sintesi=Replace(Sintesi,"</script>","")
	end if
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
	objCreatedFile.WriteLine(Sintesi)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
	'response.write(url)
	
	'if Tipo="1" then 'CREAZIONE FILE DI TESTO PER INSERIRE LA DOMANDA
	'
	'	url4=Replace(url4,"\","/")
	'	 
	'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
	'	' Write a line with a newline character.
	'	objCreatedFile.WriteLine(Domanda)
	'	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	'	objCreatedFile.Close
	'end if 
	'response.write("<br>" & url)
	
	'On Error Resume Next
	
If Err.Number = 0 Then
	    'https://www.umanet.net/expo2015/UECDL/script/quaderno_metafore.asp?
	session("inserita")=true
	Response.Write "Inserimento avvenuto! "
	
	Else
	Response.Write Err.Description 
	Err.Number = 0
	End If
	
	



   %>
	</font>   
	 
		
      <h4><a href="inserisci_metafora_topolino.asp?Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>">Continua ...</a></h4>

<%end if ' chiudo if daSimulazione<>1 then%>

	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			<div class="contenuti_test">
			 <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
              <h3 class="sottotitolo"><a href="../cClasse/quaderno_metafore.asp?id_classe=<%=Session("Id_Classe")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Vai al Quaderno </a></h3>
 

			<a href="#" onClick="history.go(-1);return false;" class="sottotitolo">Indietro</a>
             <%' response.Redirect request.serverVariables("HTTP_REFERER") ' torno indietro direttamente senza chiedere%>
<%else
  if (errore=1) then
     'response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
  end if 
  if (errore=2) then
    response.write("Controlla che non ci siano campi lasciati vuoti")
  end if %>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
  
end if 			


if daTopolino<>"" then
QuerySQL ="UPDATE M_Topolino SET Pf="&ID&" WHERE CodiceMetafora =" &CodiceMetafora&";"
response.write(QuerySQL& "<br>")
ConnessioneDB.Execute QuerySQL 

QuerySQL ="UPDATE M_Topolino SET Pi="&CodiceMetafora&" WHERE CodiceMetafora =" &ID&";"
ConnessioneDB.Execute QuerySQL


' metto anche i link nella tabella link forse serve solo uno dei due modi
LP=Li ' Livello Partenza Strada_OK nella metafora prima, messo IN BASE AI RADIO BOX
LD=2 ' Livello Arrivo Destinazione nella metafora 
T2=""
QuerySQL="INSERT INTO LinkTopolino (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(CodiceMetafora) & "','" &LP & "', '" & CInt(ID) & "','" & LD & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL&"<br>")
' non lo metto perch� non prevedo di tornare indietro ma solo in avanti, altrimenti mi mette blu anche il livello di arrivo del link
'QuerySQL="INSERT INTO LinkNavigazione (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(ID) & "','" &LD & "', '" & CInt(CodiceMetafora) & "','" & LP & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
'ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL)

end if

if daSviluppa<>"" then

     ID=Session("CodiceMetafora")
	 Capitolo=Session("Capitolo")
	 Paragrafo=Session("Paragrafo")
	 CodiceTest=Session("CodiceTest")
	 
	  Session("CodiceMetafora")=""
	 Session("Capitolo")=""
	 Session("Paragrafo")=""
	 Session("CodiceTest")=""
	 
	 
response.Redirect "inserisci_valutazione_metafore.asp?Cartella="& Session("Cartella")&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&ID&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo
else 
Session("Modulo")=Modulo
response.Redirect request.serverVariables("HTTP_REFERER") 
end if

  


%>
            </div>
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	