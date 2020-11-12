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
 ' On Error Resume Next  
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
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet,Data
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Sintesi
   Dim Autista,Carburante,Luogo,Strada,Strada_KO,Strada_OK,Cespugli,Lupo,Cestino,Distanza
   Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  						
  
   'Nome=Request.QueryString("Nome")
   'Cognome=Request.QueryString("Cognome")
   daSimulazione = Request.QueryString("daSimulazione")
   daNavigazione = Request.QueryString("daNavigazione") ' settato se sono chiamato per lo sviluppo della metafora
   daSviluppa = daNavigazione
   CodiceTest = Request.QueryString("CodiceTest")
 
   if Request.QueryString("Li")<>"" then
    Li = cint(Request.QueryString("Li"))
   else
   li=0
   end if
   
  ' Function gira_data()
'  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
'   End Function
'   Data = gira_data()
   
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   
  
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   'CodiceCap=Request.Cookies("Dati")("CodiceCap")
 ' Num=Request.QueryString("Num")
Cartella=Request.QueryString("Cartella")
if Request.QueryString("Num") <>"" then
Num=cint(Request.QueryString("Num"))
else
Num=1
end if
'response.write("NUM="&Num)
CodiceTest= Request.QueryString("CodiceTest")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
Capitolo=Request.QueryString("Capitolo")
CodiceMetafora=Request.QueryString("CodiceMetafora")

DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
 daSviluppa=request.QueryString("daSviluppa")
 
  
   Autista = ucase(Request.Form("txtAutista"))
   Autista = Replace(Autista, Chr(34), "'")
   Autista=  Replace(Autista,"'",Chr(96))
  
   Destinazione = ucase(Request.Form("txtDestinazione"))
   Destinazione = Replace(Destinazione, Chr(34), "'")
   Destinazione=  Replace(Destinazione,"'",Chr(96))

  

   Carburante = ucase(Request.Form("txtCarburante"))
   Carburante = Replace(Carburante, Chr(34), "'")
   Carburante=  Replace(Carburante,"'",Chr(96))

   Luogo = ucase(Request.Form("txtLuogo"))
   Luogo = Replace(Luogo, Chr(34), "'")
   Luogo=  Replace(Luogo,"'",Chr(96))
   Strada = ucase(Request.Form("txtStrada"))
   Strada = Replace(Strada, Chr(34), "'")
   Strada=  Replace(Strada,"'",Chr(96))

   Strada_KO = ucase(Request.Form("txtStrada_ko"))
   Strada_KO = Replace(Strada_KO, Chr(34), "'")
   Strada_KO=  Replace(Strada_KO,"'",Chr(96))
   
   Strada_OK = ucase(Request.Form("txtStrada_ok"))
   Strada_OK = Replace(Strada_OK, Chr(34), "'")
   Strada_OK=  Replace(Strada_OK,"'",Chr(96))
   
   Cespugli = ucase(Request.Form("txtCespugli"))
   Cespugli = Replace(Cespugli, Chr(34), "'")
   Cespugli=  Replace(Cespugli,"'",Chr(96))
   
   Lupo = ucase(Request.Form("txtLupo"))
   Lupo = Replace(Lupo, Chr(34), "'")
   Lupo=  Replace(Lupo,"'",Chr(96))
   
   Cestino = ucase(Request.Form("txtCestino"))
   Cestino = Replace(Cestino, Chr(34), "'")
   Cestino=  Replace(Cestino,"'",Chr(96))
   
   
   Distanza = ucase(Request.Form("txtDistanza"))
   Distanza = Replace(Distanza, Chr(34), "'")
   Distanza=  Replace(Distanza,"'",Chr(96))
   
   Sintesi=ucase(Request.Form("S1"))
   Sintesi= Replace(Sintesi, Chr(34), "'")
   Sintesi=  Replace(Sintesi,Chr(39),Chr(96))

 '  Sintesi=  Replace(Sintesi,"'",Chr(96))

    
   
   
   if ( (len(Autista)=0) or (len(Destinazione)=0) or (len(Carburante)=0) or (len(Luogo)=0) or (len(Strada)=0) or (len(Strada_KO)=0) or(len(Strada_OK)=0) or(len(Cespugli)=0) or(len(Lupo)=0) or(len(Cestino)=0) or(len(Distanza)=0)) then
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
 if daSimulazione=1 then ' aggiorno 
 
 QuerySQL ="UPDATE M_Navigazione SET Autista = '" & Autista & "', Destinazione= '" & Destinazione & "',Carburante= '" & Carburante & "',Luogo= '" & Luogo & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK &  "', Strada_KO = '" & Strada_KO & "', Cespugli = '" & Cespugli &"', Lupo= '" & Lupo & "',Cestino='"&Cestino&"',Distanza='"&Distanza&"' WHERE CodiceMetafora =" &CodiceMetafora&";"
					 
 else	
	
  QuerySQL="INSERT INTO M_Navigazione (Autista, Destinazione, Carburante,Luogo,Strada,Strada_OK,Strada_KO,Cespugli,Lupo,Cestino,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora) SELECT '" & Autista & "','" & Destinazione & "', '" & Carburante & "','" & Luogo & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Cespugli & "','" & Lupo & "','" & Cestino & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) &"';" 
  
 
 '  end if 
end if

  'Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\anno_2012-2013\logMETAF2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close  
'
response.write(QuerySQL)
ConnessioneDB.Execute QuerySQL 

if daSviluppa=1 then
' devo collegare alla metafora precedente il Pf dato dal codice della metafora appena inserit.
 QuerySQL="Select Max(CodiceMetafora) from M_Navigazione;"
  set Rs=ConnessioneDB.Execute(QuerySQL)
  codMax=Rs(0)
  
  QuerySQL ="UPDATE M_Navigazione SET Pf = '" & codMax & "' WHERE CodiceMetafora =" &CodiceMetafora&";"
  ConnessioneDB.Execute QuerySQL			
 end if 
 
 if daSimulazione<>1 then ' cio� nullo =""  quindi non simulazione(aggiornamento) ma nuovo inserimento
 ' prelevo ID della metafora appena inserita 
  
    QuerySQL = "SELECT CodiceMetafora,Cartella FROM M_Navigazione WHERE CodiceMetafora=(Select Max(CodiceMetafora) FROM M_Navigazione);" 
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    ID=rsTabella(0)
    CARTA=rsTabella(1)
    Session("CodiceMetafora")=ID
 ' per ritornare il valore al chiamante per visualizzare alert verde con link diretto a metafora 
	 Session("Capitolo")=Capitolo
	 Session("Paragrafo")=Paragrafo
	 Session("CodiceTest")=CodiceTest
	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" 'per il server on line

	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DEL NODO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	url3=Replace(url,"\","/")
	url=url3
	'response.write(url)
  
  if instr(Sintesi,"<script>")<>0 then
	   Sintesi=Replace(Sintesi,"<script>","")
	   Sintesi=Replace(Sintesi,"</script>","")
	end if
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	objCreatedFile.WriteLine(Sintesi)
	objCreatedFile.Close
	'response.write(url)

	'if Tipo="1" then 'CREAZIONE FILE DI TESTO PER INSERIRE LA DOMANDA
	'	url4=Replace(url4,"\","/")
	'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
	'	' Write a line with a newline character.
	'	objCreatedFile.WriteLine(Domanda)
	'	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	'	objCreatedFile.Close
	'end if 
	'response.write("<br>" & url)

		'On Error Resume Next
		If Err.Number = 0 Then
		session("inserita")=true
		
		Response.Write "Inserimento avvenuto! "
		'response.Redirect request.serverVariables("HTTP_REFERER") 
		Else
		Response.Write Err.Description 
		Err.Number = 0
		End If
		
	



   %>
	</font>   
	 
		
      <h4><a href="inserisci_metafora_patente.asp?Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>">Continua ...</a></h4>
<%end if ' chiudo if daSimulazione<>1 then%>
	
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			<div class="contenuti_test">
			<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3 class="sottotitolo">
<a href="inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&DATA=<%=DataTest%>&Cartella=<%=Cartella%>&cod=<%=CodiceAllievo%>&CodiceTest=<%=CodiceTest%>&CodiceMetafora=<%=ID%>&CodiceAllievo=<%=CodiceAllievo%>&Capitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Paragrafo=<%=Paragrafo%>&Autista=<%=Autista%>&Destinazione=<%=Destinazione%> &Carburante=<%=Carburante%>&Luogo=<%=Luogo%>&Strada=<%=Strada%>&Strada_OK=<%=Strada_OK%>&Strada_KO=<%=Strada_KO%>&Cespugli=<%=Cespugli%>&Cestino=<%=Cestino%>&Lupo=<%=Lupo%>&Distanza=<%=Distanza%>&MO=<%=Modulo%>&VAL=<%=Voto%>&Pippo=1 ">Vai a Simulazione e Narrazione</a></h3>
 <h3 class="sottotitolo"><a href="../cClasse/studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Vai al Quaderno </a></h3>
<h3 class="sottotitolo"><a href="#" class="sottotitolo" onClick="history.go(-1);return false;">Torna Indietro</a></h3>

<h3 class="sottotitolo"><a href="../../U-ECDL/home_uecdl_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 
            
<%else
  if (errore=1) then
     'response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
  end if 
  if (errore=2) then
    response.write("Controlla che non ci siano campi lasciati vuoti")
  end if %>
	<a class="sottotitolo" href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
end if 			

' se sonochiamata da sviluppa devo aggiornare il puntatore della metafora precedente con quello della nuova
if daNavigazione<>"" then
QuerySQL ="UPDATE M_Navigazione SET Pf="&ID&" WHERE CodiceMetafora =" &CodiceMetafora&";"
ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL& "<br>")
QuerySQL ="UPDATE M_Navigazione SET Pi="&CodiceMetafora&" WHERE CodiceMetafora =" &ID&";"
ConnessioneDB.Execute QuerySQL

'??? DA CAPIRE MEGLIO VALUTARE LINK n-m anziche lista fare albero !!!
' metto anche i link nella tabella link forse serve solo uno dei due modi
LP=Li ' Livello Partenza Strada_OK nella metafora prima, messo IN BASE AI RADIO BOX
LD=2 ' Livello Arrivo Destinazione nella metafora 
T2=""
QuerySQL="INSERT INTO LinkNavigazione (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(CodiceMetafora) & "','" &LP & "', '" & CInt(ID) & "','" & LD & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
ConnessioneDB.Execute QuerySQL 


 				randomize()
				rand=rnd()
				rand=cint(left(rand*1000,1))
				' per evitare che la pagina venga memorizzata in cache e quindi non aggiornata, il parametro dice che � una pagina diversa
  				 
 response.Redirect "6_simula_metafora_navigazione.asp?CodiceMetafora="&CodiceMetafora&"&nocache="&rand
'response.write(QuerySQL&"<br>")
' non lo metto perch� non prevedo di tornare indietro ma solo in avanti, altrimenti mi mette blu anche il livello di arrivo del link
'QuerySQL="INSERT INTO LinkNavigazione (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(ID) & "','" &LD & "', '" & CInt(CodiceMetafora) & "','" & LP & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
'ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL)

end if
'Response.Redirect "quaderno_metafore.asp?id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&Session("CodiceAllievo")&"&DataClaq="&Session("DataClaq")&"&DataClaq2="&Session("DataClaq")&"&damenu=1"
	  
if daSviluppa<>"" then

     ID=Session("CodiceMetafora")
	 Capitolo=Session("Capitolo")
	 Paragrafo=Session("Paragrafo")
	 CodiceTest=Session("CodiceTest")
	 
	  Session("CodiceMetafora")=""
	 Session("Capitolo")=""
	 Session("Paragrafo")=""
	 Session("CodiceTest")=""
	 
	 
response.Redirect "inserisci_valutazione_metafore.asp?Cartella="& Session("Cartella")&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&ID&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo&"&Modulo="&Modulo
else 
'response.Redirect request.serverVariables("HTTP_REFERER") 
 response.Redirect "inserisci_valutazione_metafore.asp?damodifica=1&Cartella="& Session("Cartella")&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&ID&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo

end if

%>
           </div>
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	