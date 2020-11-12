 
<%@ Language=VBScript %>
   
<%divid_apro=request.QueryString("divid_apro")
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario

		
	'	url="C:\Inetpub\umanetroot\Anno_2012-2013\log_session.txt"
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				 
'				objCreatedFile.WriteLine("CA)" &Session("CodiceAllievo"))
'				objCreatedFile.WriteLine("Cla)" &Session("Id_Classe"))
'				objCreatedFile.Close
		    
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
%>
<html>
<head> 

	<script src="../lib/prototype.js" type="text/javascript"></script> 
    <script src="../src/scriptaculous.js" type="text/javascript"></script> 
<script src="../src/unittest.js" type="text/javascript"></script>
    <script src="../../../SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript"> 
function SessioneScaduta(){  
            window.alert("Sessione  scaduta, effettua nuovamente il Login!");
             location.href="../home.asp";
             }
</script>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Studenti </title>
<link rel="stylesheet" type="text/css" href="../../stile.css">
 <style>
	
	 2
 /*
 div#header, div#content { 
    BORDER-RIGHT: #29abe2 1px solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #29abe2 1px solid;
    PADDING-LEFT: 5px;
    FONT-SIZE: 10px;
    PADDING-BOTTOM: 10px;
    margin: 0 auto 0 auto;
    BORDER-LEFT: #29abe2 1px solid;
    WIDTH: 75%;
   
    BORDER-BOTTOM: #29abe2 1px solid;
    FONT-FAMILY: Verdana, Arial;
    TEXT-ALIGN: left;
	
	border-radius: 8px;
	-moz-border-radius: 8px;
    -webkit-border-radius: 8px;
	
	-webkit-box-shadow : 0 0px 10px #999;
	-moz-box-shadow : 0px 0px 10px #999;}*/  
 
 
 #superheader { 
 
 position:fixed; 	
    z-index:1; 
    width:70%;
    height:3em;
     background: #fff; 
    
     top:0;
	  
	  text-align:center;
	  
   overflow:visible;
 }  
  

 
	</style>

<link href="../../../SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css">

</head>
 
<!-- #include file = "../extra/test_server.asp" --> 
<!-- #include file = "../include/formattaDataCla.inc" --> 
<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
<!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
<!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->
<!-- #include file = "../var_globali.inc" -->

	           

<%
' non si capisce per quale cazzo di motivo non funziona il controlla sessione come in home_app.asp ??????
' non va il window.alert dentro if , se lo metto fuori funziona

if Session("CodiceAllievo")="" or Session("Id_Classe")="" then %>
                  		 
		 
				<script language="javascript" type="text/javascript"> 
				    window.alert("Sessione  scaduta, effettua nuovamente il Login!");
                    location.href="../home.asp";
				</script>
				<%
				response.Redirect "../home.asp"
				 
				 %>
 
<% end if%>
  

 <%
    
	'SetLocale(1040)  ' imposta il formato data corretto
    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	ScalaValutaz=rsTabellaCI("ScalaValutaz")
	rsTabellaCI.close
 Dim esecuzione
 set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito


%>




<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="Effect.toggle('dAvvisi','appear'); return false;">
-->
<%
dividApri=request.querystring("dividApri")
if (CIAbilitato=0) and Session("Admin")=False then  ' else è alla fine disabilito copia incolla %>
<body  oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="Effect.toggle('dAttività','appear');Effect.toggle('dAvvisi','appear'); return false;">  
<%else%>
<body  onLoad="Effect.toggle('dAttività','appear');Effect.toggle('dAvvisi','appear'); return false;"> 
 
<%end if%>
<!-- #include file = "../service/controllo_sessione.asp" -->
 
<% 
Dim periodi() ' vettore delle date per il calcolo della classifica, più avanti farò il redim
Dim vetstud(35) ' massimo numero di studenti possibile
vetstud(0)="?"
'divid(0)="terzacom"
'Response.Buffer=True
 
  On Error Resume Next
xEstrazione=request.querystring("xEstrazione")
id_classe=request.querystring("id_classe")
classe=request.querystring("classe")
divid=request.querystring("divid")
if divid="" then divid=Session("divid")
divid2=request.querystring("divid")

PS=request.querystring("PS") ' vale 1 se devo mostrare anche i Punti Social chiamato da javasscript
if PS="" then ' per la prima chiamata mostrio i PS
   PS=1
end if
cod=Request.QueryString("cod")
daStud=Request.QueryString("daStud")
daMenu=Request.QueryString("daMenu")
DataCla=request.form("txtData") 
DataCla2=request.form("txtData2")
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")
if daMenu<>"" then
    DataCla=request.QueryString("DataClaq") 
    DataCla2=request.QueryString("DataClaq2")
end if
if daStud<>"" then
   DataClaq= DataCla
   DataClaq2=DataCla2
end if



Session("DataClaq")=DataClaq
Session("DataClaq2")=DataClaq2
 

'response.write("<br>Datacla="&DataCla)
'response.write("<br>Datacla2="&DataCla2)
'response.write("<br>Dataclaq="&DataClaq)
'response.write("<br>Dataclaq2="&DataClaq2)

' aggiungo per IIS7 ?
'if DataCla="" then
'  DataCla=DataClaDefault
'end if
'if DataCla2="" then
'  DataCla2=DataCla2Default
'end if
' se è la prima chiamata il valore del form sopra la classifica è nullo
if (DataCla<>"") and (DataCla2<>"") then
	Session("DataCla")=DataCla
	Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
else
   Session("DataCla")= Session("DataClaq")
   Session("DataCla2")= Session("DataClaq2")
end if



				'dim objFSO,objCreatedFile
				'Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'Dim sRead, sReadLine, sReadAll, objTextFile
			'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\Anno_2012-2013\log_id.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="cla=" &cla &  " d="&d& " id_classe="&id_classe & " DataCla="& DataCla
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
'



' raccolgo il parametro data per clacolarela classifica a scalare quando la pagina viene richiamata dal form che setta la data
 'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query
 Dim ConnessioneDB,ConnessioneDB1, rsTabella,rsTabella1,rsTabella2,rsTabella0,rsTabella3,rsTabella4,rsTabellaForum, QuerySQL,QuerySQL1 ,CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,DataCla
 'StringaConnessione= Response.Cookies("Dati")("StrConn")

 ' per il file di log
 dim objFSO,objCreatedFile
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim sRead, sReadLine, sReadAll, objTextFile
 Set objFSO = CreateObject("Scripting.FileSystemObject")  
 
 %>
<!-- <center>-->
 


<div id="bloc_sinistra">
	<!--<div id="bloc_sinistra_int" style="margin-top:35px;margin-left:-5px;">-->
		<div id="bloc_sinistra_int">
         <div id="bloc_sinistra_cont">	
            
                    <div id="logo_space">
                        <div class="menu_title"><div id="home_page">
                            <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx">
                            </div></div>
                            <div class="menu_cont_one"><div id="comune"><b>
                                <a href="../../home.asp"><font color=#000000>HOME PAGE</font></a></b></div></div>
                            <div class="menu_cont_two">
                                <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx">
                      </div>
      </div>
                        <div id="logo_space1">
                        <p align="center">
                        <%
    
    ' HO CREATO UNA CLASSE
    ' per sapere se sono in esecuzione sul server o in locale, serve per distinguere gli url per le risorse
     ' pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
    '  if (left(pathEnd1,10)="c:\inetpub") then
    '     locale=1
    '  else
    '     locale=0
    '  end if 	
    ' 	
                
                     if esecuzione.locale=1 then%>
                         <a href="../../../U-ECDL/UECDL/index.html" target="_blank"><img src="../../img/umanet2.png" width="90%" ></a></div>
                        
                         <% else%>
                             <a href="https://www.umanet.net/informistica/UWWW/Benvenuto.html" target="_blank"><img src="../../img/umanet2.png" width="90%" ></a></div>
                        
                         <% end if %>
                        
                         <% 
                
                    
    
           
                     
                       
                       'Apertura della connessione al database
                       Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>   
                       <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
                        
                        <%
                            id_classe=Session("Id_Classe")
                            QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
                            
                            
                        'dim objFSO,objCreatedFile
    '				Const ForReading = 1, ForWriting = 2, ForAppending = 8
    '				Dim sRead, sReadLine, sReadAll, objTextFile
    '				Set objFSO = CreateObject("Scripting.FileSystemObject")
    '				url="C:\Inetpub\umanetroot\anno_2012-2013_2\log.txt"
    '				Set objCreatedFile = objFSO.CreateTextFile(url, True)
    '				objCreatedFile.WriteLine(QuerySQL) 
    '				objCreatedFile.Close
                        
                        
                    
                        Set rsTabella = ConnessioneDB.Execute(QuerySQL)
                        'divid=request.querystring("divid")
                        cartella=rsTabella.fields("Cartella")%>
                        
                            <div class="menu_sinistra">
                                    
                                  <div class="menu_title"><div id="<%=divid%>"><%=rsTabella.fields("Classe")%></div>
                              </div>
                                    <div class="menu_cont_one">
                                       <a href="../lavagna/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Lavagna&nbsp;</a>
                                    </div>
                                    <div class="menu_cont_two"  >
                                        <a   href="home_app.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>">Apprendimento</a>
                                    </div>	
                                        
                                    <div class="menu_cont_one"  >
                                        <a href="../../home_ver.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>&cartella=<%=rsTabella.fields("cartella")%>">Verifica</a> 
                                    </div>	
                                    <div class="menu_cont_two"  >
                                        <a href="../forum/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Forum&nbsp;</a> 
                                    </div>	
                                    <div class="menu_cont_one"  >
                                        <a href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">&nbsp;Chat</a>
                              </div>
                                        
                                         <div class="menu_cont_two"  >
                                          			<a href="../diario/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Diario</a> </div>
                                        <div class="menu_cont_one"  >
						
                                     <a class="menu_selected" href="studente_domande.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>">Classe</a></div>
                                       
                                                                     
      </div>	
                             
                            </p>
                            
                            
                            
                            
                            <div class="menu_sinistra">
                            <div class="menu_title"><div id="quintacom">U-ECDL</div></div>
                            <div class="menu_cont_one">
                                <a href="home_uecdl_app.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>">Apprendimento</a></div>
                            <div class="menu_cont_two">
                                <a href="../../U-ECDL/home_uecdl_ver.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>">Verifica</a></div>	
                    </div>
                    <div class="menu_sinistra">
                            <div class="menu_title"><div id="quarta">GESTIONE</div></div>
                            <div class="menu_cont_one">
                                <a href="../service/logout.asp">Logout</a></div>
                                
                            <%if (session("Admin")=true) then %>
                            <div class="menu_cont_two">
                            <a href="studente_domande_gruppi.asp">Gruppi</a>
                            </div>
                            
                        
                             <div class="menu_cont_one">
                            <a href="../cAdmin/admin.asp?Id_Classe=<%=id_classe%>&divid=<%=divid%>">Admin</a>
                            </div>
                             
                            <%end if %>
                    </div>
                    
                        <%
                        rsTabella.Close()
                        Set rsTabella = nothing
                    
                    
                         
                        QuerySQL="SELECT Classi_Moduli_Paragrafi.Id_Classe, Moduli.Titolo, Paragrafi.Titolo, Moduli.ID_Mod, Paragrafi.ID_Paragrafo,Moduli.Cartella,Moduli.URL,Moduli.URL_OL,Classi.Classe,Paragrafi.URL_L,Paragrafi.URL_O,Moduli.Posizione from MODULI_NOT_UMANET where Id_Classe='"&id_classe&"';"
                        Set rsTabella = ConnessioneDB.Execute(QuerySQL)
                        
                        %>									
    
                         
  </div>
</div>
        </div>
        </div>
</div>
 



<div id="bloc_destra">
		<div id="bloc_destra_int">
			<div id="bloc_destra_cont">
            
 
<!--Visualizza i report sopra la CLASSIFICA -->

<!-- #include file = "studente_domande_include/1_report.asp" --> 
     
          
     
	<br><b>
	<% 'if (request.form("txtData")<>"") then response.write("Caklcola data") end if 
	   response.Write("Classifica al " &day(date())&"/"&month(date())&"/"&year(date()))%> <br><br></b>

<!-- #include file = "studente_domande_include/1_periodi.asp" --> 
<input type="button"  style="width:60px;height:25px;" value="Invia" name="B1" onClick="aggiorna()"> 
 <input type="checkbox"  name="cbPS" value="1" checked="true" title="Deseleziona per escludere i Punti Social dalla classifica">  <b> 
	Includi PS
 </p> 
</form>

	<form method="POST" form action="aggiorna_punteggio.asp?classe=<%=classe%>&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
	 

<!-- #include file = "studente_domande_include/1_classifica.asp" --> 


	
		 <!--<table align=center border=1 bordercolor=pink>-->
	 <table id="zebra_stud" summary="Classifica studenti" align=center border=1>
	 
    <thead>
	<tr><th  title="Posizione"><b>N.</b></th><th><center><b>Cognome Nome</b><th title="Totale"><b>TOT</b></th><th title="Punti Domande"><b>PD</b></th><th title="Punti Nodi"><b>PN</b></th><th title="Punti Frasi"><b>PF</b></th><th title="Punti Metafore"><b>PM</b></th><th><b title="Punti Crediti">PC</b></th><th><b title="Punti Social Network">PS</b></th><th title="Voto Virtuale"><b>VV</b></th><th title="Percentuale rispetto al massimo"><b>%</b></th></tr>
	</thead>
	
	<%
	'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logfrasi.txt"
	'Set objCreatedFile = objFSO.CreateTextFile(url, True)
	 i=0
	 do while not rsTabella.eof 
	 CodiceAllievo=rsTabella("CodiceAllievo")
	
	
	'
'	QuerySQL="INSERT INTO Domande (Quesito, Risposta1, Risposta2,Risposta3,Risposta4,RispostaEsatta,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Domanda & "','" & R1 & "', '" & R2 & "','" & R3 & "','" & R4 & "','" & RE & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella &  "','"& In_Quiz &"';"
'		ConnessioneDB.Execute QuerySQL 
''		
''		
'		    QuerySQL="INSERT INTO Nodi (Chi, Cosa, Dove,Quando,Come,Perche,Quindi,Id_Stud,Id_Arg,Id_Mod,Data,Cartella) SELECT '" & Chi & "','" & Cosa & "', '" & Dove & "','" & Quando & "','" & Come & "','" & Perche & "','" & Quindi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Cartella & "';"
'' 
'   ConnessioneDB.Execute QuerySQL 
''		   
'QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & Cartella & "','" & In_Quiz & "';"
'' 
'   ConnessioneDB.Execute QuerySQL 
''	
'	
'	'response.write(QuerySQL)
'			'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logfrasi.txt"
'				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
'	'			objCreatedFile.WriteLine(QuerySQL)
'				'objCreatedFile.Close
'	
'	
	'QuerySQL="INSERT INTO M_Topolino (Topolino,Id_Stud,Id_Arg,Id_Mod,Data,Voto,In_Quiz) SELECT '" & Topolino & "','" &CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & In_Quiz & "';"
'   ConnessioneDB.Execute QuerySQL 
'	
'	QuerySQL="INSERT INTO M_Navigazione (Autista,Id_Stud,Id_Arg,Id_Mod,Data,Voto,In_Quiz) SELECT '" &Autista & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & In_Quiz & "';"
'' 
'   ConnessioneDB.Execute QuerySQL 
''	
'QuerySQL="INSERT INTO M_Desideri (SoggettoC, DomandaC, MotivazioneC,DesiderioC,BisognoC,SoggettoS,RispostaS,MotivazioneS,DesiderioS,BisognoS,TipoEvento,TolleranzaC,URL_Teoria,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora) SELECT '" & SoggettoC & "','" & DomandaC & "', '" & MotivazioneC & "','" & DesiderioC & "','" & BisognoC & "','" & SoggettoS & "','" & RispostaS & "','" & MotivazioneS & "','" & DesiderioS & "','" & BisognoS  & "'," & TipoEvento & "," & TolleranzaC &",'"  & URL_teoria &"','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) &"';" 

	
	if (rsTabella(5)=0) then
	     rsTabella(5)=1
	end if
	if (i mod 2) = 0  then 
	    classe_riga="zebra-dispari"
	else
	    classe_riga=""
	end if 
	 
		%>
	   
 
        
         
				<tr class="<%=classe_riga%>"><td><%=i+1%></td><td><a href="../?divid=<%=divid%>&DataClaq=<%=DataCla%>&DataClaq2=<%=DataCla2%>&id_classe=<%=id_classe%>&classe=<%=classe%>&cod=<%=rsTabella("CodiceAllievo")%>">   <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%>  </a></td><td><%=rsTabella(5)%>  </td><td><%=rsTabella(0)%></td><td><%=rsTabella(6)%></td><td><%=rsTabella(7)%></td><td><%=rsTabella(8)%></td><td><%=rsTabella(4)%></td><td><%=(rsTabella("PuntiForum")+rsTabella("PuntiDiario"))%></td><td><%=fix((rsTabella(5)*ScalaValutaz/max) * 10) / 10 %></td><td><%=round(fix((rsTabella(5)*100/max) * 10) / 10 ) %></td></tr>
      <%      
		' aggiungo al vettore che servirà per estrarre a sorte per l'orale
		

i=i+1
vetstud(i)=rsTabella.fields("Cognome")
rsTabella.movenext
loop
' numero studenti per quella classe
NumStud=i





rsTabella.close()
if (DataCla<>"") then 
' 'dopo aver caricato la classifica cancello le tabelle create
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_DOMANDE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_FRASI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_NODI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MTOPOLINO")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MNAVIGAZIONE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MDESIDERI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_METAFORE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_CREDITI")
    ConnessioneDB.Execute("Delete * From TPUNTEGGI_STUDENTI_FORUM")
	ConnessioneDB.Execute("Delete * From TPUNTEGGI_STUDENTI_DIARIO")

	'response.write("Cancellate")
end if%> 



</table> 
<br>


<a target="_new" href="../cGrafici/genera_grafico.asp?byGrafico=1&PS=<%=PS%>&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>&indice_periodo=<%=indice_periodo%>&indice_periodo2=<%=indice_periodo2%>">Visualizza grafico</a><br></p>
<center><a target="_blank" href="../cAdmin/consulta_profili_new.asp?id_classe=<%=id_classe%>&divid=<%=divid%>">Visualizza Classe </a></center>  


<%
'else
' PRELEVO IN ANTICIPO IL CONGOME NOME NEL CASO LA QUERY 2 NON TROVI NULLA IN QUEL PERIODO E QUINDI RESTITUISCA NULL

QuerySQL="SELECT Allievi.Cognome,Allievi.Nome " &_
" FROM Allievi INNER JOIN Domande ON Allievi.CodiceAllievo=Domande.Id_Stud" &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

 'url="C:\Inetpub\umanetroot\anno_2012-2013\logCod.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
' 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
cognome = rsTabella.fields("Cognome")
nome = rsTabella.fields("Nome")  



%>
 
 
<% i=0 ' serve per decidere quando aggiungere la riga con il modulo
 divid=0
%>
<br>

  <div  align="center">
<div class="hr" style="width:9%;"><hr /></div><br>
 </div>                         
<b> <span class="sottotitoloquaderno">  QUADERNO DI   <%=ucase(cognome)%>
<%=ucase(nome)%> 
</B></span> 
 
	 
 <b>
	<% 'if (request.form("txtData")<>"") then response.write("Caklcola data") end if 
	   'response.Write("Classifica al " &day(date())&"/"&month(date())&"/"&year(date()))%> <br><br></b>
<!-- #include file = "studente_domande_include/1_periodi.asp" --> 
<input type="button"  style="width:60px;height:25px;" value="Invia" name="B1" onClick="aggiornaStud()"> 
 <input type="checkbox"  name="cbPS" value="1" checked="true" title="Deseleziona per escludere i Punti Social dalla classifica">  <b> 
	Includi PS
 </p> 
</form>



<%
'response.write("Default LCID is: " & Session.LCID & "<br>")
'response.write("Date format is: " & date() & "<br>")
'response.write("Date forma2t is: " &FormatDateTime(now(),2) & "<br>") 
%>

<!--<div id="superheader"> per tenere fermo il menu styudebnte : difficile gestire il bloc_destra_cont nei diversi utilizzi pagina-->
   <table id="zebra_stud" align=center style="padding-right:0px; border:none;"  width="95%">
<tr><td style="padding-right:140px; background-color:#F5FAFC; border:none;" align="left"> <div>
<ul id="MenuBar1" class="MenuBarHorizontal">
   <li><a   href="../cUtenti/form_cambia_pwd.asp?CodiceAllievo=<%=cod%>&id_classe=<%=id_classe%>&classe=<%=classe%>"><img src="../lavagna/img/icon_profile.gif" width="12" height="12">&nbsp;&nbsp;Profilo</a>
  </li>
   <li><a class="MenuBarHorizontal" href="../forum/default.asp?bacheca=<%=cod%>&nome=<%=nome%>&cognome=<%=cognome%>&id_classe=<%=id_classe%>&divid=<%=Session("divid")%>&cartella=<%=cartella%>">
   <img src="../lavagna/img/facebook1.jpg" width="12" height="13">&nbsp;&nbsp;Bacheca</a>
      
   </li>
  <li><a class="MenuBarItemSubmenu" href="#"><img src="../lavagna/img/icon_aim.gif" width="12" height="12">&nbsp;&nbsp;Attività</a>
  
     <ul>
    <!--   <li><a href="#" onclick="Effect.toggle('dAttività','slide');Effect.toggle('dAvvisi','slide'); return false;">Avvisi</a></li>-->
         <li><a href="#ancora_avvisi"><img src="../lavagna/img/facebook1.jpg" width="18" height="19">&nbsp;&nbsp;Avvisi</a></li>
        <li><a href="#ancora_forum"><img src="../lavagna/img/facebook3_msg.jpg" width="18" height="19">&nbsp;&nbsp;Forum</a> 
    <!--   <li><a href="#forum" onclick="Effect.toggle('forum','slide'); return false;">Forum</a>-->
       <li><a href="#ancora_crediti"> <img src="../lavagna/img/icon_star_red.gif" width="13" height="12">&nbsp;&nbsp;&nbsp;Crediti</a></li>
       <li><a href="#ancora_quiz">Quiz</a></li>
       <li><a href="#ancora_video">Video</a></li>
       
     </ul>
  </li>
  
   <li><a class="MenuBarItemSubmenu" href="#"> <img src="../lavagna/img/icon_pencil.gif" width="12" height="12">&nbsp;&nbsp;Compiti</a>
     <ul>

           <li><a href="#ancora_domande">Domande</a></li>
           <li><a href="#ancora_nodi">Nodi</a> 
    <!--   <li><a href="#forum" onclick="Effect.toggle('forum','slide'); return false;">Forum</a>-->
           <li><a href="#ancora_frasi">Frasi</a></li>
           <li><a href="#ancora_metafore">Metafore</a></li>
           <li><a href="#ancora_immaginario">Immaginario</a></li>
           
         
     </ul>
  </li>
   
   
 </ul>
</div>
</td></tr>
</table>
<p></p>
<br /><br /><br> 




<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND  style="width:15%;"> 
<!--<a href="#" onClick="Effect.toggle('dAttività','appear'); return false;">-->
<span style="font-style:normal;" class="sottotitoloquaderno">&nbsp;&nbsp;ATTIVITA'</span></a> 
</legend>

<div id="dAttività" style="display:none;">


<% 
 ' carico i  messaggi personali su Copiaditestonline
%>
<!-- #include file = "studente_domande_include/2_messaggi.asp" --> 

 <br>
 
<% ' logica per mostrare le ATTIVITA dello studente nel forum%>
 
 
<!-- #include file = "studente_domande_include/2_forum.asp" --> 


<% ' logica per mostrare i CREDITI dello studente nelle varie attività%>
    
	
<!-- #include file = "studente_domande_include/2_crediti.asp" --> 

 
<br>&nbsp

 
<% ' logica per mostrare la cronologia delle classifiche%>


 <!-- #include file = "studente_domande_include/2_cronologia.asp" --> 



<% ' logica per mostrare i risultati nei quiz dello studente relativi ai singoli paragrafi e moduli%>
    


 <!-- #include file = "studente_domande_include/2_quiz.asp" --> 


<% ' logica per mostrare le visualizzazioni dello studente relativi ai singoli paragrafi%>

<!-- #include file = "studente_domande_include/2_visualizzazioni.asp" --> 


</fieldset>
<br>
</div>




<%' prelevo l'elenco delle domande dello studente%>

<!-- #include file = "studente_domande_include/2_domande.asp" --> 

<%

' prelevo l'elenco dei nodi dello studente %>


<!-- #include file = "studente_domande_include/2_nodi.asp" --> 


<%
' RIPETO LA STESSA LOGICA PER L?ELENVCO delle frasi
%>

<!-- #include file = "studente_domande_include/2_frasi.asp" --> 


<%
' RIPETO LA STESSA LOGICA PER L?ELENVCO METAFORE%>


<!-- #include file = "studente_domande_include/2_metafore.asp" --> 


<!-- fine menu metafore  -->
<br>

<%
' RIPETO LA STESSA LOGICA PER L?ELENCO DELLE IMMAGINI IN IMMAGINARIO 
%>


<!-- #include file = "studente_domande_include/2_immaginario.asp" --> 



<%'End if

'End if
'rsTabella.Close()
'Set rsTabella = nothing%>
<br>
<!--</FIELDSET> --><!-- chiudo il quaderno dello studente -->

<!-- #include file = "studente_domande_include/2_altro.asp" --> 

</center>
<script type="text/javascript">
 
function aggiorna() {
	 
		with (document.dati) { 
		 
		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=divid%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1";
		 else
		   document.dati.action = "?divid=<%=divid%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0";
	
	    }
		document.dati.submit();		
}
 
 function aggiornaStud() {
	 
		with (document.dati) { 
		 
		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&daStud=1";
		 else
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&daStud=1";
	
	    }
		document.dati.submit();		
}

  
 
 </script>
 
 <script language="javascript" type="text/javascript">
function cancella_avviso() {
	
	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {  
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();	
	 }
}
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", {imgDown:"../../SpryAssets/SpryMenuBarDownHover.gif", imgRight:"../../SpryAssets/SpryMenuBarRightHover.gif"});
 </script>
 
</html>




