<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
    <meta charset="UTF-8">
   <title>Valutazione domanda</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	
<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
	
	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

  


   <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
</head>


<% Response.Buffer = true
 ' On Error Resume Next  
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
   <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

  <% end if %>
  
  <%   
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   Cartella=Request.QueryString("Cartella") 
    ' DATA=clng(Request.Form("txtDATA"))
   CodiceAllievo = Request.QueryString("cod")
   DATA = Request.Form("txtDATA")
   cla=Request.QueryString("cla")
   id_classe=Request.QueryString("id_classe")
   'CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
Tipodomanda=Request.QueryString("Tipodomanda")
 Multiple=Request.QueryString("Multiple")    
   Domanda1 = Request.Form("txtDomanda")
    VF=Request.QueryString("VF") 
	
	  tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 
 tDom=request.querystring("tDom")
  
 				
    
    %>
	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <!--#include file="../service/gestione_errori.asp" -->
         <!-- #include file = "tabella_corrispondenze.inc" -->
        	  
          
         
	</div>
    
 
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Valuta domanda </h1> 
                    
					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->	 
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>
						<li>
							<a href="#more-login.html">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-files.html">Quaderno</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Domande</a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%> : <%=Paragrafo%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			   <div class="box-content"> 
 <%                    
 
   
   ' sembra che se l'input box è disabilitato il suo valore non viene passato, 
   'questo comporta l'azzeramento della valutazione quando lo studente modifica la sua domanda, 
   'in questo modo la valutazione viene passata come parametro e viene correttamente conservata
   
       voto=clng(Request.QueryString("Voto"))
   
     'INQUIZ=clng(Request.QueryString("INQUIZ"))
	 INQUIZ=clng(Request.Form("txtINQUIZ"))
  
  ' %><br><%
  'response.write("INQUIZbis:"&INQUIZbis)
   Segnalata=Request.Form("txtSegnalata")
    Voto=Request.Form("txtVAL")
	' se non sono admin il campo Voto è disabilitato e restituisce valore "" quindi lo prendo da qudery string
	if Voto="" then
	    Voto=clng(Request.QueryString("Voto"))
	end if	
	VECCHIOVOTO=clng(Request.QueryString("VECCHIOVAL"))
	deltaVoto= Voto-VECCHIOVOTO
	
	
   'voto=clng(Request.QueryString("Voto"))
   Domanda=Replace(Domanda1,"'","''")
   ID=Request.QueryString("CodiceDomanda")

   R11 = Request.Form("txtR1")
   R1=Replace(R11,"'",Chr(96))
    


   R22 = Request.Form("txtR2")
   R2=Replace(R22,"'",Chr(96))

   R33 = Request.Form("txtR3")
   R3=Replace(R33,"'",Chr(96))

   R44 = Request.Form("txtR4")
   R4 = Replace(R44,"'",Chr(96))
  
   Spiegazione=Request.Form("S1")
   TestoDomandaPlus=Request.Form("TestoDomandaPlus")
   errore=0
   
   
   	  
		 function controlla(RisposteEsatte)
		' response.write("cidji")
			 controlla=0
			 i=0
			 esiste=false
			 while (i<=16) and not(esiste)
			' response.write("<br>cidji")
			 '  response.write(v2(i)&"=?"&RisposteEsatte &"<br>")
				if v2(i)= RisposteEsatte then 
				    esiste=true
				   controlla=1
				end if
				i=i+1
			 wend
		 end function
   
   
    if (len(Request.Form("txtRE"))=0)  or (len(Spiegazione)=0)  then 
			   errore=2
	 else
	       RE = clng(Request.Form("txtRE"))
		   errore=0
	 end if 
	' response.write("strcomp(VF,0)="&strcomp(VF,"0"))
if strcomp(VF,"0")<>0 then ' se ì vero o falso devo fare meno controlli
  if  (len(Domanda)=0)  then 
	  errore=2
  
   elseif (RE<>0) and (RE<>1) then ' risposta vero falso 0 o 1
     errore=4
  end if

else  
	 
		  ' if Multiple<>"" then
		 ' response.write("strcomp(Multiple,1)="&strcomp(Multiple,"1") &" RE="&RE)
		  if strcomp(Multiple,"1")=0 then
			   ' controllo validità numero che indica la risposta esatta deve appartenere alla tabella di corrispondenza
			   esiste=controlla(RE)
			   if esiste = 0 then
				  errore = 3
			   end if 	   
			   
		   else
			   if ((RE<1) or (RE>4)) then 
				  errore=1
			   end if
		   end if 
		   'response.write(len(Domanda)&" " &len(R1)&" " &len(R2)&" " &len(R3)&" " &len(R4)&" " &len(Domanda)&" " &len(Spiegazione) & "==="&len(Request.Form("txtRE")))
		   if ( (len(Domanda)=0) or (len(R1)=0) or (len(R2)=0) or (len(R3)=0) or (len(R4)=0) or (len(Spiegazione)=0) ) then 
			   errore=2
		   end if
		 
		   'Domanda1=Domanda
		   'response.write("Domanda="&Domanda1)
		 if Multiple<>"" then
			' se non devo inserire domanda multipla pongo a 0 il campo 
			Multiple=1
		 else
			Multiple=0
		 end if 
 end if
 
 
 
  if (errore<>0) then
	  if (errore=1) then
		 response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4, RE="&RE)
	  end if 
	  if (errore=2) then
		response.write("Controlla che non ci siano campi lasciati vuoti")
	  end if 
	  if (errore=3) then
		response.write("Controlla le risposte esatte (max 3 vere)")
	  end if 
	   if (errore=4) then
		response.write("Controlla le risposte esatte, valori ammessi 0 per (Falso) o 1 per (Vero)")
	  end if 
   
   %>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
   else
   
   
            QuerySQL1="Select * from Setting where Id_Classe='" & id_classe&"';"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
			Valutato=rsTabella.fields("Valutato") 
			DVAbilitato=rsTabella.fields("DVAbilitato")
			rsTabella.close
			if (strcomp(session("Admin"),"false")=0) then ' se non sono amministratore alloro conto i caratteri altrimenti metto il voto del form
					
					if Valutato=1  then	
					
						if len(Spiegazione)<100 then
							voto=0
							Segnalata=1									
						else
							Segnalata=0
							Voto=1  
						end if
								 
						if (DVAbilitato=1)and (len(Spiegazione)>350)   then
								voto=2
								Segnalata=0
						 end if 

					else
						Voto=0
					end if
			end if
			 
   
   
   
   
   
   
   ' per la spiegazione della domanda 
     
   
    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url1= "../Materie/"&Session("ID_Materia")& "/" & Cartella & "/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	url3=Replace(url,"\","/")
	url=url3

  ' per il testo della domanda plus
     url4=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url4=Replace(url4,"\","/")
	 
    

 


   
      
     ' QuerySQL ="UPDATE Domande SET Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "', Data = '" & DataTest & "', Voto = '" & voto & "', In_Quiz = '" & INQUIZ &"' Data= '" & DATA & "'  WHERE CodiceDomanda =" &ID&";"
'	
 if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then  
 
 QuerySQL ="UPDATE Domande SET Segnalata = '" & Segnalata & "',Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "', Voto = " & voto & ", In_Quiz = " & INQUIZ &" ,Data= '" & cdate(DATA) & "'  WHERE CodiceDomanda =" &ID&";"
'QuerySQL ="UPDATE Domande SET Segnalata = '" & Segnalata & "',Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "',  In_Quiz = " & INQUIZ &" ,Data= '" & cdate(DATA) & "'  WHERE CodiceDomanda =" &ID&";"
	'response.write(QuerySQL)	
	ConnessioneDB.Execute(QuerySQL)
end if	 

if  (Session("Admin")=True) then  
 
 QuerySQL ="UPDATE Domande SET Segnalata = '" & Segnalata & "',Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "', Voto = " & voto & ", In_Quiz = " & INQUIZ &" ,Data= '" & cdate(DATA) & "'  WHERE CodiceDomanda =" &ID&";"
'QuerySQL ="UPDATE Domande SET Segnalata = '" & Segnalata & "',Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "',  In_Quiz = " & INQUIZ &" ,Data= '" & cdate(DATA) & "'  WHERE CodiceDomanda =" &ID&";"
	'response.write(QuerySQL)	
	ConnessioneDB.Execute(QuerySQL)
end if	 
 'response.write(Session("Admin"))
	'response.write(QuerySQL)	
 
if Segnalata=1 then
' notifica
Azione="<a  target=blank href=inserisci_valutazione.asp?daQuaderno=1&cod="&CodiceAllievo&"&CodiceDomanda="&ID&"&cla="&Session("Id_Classe")&">Ho segnalato una tua domanda !</a>"
	 Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."
	 QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
	 ConnessioneDB.Execute(QuerySQL)

end if


If Err.Number <> 0 then
  DescrizioneErrore = Err.Description
  Pagina = Request.ServerVariables("url")
  Spiegazione1="Errore nell'esecuzione della query di aggiornamento"
  Riga=160
  Call GestisciErrore(DescrizioneErrore,Spiegazione1,Pagina,Riga)
  Err.Number=0
End If


'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
'Create the FSO.
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(Cartella)
'response.write(url)
Err.Number=0
objFSO.DeleteFile url

If Err.Number <> 0 then
'  NumeroErrore = Err.Number
  DescrizioneErrore = Err.Description
  Pagina = Request.ServerVariables("url")
  Spiegazione1="Tentativo di cancellare la vecchia versione del file inesistente"
  Riga=179
   
'  Source=Err.Source
  Call GestisciErrore(DescrizioneErrore,Spiegazione1,Pagina,Riga)
  Err.Number=0
End If

Set objCreatedFile = objFSO.CreateTextFile(url, True)
' Write a line with a newline character.
objCreatedFile.WriteLine(Spiegazione)
'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
objCreatedFile.Close
' per aggiornare la domanda plus
if Tipodomanda=1 then
	objFSO.DeleteFile url4
	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
	' Write a line with a newline character.
	objCreatedFile.WriteLine(TestoDomandaPlus)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
end if 

'On Error Resume Next
If Err.Number = 0 Then

Response.Write "<span class='alert-success'>Modifica avvenuta!</span> "
'response.redirect "../Cclasse/quaderno.asp?stile="&session("stile")&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&DataClaq2="&Session("DataClaq2")&"&DataClaq="& Session("DataClaq")&"&tCap="&tCap&"&tSot="& tSot&"&tDom="& tDom

Else
Response.Write Err.Description 
Err.Number = 0
End If%>


 <h5><a href="../cClasse/quaderno.asp?DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>&id_classe=<%=id_classe%>&cla=<%=cla%>&CodiceAllievo=<%=CodiceAllievo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua 
		a valutare o modificare le domande...</a></h5>
	<p>&nbsp;</p>
    <% end if%>                     
                      
              
                      </div>         
			        </div>
			      </div>
			    </div>
	
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         

			 
	</body>

 </html>

