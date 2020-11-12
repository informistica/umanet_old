<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
      <meta charset="UTF-8">
   <title>Inserisci test V/F</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	
	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">

    
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	
    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
    
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
 <script language="javascript" type="text/javascript"> 
function showText3() {window.alert("La domanda è già stata inserita, la puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
 </script>
	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

  


   
</head>
<% if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
   <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

  <% end if %>
  
	<div id="navigation">
     
        <% 
		
  Dim Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,errore
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   Quesito=Request.Form("txtDomanda")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
      by_UECDL=Request.QueryString("by_UECDL")
	   Capitolo=request.QueryString("Capitolo")
   Paragrafo=request.QueryString("Paragrafo")
    Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
	  
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserisci test </h1> 
                    
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
							<a href="#more-files.html">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Verifica</a>
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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%> : <%=Paragrafo%>
                        <% if  CodiceSottopar<>"" then %>
                          /&nbsp;<%=Sottoparagrafo%>
                         <% end if%>
                         </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			   <div class="box-content"> 
                  <%   
                       Function controlla(RisposteEsatte)
	 controlla=0
	 i=0
	 while (i<=16) and not(esiste)
	 'response.Write(v2(i) & "=" & RisposteEsatte & "<br>")
		if strcomp(v2(i),RisposteEsatte)= 0 then 
		    esiste=true
		    controlla=1
			'response.Write(" Trovato <br>")
		end if
		i=i+1
	 wend
 end function
 
 function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"a"&Chr(96))
  sAns = Replace(sAns,chr(225),"a"&Chr(96))
  sAns = Replace(sAns,chr(232),"e"&Chr(96))
  sAns = Replace(sAns,chr(233),"e"&Chr(96))
  sAns = Replace(sAns,chr(236),"i"&Chr(96))
  sAns = Replace(sAns,chr(237),"i"&Chr(96))
  sAns = Replace(sAns,chr(242),"o"&Chr(96))
  sAns = Replace(sAns,chr(243),"o"&Chr(96))
  sAns = Replace(sAns,chr(249),"u"&Chr(96))
  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  
  
'ReplaceCar = sAns
ReplaceCar=sInput
end function
    
	 
' controllo se la domanda è già stata inserita 
' lo tolgo perchè a volte non funziona
   
	'	querySQL="Select * from Domande where Id_Stud='" & Session("CodiceAllievo") & "' and (Id_Predomanda="&clng(ID_Predomanda)&" and ID_Predomanda<>0 or Quesito='" &Quesito& "');"
		
	 'Set objFSO = CreateObject("Scripting.FileSystemObject")
'					url1="C:\Inetpub\umanetroot\anno_2012-2013\logfrasi.txt"
'					Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'					objCreatedFile.WriteLine(querySQL)
'					objCreatedFile.Close	
		
		' on error resume next 
	'	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'If not(rsTabella.BOF=True And rsTabella.EOF=True) Then 
				' esiste già non la faccio inserire  
			%>
           <!-- <BODY onLoad="showText3();"> 
		  Stai per essere reindirizzato all'home page ... </BODY>--> 
		  <% 
		 
    

'else
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
  
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   ' devo sapere a quale quiz contribuisce lo studente
   QuerySQL1="Select In_Quiz from Allievi where CodiceAllievo='" & CodiceAllievo &"';"
   Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
   
   
   ' devo farlo parametrico in base a max_in_quiz
  ' if Session("Admin")=True then
                  '  QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") & "';" 
'					Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
'					' se raggiungo il limite ricomncio
'					 
'					max_in_quiz=clng(rsTabella1("Max_In_Quiz"))%>
					 
   
   
     <% SELECT CASE Request.Form("inQuiz")
		 CASE "-1"
			 In_Quiz_Stud=-1
		' for i=1 to max_in_quiz 
		'	CASE i					 
           '     In_Quiz_Stud=i       		 
		' next 
		 
		 
		 CASE "1"
		   In_Quiz_Stud=1
		 CASE "2"
		   In_Quiz_Stud=2
		 CASE "3"
		   In_Quiz_Stud=3
		 CASE "4"
		   In_Quiz_Stud=4
		 CASE "5"
		   In_Quiz_Stud=5
		 CASE "6"
		   In_Quiz_Stud=6  
   	  CASE ELSE
	 END SELECT   
     
	 lingua=request.form("lingua") ' per la versione inglese
  ' else
  ' 		In_Quiz_Stud=rsTabella.fields("In_Quiz") 
  ' end if
   
   rsTabella.close
	'response.write("In_Quiz_Stud"&In_Quiz_Stud)
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
  Num=Request.QueryString("Num")
Capitolo=Request.QueryString("Capitolo")
Multiple=Request.QueryString("Multiple")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
predomanda = Request.QueryString("predomanda") 
ID_Predomanda=Request.QueryString("ID_Predomanda") 
 Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
 ' serve per controllare la validità della RispostaEsatta, se esiste nel vettore è giusta altrimenti no


'response.Write("Tipo="&tipo)
   if strcomp(Tipo,"0")=0 then
	   Domanda = Request.Form("txtDomanda")
	   'response.write("primo="&Domanda)
	   Domanda = Replace(Domanda, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Domanda=  Replace(Domanda,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
	   'response.Write("Domanda="&Domanda)
   else
       Titolo=   Request.Form("txtDomanda")
	    ' response.write("secondo")
	   Titolo=  Replace(Titolo, Chr(34), "'")
	   Titolo=  Replace(Titolo, "'",Chr(96))
	   Domanda = Request.Form("txtDomandaplus")
	   Domanda = Replace(Domanda, Chr(34), "'")
	   Domanda=  Replace(Domanda,"'",Chr(96))
	   'response.Write("Domanda="&Domanda)
	   ' response.Write("<br>Titolo="&Titolo)
   end if
  ' R1 = Request.Form("txtR1")
'   R1 = Replace(R1, Chr(34), "'")
'   R1=  Replace(R1,"'",Chr(96))
'
'
'   R2 = Request.Form("txtR2")
'   R2 = Replace(R2, Chr(34), "'")
'   R2=  Replace(R2,"'",Chr(96))
'
'   R3 = Request.Form("txtR3")
'   R3 = Replace(R3, Chr(34), "'")
'   R3=  Replace(R3,"'",Chr(96))
'   
'   R4 = Request.Form("txtR4")
'   R4 = Replace(R4, Chr(34), "'")
'   R4=  Replace(R4,"'",Chr(96))


   Spiegazione=Request.Form("S1")
 '  Spiegazione= Replace(Spiegazione, Chr(34), "'")
 '  Spiegazione=  Replace(Spiegazione,"'",Chr(96))
 ' se non è una risposta multipla faccio il solito controllo di validità sulla Risposta esata
   errore=0
  RE=Request.Form("VF") 
   
    ' RE=Request.Form("txtRE")
	' response.write("RE="&RE)
  ' if (len(Request.Form("txtRE"))=0) then 
  '     errore=2
  ' end if 
   'errore=0 
   
   'qua metto il controllo per verificare se la domando o frase è stata già inserita. 
   
 
%>
 
<!--#include file="inserisci_test_vf1_include.asp"-->  

 <h4><a href="inserisci_test_vf.asp?Multiple=<%=Multiple%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>">Continua ad inserire...</a></h4>
	<p>&nbsp;</p>
	
	
	<div id="piede_pagina" align="left">
 <%if instr(CodiceTest,"_U_")>0 then  
    	 if by_UECDL<>"" then %>
                   <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
    <h3><a href="../cClasse/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla Libro... </a></h3> 
                
   
                <%end if	
        
   else%>
    	       <h4 class="sottotitolo"><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&cartella=<%=Cartella%>"> Torna al Libro </a></h4> 
        
     
     
  			  
 <%end if %>				 

   
<%else ' if (errore=0)
   'response.write("e="&errore)
  if (errore=1) then
     response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
  end if 
  if (errore=2) then
    response.write("Controlla che non ci siano campi lasciati vuoti")
  end if 
  if (errore=3) then
    response.write("Controlla le risposte esatte (max 3 vere)")
  end if 
  
  
  %>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
end if 			
%>
                      
              
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

