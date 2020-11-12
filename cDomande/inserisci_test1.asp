<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
    <meta charset="UTF-8">

   <title>Inserisci test</title>   
     <meta https-equiv="Content-Type" content="text/html;" />

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

		<!-- Bootstrap -->
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

	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
  
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione � scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
 <script language="javascript" type="text/javascript"> 
function showText3() {window.alert("La domanda � gi� stata inserita, la puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
 </script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<% if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
   <body class='theme-<%=session("stile")%>'>
  <% end if %>


<%  'on error resume next
   Dim Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,errore
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
    Capitolo = Request.QueryString("Capitolo")
	 Paragrafo = Request.QueryString("Paragrafo")
   Quesito=Request.Form("txtDomanda")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
      by_UECDL=Request.QueryString("by_UECDL")
	   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
   predomanda = Request.QueryString("predomanda") 
  ID_Predomanda=Request.QueryString("ID_Predomanda") 
  
%>
	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <!-- #include file = "tabella_corrispondenze.inc" -->
        
		 <!-- #include file = "../service/replacecar.asp" -->
        	  
          
         
	</div>
    
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


  %>
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserisci test</h1> 
                    
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
							<a href="#">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Verifica</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							<a href="#">Inserisci test</a>
                           
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%> : <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 <%QuerySQL1="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
	Valutato=rsTabella.fields("Valutato") 
	DVAbilitato=rsTabella.fields("DVAbilitato")



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
  
 '  else
  ' 		In_Quiz_Stud=rsTabella.fields("In_Quiz") 
   'end if
 'response.write("In_Quiz_Stud"&In_Quiz_Stud)
   rsTabella.close
	
CodiceCap=Request.Cookies("Dati")("CodiceCap")
Num=Request.QueryString("Num")
Capitolo=Request.QueryString("Capitolo")
Multiple=Request.QueryString("Multiple")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
predomanda = Request.QueryString("predomanda") 
ID_Predomanda=Request.QueryString("ID_Predomanda") 
 ' serve per controllare la validit� della RispostaEsatta, se esiste nel vettore � giusta altrimenti no

lingua=request.form("lingua")




'response.Write("Tipo="&tipo)
   if strcomp(Tipo,"0")=0 then
	   Domanda = Replace(Request.Form("txtDomanda"),"'",Chr(96))	
	   Domanda = Replace(Request.Form("txtDomanda"),Chr(34),Chr(96))
	  ' response.write("ciao"&domanda)
	   Domanda=  ReplaceCar(Domanda)
	    
   else
       Titolo=   Replace(Request.Form("txtDomanda"),Chr(34),Chr(96))
	    Titolo=   Replace(Request.Form("txtDomanda"),"'",Chr(96))	   
	   Titolo=  ReplaceCar(Titolo) 
	   Domanda = Replace(Request.Form("txtDomandaplus"),"'",Chr(96)) 
	   Domanda = Replace(Request.Form("txtDomandaplus"),Chr(34),Chr(96)) 
   end if

 'response.write("ciao")
   
  
   
   
  'Domanda=formattaChar(Domanda)
    Domanda=ReplaceCar(Domanda)
	'response.write("ciao0")
   R1 = Replace(Request.Form("txtR1"),"'",Chr(96))
   R2 = Replace(Request.Form("txtR2"),"'",Chr(96))
   R3 = Replace(Request.Form("txtR3"),"'",Chr(96))
   R4 = Replace(Request.Form("txtR4"),"'",Chr(96))
   Spiegazione=Request.Form("S1")
  ' Spiegazione= Replace(Spiegazione, Chr(34), Chr(96))
  Spiegazione= ReplaceCar(Spiegazione)
   
 ' se non � una risposta multipla faccio il solito controllo di validit� sulla Risposta esata
   errore=0
   RE=Request.Form("txtRE")
	' response.write("RE="&RE)
  ' if (len(Request.Form("txtRE"))=0) then 
  '     errore=2
  ' end if 
   'errore=0 
   
   'qua metto il controllo per verificare se la domando o frase � stata gi� inserita. 
   
 
 
 
 
 
%>
 
<!--#include file="inserisci_test1_include.asp"-->  
 

	</font>   
	 
		
      <h5><a href="inserisci_test.asp?Multiple=<%=Multiple%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&CodiceSottoPar=<%=CodiceSottoPar%>&Sottoparagrafo=<%=Sottoparagrafo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua ad inserire...</a></h5>
	<p>&nbsp;</p>
	
	
	<div id="piede_pagina" align="left">
 <%if left(CodiceTest,1)="U" then  
    	 if by_UECDL<>"" then %>
                   <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
    <h4><a href="../cClasse/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna al Libro... </a></h4> 
                <%else%>
                    <!-- REDIRECT INTELLIGENTE  -->
    <h5 ><a href="../cClasse/scegli_azione_test.asp?by_UECDL=1&id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottoPar%>"> Torna a scelta verifica </a></h5> 
                <%end if	
        
   else%>
 
    	 <% if predomanda<>"" then %>
    	       <h5><a href="../cClasse/home_app.asp?divid=<%=Session("divid")%>&id_classe=<%=Session("Id_Classe")%>&cartella=<%=Cartella%>"> Torna al Libro </a></h5> 
         <% else%>
      
  			 <h5 ><a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottoPar%>"> Torna a scelta verifica </a></h5> 
       <% end if%>
 <%end if %>				 

   
<%else ' if (errore=0)
   'response.write("e="&errore)
   %>
    <div class="alert-error">
	<% 
  if (errore=1) then
     response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
  end if 
  if (errore=2) then
  response.write("domanda="&Domanda &"r1="&R1 & "r2="&R2& "r3="&R3 & "r1="&R4&  "r1="&R4)
    response.write("Controlla che non ci siano campi lasciati vuoti")
  end if 
  if (errore=3) then
    response.write("Controlla le risposte esatte (max 3 vere)")
  end if 

  %>
  </div>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
end if 			
%>
 									 
	 
	 
				 
				 
                   
                   
 
		  			 
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        

			 
	</body>

 </html>

