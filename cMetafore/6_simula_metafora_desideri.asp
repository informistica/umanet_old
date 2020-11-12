<%@ Language=VBScript %>



<!doctype html>
<html>
<head>
   
   <title>Simula Client/Server</title>   
   <meta charset="UTF-8">
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
<!-- Le styles -->
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap.css" rel="stylesheet">
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap-responsive.css" rel="stylesheet">
    
    <link href="../../../guida/css/pageguide.css" rel="stylesheet">
   <link rel="stylesheet" href="../../css/style.css">
    

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
	 

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
	
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!");

location.href="studente_domande.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"

//location.href=window.history.back();
 }
 </script>
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 
 </script>
  
</head>

 
<%  
  
  CodiceMetafora = Request.QueryString("CodiceMetafora")
  Num = Request.QueryString("Num")
  Num=Num+1
  
 %> 
 
 <script language="JavaScript">
 var contSi,contNo,Topolino,Formaggio,Fame,Labirinto,Strada,Strada_OK,Strada_KO,Testata,Motivato,v;
 //variabili per singolare plurale
 var plurale,plurale1,volere,fare,avere,raggiungere,scegliere,allontanarsi,allontanarsi1,scontrarsi,continuare;

function popup(pagina) 
{ 
window.open(pagina,'','width=750,height=500, left=700,top=50,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
} 
 
 function popup2(pagina) 
{ 
window.open(pagina,'','width=550,height=230, left=30,top=50,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
} 
 function popup3(pagina) 
{ 
window.open(pagina,'','width=550,height=180, left=30,top=350,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
}
 

function inizio() {
     SoggettoC=document.parametri.txtSoggettoC.value.toUpperCase();
	 DomandaC=document.parametri.txtDomandaC.value.toUpperCase();
     distanza = parseInt(document.parametri.txtTolleranzaC.value.toUpperCase());
	 MotivazioneC=document.parametri.txtMotivazioneC.value.toUpperCase();
	 DesiderioC=document.parametri.txtDesiderioC.value.toUpperCase();
	 BisognoC=document.parametri.txtBisognoC.value.toUpperCase();
	 SoggettoS=document.parametri.txtSoggettoS.value.toUpperCase();
	 RispostaS=document.parametri.txtRispostaS.value.toUpperCase();
	 MotivazioneS=document.parametri.txtMotivazioneS.value.toUpperCase();
	 DesiderioS=document.parametri.txtDesiderioS.value.toUpperCase();
	 BisognoS=document.parametri.txtBisognoS.value.toUpperCase();
	 TipoEvento=document.parametri.txtTipoEvento.value.toUpperCase();
	// Terremoto=document.parametri.txtTerromoto.value.toUpperCase();
	
	// plurale=SoggettoC.search(/ e /i); //se è presente e oppure E è >0
     plurale=SoggettoC.search(";"); //faccio mettere , per indicare il prurale
 
	 if (plurale == -1){
	     volere="vuoi";
		 raggiungere="raggiungerai";
		 avere="hai";
		 scegliere="scegli";
		 avvicinarsi="ti avvicina";
		 allontanarsi="ti allontana";
		 allontanarsi1="ti sei allontanato troppo hai";
		 scontrarsi="e ti sei scontrato";
		 continuare="continua";
		 fare="ci sei quasi fai";  
	  }  
	  
	  else 
      { 
 
         volere="volete";
		 raggiungere="raggiungerete";
		 avere="avete";
		 scegliere="scegliete";
		 avvicinarsi="vi avvicina";
		 allontanarsi="vi allontana";
		 allontanarsi1="vi siete allontanati troppo avete";
		 scontrarsi="e vi siete scontrati";
		 continuare="continuate";
		 fare="ci siete quasi fate";
	     
	 
	   } 
	  
     contSi=0;
     contNo=0;
	 Motivato=0;
     document.parametri.Storia.value="Distanza Iniziale dall'obiettivo = (" + distanza + ")"; 
	 document.parametri.Storia.value=document.parametri.Storia.value+"\n\n " + SoggettoC + " e " + SoggettoS + " volete stabilire relazione Client Server ?";
     document.parametri.BSx.value='No';
     document.parametri.BDx.value='Si';
	 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
}

function bottoneSx() {
	  
if (Motivato == 0 )
{
   document.parametri.Storia.value=document.parametri.Storia.value+ " NO! \n\n Mancando relazione tra " + SoggettoC + " e "  + SoggettoS + " non e' possibile far interagire la richiesta di '"+DomandaC+"' (manifestata da " + SoggettoC+ ") con la risposta '"+RispostaS +"' (manifestata da " + SoggettoS+").";
   document.parametri.Storia.value=document.parametri.Storia.value+"\n\n " + SoggettoC + " e " + SoggettoS + " volete stabilire relazione Client Server ?";
   document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
}
else
    {
    if ((contSi-contNo)==(-distanza)) { 	
    document.parametri.Storia.value=document.parametri.Storia.value +"\n\n :-(  " + SoggettoC + " e " + SoggettoS + " vanno incontro ad un 'Terremoto Culturale' .";
	 popup("../../U-ECDL/img/PaginaClienteOsteNO.htm");  
     popup2("../../U-ECDL/img/paginaMessaggioClienteNO.asp?CodiceMetafora=<%=CodiceMetafora%>"); 
     popup3("../../U-ECDL/img/paginaMessaggioOsteNO.asp?CodiceMetafora=<%=CodiceMetafora%>"); 
		        
	 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;        
    }
    else {
    	contNo++;
       document.parametri.Storia.value=document.parametri.Storia.value + "\n\nATTENZIONE  " + SoggettoS + "  la scelta  '"+ document.parametri.BSx.value +"' " + allontanarsi +" da soddisfare l'aspettativa '"+ DomandaC + "' come richiesto da " + SoggettoC;
        document.parametri.Storia.value=document.parametri.Storia.value + "\n\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
		document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
        
    }	
    document.parametri.Storia.value=document.parametri.Storia.value + "\n\n  " + SoggettoS + "  quale risposta " + scegliere +" ?  ";
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
    }	  
		
}

function bottoneDx() {
	  
		
if (Motivato==1)  {
      if ((contSi-contNo)==distanza) {
           document.parametri.Storia.value=document.parametri.Storia.value + "\n :-) COMPLIMENTI  " + SoggettoC + " "+ SoggettoS + " avete stabilito una relazione coerente !!! \n\n ";
		   popup("../../U-ECDL/img/PaginaClienteOsteSI.htm");  
		   popup2("../../U-ECDL/img/paginaMessaggioClienteSI.asp?CodiceMetafora=<%=CodiceMetafora%>"); 
		   popup3("../../U-ECDL/img/paginaMessaggioOsteSI.asp?CodiceMetafora=<%=CodiceMetafora%>"); 
		        
		 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;

		  }
      else
      {
      contSi++;
      document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n  "+SoggettoS+"  la scelta  '"+document.parametri.BDx.value+"' " +avvicinarsi +" a  '"+ DomandaC+ "' come richiesto da " + SoggettoC +  " "  + continuare + " cosi' !  ";
      document.parametri.Storia.value=document.parametri.Storia.value + "\n\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
	  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		  if (contSi-contNo ==(distanza))
		  {
			 document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n Coraggio " + fare + " l'ultimo passo ! '";
		  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		  }
		  document.parametri.Storia.value=document.parametri.Storia.value + "\n\n  " + SoggettoS + "  quale risposta " + scegliere +" ?  ";
	
		  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
      }
    }
    else{
	document.parametri.Storia.value=document.parametri.Storia.value + "SI!  \n\n " + SoggettoS + "  quale risposta  " + scegliere +" ? '";
	if (TipoEvento.search("COERENTE") >0)
	 {
		 document.parametri.BSx.value="Non " + RispostaS;
	     document.parametri.BDx.value=RispostaS;
		 
	 }
	 else
	 {
		 document.parametri.BSx.value=RispostaS;
	     document.parametri.BDx.value="Non " + RispostaS;
		 
	 }
	 
   
    Motivato=1;
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
	    }
}


</script>


%>
<%
  Response.Buffer = true
 ' On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
   ' if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	<!-- <BODY onLoad="showText2();"> </BODY>-->
  <% 'else %>
    
    <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" >

  <% 'end if %>

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  
  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
      'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBCopiatestonline")
    
	
	
    
'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
'              "DBQ=" & Server.MapPath("../database/Copiatestonline.mdb")

 

   
%>
	<div id="navigation">
     
        <% 
		
 
		 
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
						<h1> <i class="icon-comments"></i> Metafora interattiva </h1> 
                    
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
							<a href="#more-files.html">Libro U-WWW</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Metafore</a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                              
                               
                                <%
 QuerySQL="SELECT * " &_
" From Elenco_Metafore_Desideri " &_
" Where CodiceMetafora =" & CodiceMetafora & "" 
 ' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logSimulazione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close  
	
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySQL)	
'response.write(rsTabella("SoggettoC"))
     Codice_Test=rsTabella("ID_Paragrafo")
	 Cartella=rsTabella("Cartella")


%>
              
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> <%=rsTabella("Titolo")%>:&nbsp;<%=rsTabella("Tit")%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 		<% 'response.write(Cartella&"_U_3_3")
' response.write("<br>"&Codice_Test)	%>						 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
               
						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora N.(<%=CodiceMetafora%>)</h3>
							</div>
							<div class="box-content">
							
							
    
                              	<form action="inserisci_metafora_dbdesideri1.asp?daSimulazione=1&CodiceMetafora=<%=CodiceMetafora%>"  name="parametri"  method="POST" class="form-vertical">
 
        
                                  <div class="control-group" id="btnAggiorna">
										 
										<div class="controls">
                                           <% if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True)  or (Condiviso=1) then%>
                                            <input type="button" value="Aggiorna" onClick="aggiornaMetafora();" class="btn" name="BAggiorna"><br>
                                            <%end if%>
										</div>
									</div>
 
  <fieldset id="idClient">
 
                                   <div class="control-group">
										<label for="textfield" class="control-label"><b>Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" placeholder="Soggetto che manifesta un aspettativa" class="input-xxlarge"  id="txtSoggettoC"  name="txtSoggettoC"  value="<%= rsTabella("SoggettoC") %>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Domanda</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" placeholder="Aspettativa" class="input-xxlarge"  id="txtDomandaC" name="txtDomandaC" value="<%=rsTabella("DomandaC")%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"  id="txtMotivazioneC"  name="txtMotivazioneC" value="<%=rsTabella("MotivazioneC")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"  id="txtDesiderioC" name="txtDesiderioC" value="<%=rsTabella("DesiderioC")%>">
										</div>
									</div>
                               
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge"  id="txtBisognoC" name="txtBisognoC" value="<%=rsTabella("BisognoC")%>">
										</div>
									</div>
                                     <div class="control-group">
										<label for="textfield" class="control-label"><b>Tolleranza del Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" placeholder="Indice della tensione che può sopportare" class="input-mini"  id="txtTolleranzaC"  name="txtTolleranzaC" value="<%=rsTabella("TolleranzaC") %>" >
										</div>
									</div>
                                    
                                     </fieldset>
                                     <hr>
                                     <div class="line"></div>
                                  <fieldset id="idServer">
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Server</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6"> 
											<input type="text" placeholder="Soggetto che risponde alla richiesta" class="input-xxlarge" id="txtSoggettoS"  name="txtSoggettoS" value="<%=rsTabella("SoggettoS")%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Risposta</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" checked="true">  
											<input type="text" placeholder="Risposta alla richiesta" class="input-xxlarge" id="txtRispostaS" name="txtRispostaS" value="<%=rsTabella("RispostaS")%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Ragioni che sostengono la Risposta" class="input-xxlarge" id="txtMotivazioneS"  name="txtMotivazioneS" value="<%=rsTabella("MotivazioneS")%>" >
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Desiderio che sostiene la Motivazione" class="input-xxlarge"   id="txtDesiderioS" name="txtDesiderioS" value="<%=rsTabella("DesiderioS") %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge" id="txtBisognoS" name="txtBisognoS" value="<%=rsTabella("BisognoS")%>" >
										</div>
									</div>
                                    </fieldset>
                                      <div class="control-group" id="tipoEvento">
										<label for="textfield" class="control-label"><b>Tipo di Evento</b></label>
										<div class="controls">
                                         <%if  rsTabella("TipoEvento")="PARADOSSALE" then %>
                                            <input type="text" name="txtTipoEvento" value="Paradossale" class="input-small"> 
                                          <% else%>
                                               <input type="text" name="txtTipoEvento" id="txtTipoEvento" value="Coerente"  class="input-small"> 
                                          <%end if%>
										</div>
									</div>
                                      
   
                                    
                                    
									
								
                                  
                                     
 
                                    
                                    <div class="form-actions">
		<center>							 
    <p> <span id="idInizio"> <input type="button" value="INIZIO" name="BInizia" class="btn-primary" onClick="inizio()"></span> </p>
  <p>  <span id="btnSxDx">  <input type="button" value="  " name="BSx" onClick="bottoneSx()" class="btn">
         <input type="button" value="  " name="BDx" onClick="bottoneDx()" class="btn"></span></p> <!--Definisce i due bottoni del form --></center>
	</div></p> 
	<!--<input type="button" value="clicca" onClick="<%'call ciao1()%>">-->
	 
								<div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Narr@azione</b></label>
										<div class="controls">
											<textarea name="Storia" id="textarea" rows="15" class="input-block-level"><%=					Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
								 
									</div>
								
<input type="hidden" value="<%=Cartella%>" id="cartella">
<input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
<input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
<input type="hidden" value="<%=Modulo%>" id="Modulo">
<input type="hidden" value="<%=Paragrafo%>" id="Paragrafo">
<input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
                                
                                </form>
							</div>
						</div>
					</div>
                      
                      
                 
                  
                      
                      
                      
                      
             
                      </div>         
			        </div>
			      
			    
	
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
      
      
          <script src="../../../guida/docs/lib/bootstrap/js/bootstrap-dropdown.js"></script>
    <script src="../../../guida/docs/lib/google-code-prettify/prettify.js"></script>

    <script src="../../../guida/js/jquery.pageguide.js"></script>
    <script language="javascript">
     
	 function aggiornaMetafora() {
	 
	 var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;		  
	 
			txtSoggettoC=document.getElementById("txtSoggettoC").value;
			txtDomandaC=document.getElementById("txtDomandaC").value;
			txtMotivazioneC=document.getElementById("txtMotivazioneC").value;
			txtDesiderioC=document.getElementById("txtDesiderioC").value;
			txtBisognoC=document.getElementById("txtBisognoC").value;
			txtSoggettoS=document.getElementById("txtSoggettoS").value;
			txtRispostaS=document.getElementById("txtRispostaS").value;
			txtMotivazioneS=document.getElementById("txtMotivazioneS").value;
			txtDesiderioS=document.getElementById("txtDesiderioS").value;
			txtBisognoS=document.getElementById("txtBisognoS").value;
			txtTipoEvento=document.getElementById("txtTipoEvento").value;
			txtTolleranzaC=document.getElementById("txtTolleranzaC").value;
			 
			dati2="&txtSoggettoC="+txtSoggettoC+"&txtDomandaC="+txtDomandaC+"&txtMotivazioneC="+txtMotivazioneC+"&txtDesiderioC="+txtDesiderioC+"&txtBisognoC="+txtBisognoC+"&txtSoggettoS="+txtSoggettoS+"&txtRispostaS="+txtRispostaS+"&txtMotivazioneS="+txtMotivazioneS+"&txtDesiderioS="+txtDesiderioS+"&txtBisognoS="+txtBisognoS+"&txtTipoEvento="+txtTipoEvento+"&daSimulazione=1"+"&txtTolleranzaC="+txtTolleranzaC;		 
	 
	
	dati="cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&CodiceMetafora="+CodiceMetafora+"&Codice_Test="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo; 
    var url = "7_aggiorna_metafora_ajax.asp?"+dati+dati2;			   
	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
	  if (xhttp.readyState == 4 && xhttp.status == 200) {
		  var testo = xhttp.responseText;		
		  alert(testo);			 
	  }
	};
	xhttp.open("GET", url, true);
	xhttp.send();	

	 
	 }
	 
	  /**
       * Helper Functions
       */

      // View source of current page in a new window
      function viewsource(e){
        window.open("view-source:" + window.location, 'jquery.pageguide.source');
      }

      // Smooth scroll to anchor
      function scrollTo(e) {
        e.preventDefault();

        var anchor = e.currentTarget.hash.slice(1);
            $t = $('a[name=' + anchor + ']');

        if (!$t.size()) return;

        var dvh = $(window).height(),
            dvtop = $(window).scrollTop(),
            eltop = $t.offset().top,
            mgn = {top: 100, bottom: 100};

        var scrollTo = eltop - mgn.top;

        $('html,body').animate({
          scrollTop: scrollTo
        }, {
          duration: 500
        });
      }

      // Example guides
	  
	  </script>
         
      <script language="javascript" type="text/javascript" src="../jsguide/clientserversimula.js"> </script> 
		
		</div> <!--fine main-->
        </div>
        
         

			 
	</body>

 </html>

