<!doctype html>
<html>
<head>
<link rel="shortcut icon" href="../favicon.ico" />

<script src="../js/google.js"></script><!--<meta charset="utf-8">-->
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <meta charset="UTF-8">
	<title>Gestione Segnalazioni </title>

	<link rel="shortcut icon" href="../favicon.ico" />
 
	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- Easy pie  -->
	<link rel="stylesheet" href="../../css/plugins/easy-pie-chart/jquery.easy-pie-chart.css">
	<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">
    
     <!-- Notify -->
	 
	<link rel="stylesheet" href="../../css/plugins/gritter/jquery.gritter.css">
	
     <link href="../../../guida/css/pageguide.css" rel="stylesheet">
     <!-- Le styles -->
   <!-- <link href="../../../guida/docs/lib/bootstrap/css/bootstrap.css" rel="stylesheet">
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap-responsive.css" rel="stylesheet">
    
-->

	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery UI -->
    
     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
     <!-- Notify -->
	<script src="../../js/plugins/gritter/jquery.gritter.min.js"></script>

	 
	<!-- Flot -->
	<script src="../../js/plugins/flot/jquery.flot.min.js"></script>
	<script src="../../js/plugins/flot/jquery.flot.resize.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
    <!--
    <script src="../../js/plugins/validation/jquery.validate.min.js"></script>
	<script src="../../js/plugins/validation/additional-methods.min.js"></script>
	
    --> 
     <script language="javascript" type="text/javascript" >
	 
 function validate2() {
	 
	 //continua a dare errore TypeError: document.frm0.txtNewCodiceAllievo is undefined, il form è definito nel file di inclusione 2_modifica_login_1.asp, rinuncio faccio il controllo lato server prima di inserire in db
	 
 alert(document.frm0.txtNewCodiceAllievo.value);
 if (document.frm0.txtNewCodiceAllievo.value=="")
	{
	   alert("Non hai inserito lo username !");
	 
	}
else
 if (frm0.txtNewPwd.value=="")
	{
	   alert("Non hai inserito la nuova password");
	}
 else
  if (frm0.txtNewPwd1.value=="")
	{
	   alert("Non hai inserito la conferma password");
	 
	}else
	 if (frm0.txtNewPwd1.value != frm0.txtNewPwd.value)
	{
	   alert("Le password non coincidono");
	 
	}
	else
	
	{
	    document.frm0.action = "modifica_pwd_new.asp?stato=<%=stato%>&cla=<%=cla%>&id_classe=<%=id_classe%>&divid=<%=divid%>" ;
		document.frm0.submit();
	 	
	   
    }
	
}


    </script>
	<script>
	$(window).ready(function () {	   
	
	   $('#msg').click();
	   
	  // event.stopPropagation();
	    
	});
	
</script>
	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->
	
	<!-- Favicon 
	<link rel="shortcut icon" href="img/favicon.ico" />-->
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />

</head>

<body class='theme-<%=Session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">
     <% 
	
function ReplaceCar(sInput)
dim sAns
   
  sAns=  Replace(sInput,"è","&egrave;")
 
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
 
ReplaceCar = sAns
 
end function
	
	if Session("CodiceAllievo")="" or Session("Id_Classe")="" then %>	 
				<script language="javascript" type="text/javascript"> 
				    window.alert("Sessione  scaduta, effettua nuovamente il Login!");
                    location.href="../../home.asp";
				</script>
				<%
				response.Redirect "../../home.asp"
				 
				 %>
 
<% end if%>

<%


'Response.AddHeader "Refresh", "600"

 ' Cartella=Request.QueryString("Cartella")
  Cartella=Request.QueryString("classe")
 ' response.Cookies("Dati")("Cartella")=Cartella
 ' TitoloCapitolo=Request.QueryString("Capitolo") 
 ' Paragrafo=Request.QueryString("Paragrafo")
  'Modulo=Request.QueryString("Modulo")
 ' CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Session("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")  
  dividA=request.QueryString("dividApro")
    On Error Resume Next
xEstrazione=request.querystring("xEstrazione")
id_classe=request.querystring("id_classe")
  Response.Cookies("Dati")("Id_Classe")=id_classe
classe=request.querystring("classe")
divid=request.querystring("divid")
 
PS=request.querystring("PS") ' vale 1 se devo mostrare anche i Punti Social chiamato da javasscript
if PS="" then ' per la prima chiamata mostrio i PS
   PS=1
end if
 
daStud=Request.QueryString("daStud") ' chiamato da href della classifica
daForm=Request.QueryString("daForm") ' chiamato dal bottone invia di scelta periodi dal al
daMenu=Request.QueryString("daMenu")


DataCla=request.form("txtData") 
DataCla2=request.form("txtData2")
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")


if daForm<>"" then
 ' Session("DataClaq")=DataClaq
 ' Session("DataClaq2")=DataClaq2
  
end if
if daStud<>"" then
  
end if

if DataCla="" then
   if DataClaq2<>"" then
      DataCla=DataClaq
	  DataCla2=DataClaq2
   else
     DataCla=Session("DataCla")
	 DataClaq=Session("DataCla")
	  DataClaq=Session("DataClaq")
	 DataClaq2=Session("DataClaq2")
	end if 
end if

'if daMenu<>"" then
'    DataCla=request.QueryString("DataClaq") 
'    DataCla2=request.QueryString("DataClaq2")
'end if
'if daStud<>"" then
'   'DataClaq= DataCla
'   'DataClaq2=DataCla2
'    DataClaq=request.QueryString("DataClaq") 
'	DataClaq2=request.QueryString("DataClaq2")
'   
'end if

'response.write(DataClaq & "<br>" & DataClaq2)
'if session("DataClaq")="" then
'Session("DataClaq")=DataClaq
'Session("DataClaq2")=DataClaq2
'else
' DataClaq=Session("DataClaq")
' DataClaq2=Session("DataClaq2")
' DataCla=Session("DataClaq")
' DataClaq=Session("DataClaq2")
' end if
'' response.write("dopo session OK "& DataClaq & "<br>" & DataClaq2) 
'' se è la prima chiamata il valore del form sopra la classifica è nullo
'if (DataCla<>"") and (DataCla2<>"") then
'	Session("DataCla")=DataCla
'	Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
'else
'   Session("DataCla")= Session("DataClaq")
'   Session("DataCla2")= Session("DataClaq2")
'end if
'  
  
  
  
  cod=Request.QueryString("cod")
  if strcomp(cod&"","")=0 then
     cod=Session("CodiceAllievo")
	
	 
  end if
  
 box_apri="toggleCapitolo"&request.querystring("tCap")
 box_apri1="toggleSottoPar"&request.querystring("tSot")
 box_apri2="toggleDomande"&request.querystring("tDom")
 box_apri3="toggleFrasi"&request.querystring("tFra")
 box_apri4="toggleNodi"&request.querystring("tNod")
 
  
  
  
  
function ReplaceCar(sInput)
dim sAns
 
  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
  sAns=  Replace(sAns,"'",Chr(96))
  
ReplaceCar = sAns
end function

   
  
  
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario
 
		%> 
        <!-- #include file = "../var_globali.inc" --> 
        
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->    
           
		 
                 
		<!-- #include file = "../include/navigation.asp" --> 
            
        <!-- #include file = "../extra/test_server.asp" --> 
        
		<!-- #include file = "../include/formattaDataCla.inc" --> 

        <%
		
		
	 
		
	' PRELEVO IN ANTICIPO IL CONGOME NOME NEL CASO LA QUERY 2 NON TROVI NULLA IN QUEL PERIODO E QUINDI RESTITUISCA NULL	
		  cod=Request.QueryString("cod")
		 
		%>	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
   
      <!-- #include file = "../include/menu_left.asp" -->
			<div id="main">
				<div class="container-fluid">
				 
	           
  
              
   <div class="box-title">
				        <h3 > <a name="#"><i class="icon-reorder"></i></a>  Compiti segnalati<small title="Punti totalizzati"> (da correggere)</small></h3>
			          </div>
 <div class="row-fluid">
					 
					
</div>

 <div class="bs-docs-example">
 
 <!-- #include file = "../cUtenti/adovbs.inc" -->
 
 
 
 <%
 
 
 'per le store procedure
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd1.activeconnection = conn  
set cmd2.activeconnection = conn 
set cmd3.activeconnection = conn 

 

 
 QuerySQL="SELECT * FROM MODULI_CLASSE " &_
" WHERE Id_Classe='" & id_classe & "';"

'" WHERE Id_Classe='" & id_classe & "' or Id_Classe='"&Id_ClassePassato&"';"

  Set rsTabellaModuli = ConnessioneDB.Execute(QuerySQL)
   '  response.write(QuerySQL)
 %>
 
 <% k=0 
 p=0
   compiti=0 ' serve per mettere il box se non ci sono compiti inseriti
		     do while not rsTabellaModuli.EOF  
			 ' calcolo i punteggi frase per quel modulo
			 %>
			
			  <!-- #include file = "../cClasse/studente_domande_include/3_statistica_frasi_segnalazioni.asp" --> 
              <!-- #include file = "../cClasse/studente_domande_include/3_statistica_nodi_segnalazioni.asp" -->
              <!-- #include file = "../cClasse/studente_domande_include/3_statistica_domande_segnalazioni.asp" -->
             
           
                 
 <% 
' response.write(numrsFrasi& " --" & numrsNodi & "---" & numrsDomande)
 ' se è stato svolto almeno un compito mostro il capitolo
 if (numrsFrasi<>0) or (numrsNodi<>0) or (numrsDomande<>0)then  ' devo fare anche per nodi e domande mostro solo dove ci sono compiti svolti%>
 
               <div class="accordion-group">            
                  <div class="accordion-heading">
                    <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%>" href="#collapsenew<%=k%>"  id="toggleCapitolo<%=k%>" title="<%=k%>">
                        <%=rsTabellaModuli("Titolo") %><small> (<% Response.write(numrsFrasi+numrsNodi+numrsDomande)%>)</small>
                    </a>
                    
                  </div>
                <div id="collapsenew<%=k%>" class="accordion-body collapse"> 
                    <div class="accordion-inner">
                    <table class="table table-hover table-nomargin table-condensed">
                                                    <thead>
                                                        <tr align="center">
                                                        <th>
                     <%
on error resume next
						if numrsPreFrasi<>0 then
						percFrasi=fix((numrsFrasi/numrsPreFrasi)*10)/10*100
						else
						percFrasi=0
						end if
						if numrsPreDomande<>0 then
						percDomande=fix((numrsDomande/numrsPreDomande)*10)/10*100
						else
						percDomande=0
						end if
						if numrsPreNodi<>0 then
						percNodi=fix((numrsNodi/numrsPreNodi)*10)/10*100
						else
						percNodi=0
						end if
						numrsDomandeBack=numrsDomande%>
					 
						
						
						 <% QuerySQL="SELECT * FROM MODULI_PARAGRAFI_CLASSE " &_
" WHERE ID_Mod='" & rsTabellaModuli("ID_Mod") & "' and Id_Classe='"&id_classe&"';"
  'response.write(QuerySQL &" " & id_classe)
  Set rsTabellaParagrafi = ConnessioneDB.Execute(QuerySQL)%>                
                     
       <%
			
				' servono solo per i parametri per aprire tutti i compiti del cap, forse si può anche fare a meno usando i parametri di rsTabellaModuli 	%>
                <!-- #include file = "../cClasse/studente_domande_include/2_nodi_0.asp" -->  
        
                <!-- #include file = "../cClasse/studente_domande_include/2_domande_0_segnalazione.asp" -->  
                <!-- #include file = "../cClasse/studente_domande_include/2_frasi_0.asp" -->   
                                                    

  

                 
          
              
                       <ul class="pagestats style-3">
                     					  
											<li>
												
                                                       
                                                <div class="spark">
													<div title="% di Frasi svolte" class="chart" data-percent="<%=percFrasi%>" data-color="#368ee0" data-trackcolor="#d5e7f7">
													
													<%=percFrasi%> %
                                                    
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaFrasi.eof then %>
                                                 <span style="color:#000" title="Apri tutte le frasi del capitolo" href="../cFrasi/2inserisci_valutazioni_frasi.asp?TutteCap=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>"> 
                                                 <%end if%>
													<span class="name"><%=numrsFrasi%> su <%=numrsPreFrasi%></span>
                                                    <span class="name">PF.<%=numrsFrasi2%> </span>
                                                      </span>
												</div>
                                                
                                              
                                                
                                                
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Domande svolte" class="chart" data-percent="<%=percDomande%>" data-color="#56af45" data-trackcolor="#dcf8d7">
													<%=percDomande%> %
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaDomande.eof then %>
                                                  <a style="color:#000" title="Apri tutte le domande del capitolo" href="inserisci_valutazioni_segnalate.asp?Segnalate=1&Tutte=1&ID_MOD=<%=rsTabellaDomande("ID_MOD")%>&ID_PAR=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_MOD")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&TitoloParagrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">
                                                   <%end if%>
													<span class="name"><%=numrsDomandeBack%> su <%=numrsPreDomande%></span>
                                                    <span class="name">PD.<%=numrsDomande%> </span>
                                                    </a>
												</div>
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Nodi svolti" class="chart" data-percent="<%=percNodi%>" data-color="#f96d6d" data-trackcolor="#fae2e2"><%=percNodi%>%</div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaNodi.eof then %>
                                                 <a style="color:#000" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>"> 
													<%end if%>
                                                    <span class="name"><%=numrsNodi%> su <%=numrsPreNodi%></span>
                                                    <span class="name">PN.<%=numrsNodi2%> </span>
                                                    </a>
                                                   
												</div>
                                               
											</li>
                                            
                                            
                                     
										</ul>
                      </th>
                      
                                                             
                                                        </tr>
                                                    </thead>
                     </table>
                 					   
      
       
            
          
                     <% p=0 
		     do while not rsTabellaParagrafi.EOF  
                %>

 							    <!-- #include file = "../cClasse/studente_domande_include/2_frasi_1_segnalazioni.asp" -->   
                                <!-- #include file = "../cClasse/studente_domande_include/2_domande_1_segnalazioni.asp" -->   
                                <!-- #include file = "../cClasse/studente_domande_include/2_nodi_1_segnalazioni.asp" -->   
                                
                                
                                
                                
					 <!--Qua il controllo per vedere se ci sono compiti svolti per quel paragrafo-->    
                     <% 'Response.write(rsTabellaParagrafi("ID_Paragrafo") & numrsFrasi &" " & " " & numrsNodi & " " &numrsDomande & "<br>")%>
					<% if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0) then %>
                          
  
                                       
                          <div class="accordion-group">    
                          
                                      
                          <div class="accordion-heading">
                          
                            <a id="toggleSottoPar<%=k%><%=p%>" title="<%=k%><%=p%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%><%=p%>" href="#collapseTrenew<%=k%><%=p%>">
                            <%=rsTabellaParagrafi("Paragrafo") %> <small> (<% Response.write(numrsFrasi+numrsNodi+numrsDomande)%>)</small>
                            </a>
                            
                          </div>
                          
                          
                           
                           
                          <div id="collapseTrenew<%=k%><%=p%>" class="accordion-body collapse">       
                              <ul id="myTab3" class="nav nav-tabs">
                                <% if numrsFrasi<>0 then %>
                                  <li  class="active">
								  <%else%>
                                  <li>
								  <%end if%>
                                 <a id="toggleFrasi<%=k%><%=p%>" href="#profileFrasi<%=k%><%=p%>" data-toggle="tab">Frasi (<%=numrsFrasi%>)</a></li>
                                
                                   
                                    <% if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  <a id="toggleDomande<%=k%><%=p%>" href="#profileDomande<%=k%><%=p%>" data-toggle="tab">Domande (<%=numrsDomande%>)</a></li>
                                  
                                   
                                   
                                       <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  
                                  <a id="toggleNodi<%=k%><%=p%>" href="#profileNodi<%=k%><%=p%>" data-toggle="tab">Nodi (<%=numrsNodi%>)</a></li> 
                                       
                            </ul>
                            <div id="myTabContent2<%=k%><%=p%>" class="tab-content">
                             
                              <% if numrsFrasi<>0 then %>
                                  <div class="tab-pane fade in active" id="profileFrasi<%=k%><%=p%>">
                          
								  <%else%>
                                   <div class="tab-pane fade" id="profileFrasi<%=k%><%=p%>">
                          
								  <%end if%>
                             
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                    <a title="Apri tutte le frasi del paragrafo" style="color:#FFF"  href="../cFrasi/2inserisci_valutazioni_frasi.asp?TuttePar=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>"> 
                                                 Apri tutte le frasi: N(<%= numrsFrasi &") Pt(" & numrsFrasi2  & ") Pb("& round( numrsFrasi2/numrsFrasi,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                        <tr>
                                                            <th>Frase</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                    </thead>
                                                    <tbody>
                                                    
                                                    
                       
                     <% Sottoparagrafo=""
					' p=0
		     do while not rsTabellaFrasi.EOF  
			   if StrComp(Sottoparagrafo, rsTabellaFrasi("SotPar")) <> 0 then
			  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
			   Sottoparagrafo=rsTabellaFrasi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaFrasi("SotPar")%></b></center></td></th>  
			 <%end if%>                        
                                                        <tr>
                                                     
                                                             <%if rsTabellaFrasi("Segnalata")=1 then%>
                                                            <td > <a style="color:#F00"  href="../cFrasi/2inserisci_valutazione_frase.asp?Cartella=<%=rsTabellaFrasi("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&MO=<%=rsTabellaFrasi("ID_Mod")%>&VAL=<%=rsTabellaFrasi("Voto")%>&id_classe=<%=id_classe%>&tCap=<%=k-1%>&tSot=<%=k-1%><%=p%>&tFra=<%=k%><%=p%>"><%=rsTabellaFrasi("Chi")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaFrasi("Voto")%></td>
                                                             <%else%>
                                                              <td> <a  href="../cFrasi/2inserisci_valutazione_frase.asp?Cartella=<%=rsTabellaFrasi("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&MO=<%=rsTabellaFrasi("ID_Mod")%>&VAL=<%=rsTabellaFrasi("Voto")%>&id_classe=<%=id_classe%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tFra=<%=k%><%=p%>">   <%=rsTabellaFrasi("Chi")%></a></td>
                                                              <td><%=rsTabellaFrasi("Voto")%></td>
                                                              <%end if%>
                                                            <td><%=rsTabellaFrasi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaFrasi("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare la frase?');"  href="../cFrasi/cancella_frase.asp?cla=<%=d%>&cod=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tFra=<%=k%><%=p%>">
                                                            
                                                            
                                                           
 
                                                            
                                                            
                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     
                 <% f=f+1
				 '  p=p+1
				    rsTabellaFrasi.movenext()
				 loop%>
                 
                
                 
                 
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                              </div>
                              
                              
                              <% 
							'  p=0
							  if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <div class="tab-pane fade in active" id="profileDomande<%=k%><%=p%>">
                             
                                     <%else%>
                                          <div class="tab-pane fade" id="profileDomande<%=k%><%=p%>">
                             
                                     <%end if%>
                              
                                  
                                  
                    
                                   <!-- inizio blocco frasi che diventa domande-->  
                                  

                  
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaDomande.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte le domande"  href="inserisci_valutazioni_segnalate.asp?Segnalate=1&ID_MOD=<%=rsTabellaDomande("ID_Mod")%>&ID_PAR=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">  Apri tutte le domande:&nbsp;
                                                    N(<%= numrsDomande &") Pt(" & numrsDomande2  & ") Pb("& round( numrsDomande2/numrsDomande,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>         
                                                         <% if not rsTabellaDomande.eof then %>
                                                        <tr>
                                                            <th>Domanda</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessuna compito inserito</th>
                                                                                                                
                                                       </tr>
                                                    <%end if%>
                                                        
                                                        
                                                    </thead>
                                                    <tbody>
                   
                      <% Sottoparagrafo=""
					' p=0
					n=0
			 
		     do while not rsTabellaDomande.EOF  
			 
			   
			   if ((StrComp(Sottoparagrafo, rsTabellaDomande("SotPar")) <> 0) ) then
			   Sottoparagrafo=rsTabellaDomande("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaDomande("SotPar")%></b></center></td></th>  
	 
            <%end if%>
          
               
                                                    
                                                        <tr>
                                                                         
                                                                        
                                                            
                                                             <%if rsTabellaDomande("Segnalata")=1 then%>
                                                            <td > <a style="color:red"  href="inserisci_valutazione.asp?daQuaderno=1&Segnalate=1&&Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&URL=<%=rsTabellaDomande("URL_Teoria")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>"><%=rsTabellaDomande("Quesito")%></a> (<%=rsTabellaDomande("CodiceAllievo")%>)</a></td>
                                                             <td style="color:#F00"><%=rsTabellaDomande("Voto")%></td>
                                                             <%else%>
                                                              <td> <a   href="inserisci_valutazione.asp?daQuaderno=1&Segnalate=1&Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">  <%=rsTabellaDomande("Quesito")%></a>(<%=rsTabellaDomande("CodiceAllievo")%></td>
                                                              <td><%=rsTabellaDomande("Voto")%></td>
                                                              <%end if%>
                                                              
                                                            
                                                              
                                                             
                                                              
                                                            <td><%=rsTabellaDomande("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaDomande("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare la domanda?');"  href="cancella_domanda.asp?Verifica=0&classe=<%=classe%>&cod=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>" title="Cancella">
                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     

                 <% f=f+1
				  '  p=p+1
				  n=n+1
				    rsTabellaDomande.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->  
                                    
                                    
                                    
                                    
                                  
                                        
                              </div>
                              
                                <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                        <div class="tab-pane fade in active" id="profileNodi<%=k%><%=p%>">
                              
                                     <%else%>
                                          <div class="tab-pane fade" id="profileNodi<%=k%><%=p%>">
                              
                                     <%end if%>
                              
                                  <!-- inizio blocco nodi -->  
                                  
                               
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaNodi.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>"> 
                                               Apri tutti i nodi: N(<%= numrsNodi2 &") Pt(" & numrsNodi2  & ") Pb("& round( numrsNodi2/numrsNodi,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                           <% if not rsTabellaNodi.eof then %>
                                                        <tr>
                                                            <th>Nodi</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                           
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                  
                     
                    
                     <% Sottoparagrafo=""
					' p=0
					
					
					
		     do while not rsTabellaNodi.EOF  
			   if StrComp(Sottoparagrafo, rsTabellaNodi("SotPar")) <> 0 then
			   Sottoparagrafo=rsTabellaNodi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaNodi("SotPar")%></b></center></td></th>  
			 <%end if%> 
                                                    
                                                        <tr>
                                                                                                                       
                                                            
                                                             <%if rsTabellaNodi("Segnalata")=1 then%>
                                                   <td><a  style="color:red" title="Apri il nodo"  href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaNodi("Voto")%></td>
                                                             
                                                             <%else%>
                                                       
                                                             
                                                             <td><a title="Apri il nodo"   href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>
                                                           
                                                             <td><%=rsTabellaNodi("Voto")%></td> 
                                                             
                                                              <%end if%>
                                                              
                                                              
                                                            <td><%=rsTabellaNodi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaNodi("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare il nodo?');"  href="../cNodi/cancella_nodo.asp?cla=<%=d%>&cod=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&CodiceTest=<%=rsTabellaNodi("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>">
                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     
                 <% f=f+1
				    rsTabellaNodi.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->                               </div> <!-- fine profile nodi-->
                                                           
                            </div><!-- fine MyTabContent2-->

                          </div><!-- fine collapse(treuno)-->
                        </div> <!-- fine accordino group-- da Descrizione capitolo in giù >-->
                         <%end if %> <!--if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0) then-->
                
                
                
                
                
                         <% p=p+1
						   rsTabellaParagrafi.movenext()
						   Loop
						%>  
                        
                        
                        
                        
                    </div><!-- fine accordion inner-->
                  </div>
                </div> <!--  fine accordion group uno per ogni capitolo-->
       <%compiti=compiti+1  %>       
     <% end if  ' if numrsFrasi<>0%>
			
			<% k=k+1
			   rsTabellaModuli.movenext()
			   Loop
			%>    
            
            
             
            <% if compiti=0 then %>
            <span class="alert-error"><h5>Nessun compito inserito nel periodo dal 
            <%response.write(cdate(DataClaq)&" al ")%>
            <%response.write(cdate(DataClaq2))%>
            </h5></span>
         <!--   <ul class="tiles">
			 
            <li class="blue">
								<a href="#"><span class='nopadding'><h5>NOn ci sono compiti</h5>
                                <span class='name'><i class="icon-twitter"></i><span class="right">1min ago</span></span></a>
							</li>
            </ul>-->
            <%
			 
			end if
			%>
 
 
 
 
   
 
 
               </div> <!--<div class="bs-docs-example"> fino blocco compiti -->
		
		   
        <p> <span class="invisible">
	   <a id="msg" href="#modal-4" role="button" data-notify-time="3000" class="btn notify" data-notify-title="Utilizza la Guida!" data-notify-message="INFORMAZIONI ALLA TUA DESTRA ">
	   
	   </a></span>
         
       
        
        
        
         
		 <!-- #include file = "../include/colora_pagina.asp" -->
        
        
         
        
        
        
	</body>
    <br><br><br><br><hr>
      <!-- #include file = "../include/footer.asp" --> 
      
        <!-- #include file = "../cGuide/g_quaderno.asp" --> 
      
 
      
 <script language="javascript" type="text/javascript">
function cancella_avviso() {
	
	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {  
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();	
	 }
}
   
   
 function aggiornaStud() {
	 // alert (DataClaq);
	 var DataClaq=document.dati.txtData.value;
	 var DataClaq2=document.dati.txtData2.value;
	// alert (DataClaq);
	 // alert (DataClaq2);
		with (document.dati) { 
		 
		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=Session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daForm=1";
		 else
		   document.dati.action = "?divid=<%=Session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daForm=1";
	
	    }
		document.dati.submit();		
}
 

</script>

<script type="text/javascript">
	

		 
$(window).load(function () {
	   
	   $('#<%=box_apri%>').click();
	   $('#<%=box_apri1%>').click();
	    $('#<%=box_apri2%>').click();
		$('#<%=box_apri3%>').click();
	    $('#<%=box_apri4%>').click();
	   $("body").addClass("theme-"+"<%=stile%>").attr("data-theme","theme-"+"<%=stile%>");
  
  
	 
	  // event.stopPropagation();
	    
	});
	

/*$(".red").click(function(event){
   
   // alert("Hai cliccato sull'Elemento");
	document.location = "script/aggiorna_stile.asp?stile=red"
});
*/	
	
</script>

 
	</html>