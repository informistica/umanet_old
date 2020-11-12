<%@ Language=VBScript %>



<!doctype html>
<html>
<head>
   
   <title>Sviluppa Topolino</title>   
    
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
	 

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
	
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

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

 
 


%>
<%
  Response.Buffer = true
 ' On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
   ' if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
   %>
	<!-- <BODY onLoad="showText2();"> </BODY>-->
  <% 'else %>
  %>
    
    <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" >

  <% 'end if %>

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  
 ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
      'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBCopiatestonline")
    
	
	
    
'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
'              "DBQ=" & Server.MapPath("../database/Copiatestonline.mdb")

 

   
%>
	<div id="navigation">
     
        <% 
		
 
 
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set objFSO = CreateObject("Scripting.FileSystemObject")
    			
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo") ' TOPOLINO ED OBIETTIVI
 ' Modulo=Request.QueryString("Modulo")
 ' Cartella=Request.QueryString("Cartella")
  Num = cint(Request.QueryString("Num"))
  Num=Num+1
   daTopolino=Request.QueryString("daTopolino")
  CodiceMetafora=Request.QueryString("CodiceMetafora")
   CodiceAllievo=Request.QueryString("CodiceAllievo")
   DATA=Request.QueryString("DATA")
  Collegata=CodiceMetafora
  
  SELECT CASE Request.Form("rdSviluppa")
     CASE "1"
       Li=5
      CASE "2"
       Li=6
	  CASE "3"
       Li=7
     CASE "4"
       Li=8
	 CASE "5"
       Li=9
	 CASE "6"
       Li=10
	 CASE "7"
       Li=11
	 CASE "8"
       Li=12
     CASE ELSE
     	Li=0
     END SELECT  
  'Sviluppa=Request.QueryString("Sviluppa") ' è settato se sono chiamata da sviluppa metafora devo inserire e linkare con l codice della chiamante
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   %>   
   <!-- #include file = "../var_globali.inc" --> 
     	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" 
  		<!-- #include file = "controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          

 <% QuerySQL="Select * from M_Topolino where CodiceMetafora=" & cint(CodiceMetafora)& ";"
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	  Cartella=rsTabella.fields("Cartella")
	  Modulo=rsTabella.fields("Id_Mod")
	  Codice_Test=rsTabella.fields("Id_Arg")
	  ThreadParent=rsTabella.fields("ThreadParent") ' come nel forum
%>

		 
 
       
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Sviluppa metafora </h1> 
                    
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
  QuerySQL="SELECT Tit, ID_Paragrafo, Cognome, CodiceMetafora, ID_Mod, Topolino, Formaggio, Fame, Labirinto, Strada, Strada_OK, Strada_KO, Testata, Distanza,In_Quiz,Titolo, Posizione " &_
" From Elenco_Metafore_topolino " &_
" Where CodiceMetafora =" & CodiceMetafora & "" 
 ' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\wwwroot\Anno_2010-2011_ITC\logSimulazione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close  
	'response.write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'response.write("ciao")
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
								<h3><i class="icon-th-list"></i>  Metafora </h3>
							</div>
							<div class="box-content">
							
							
     
                              	<form  action="inserisci_metafora_topolino1.asp?CodiceMetafora=<%=CodiceMetafora%>&prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&daTopolino=<%=daTopolino%>&Li=<%=Li%>&ThreadParent=<%=ThreadParent%>&daSviluppa=1"  name="parametri"  method="POST" class="form-vertical">
								
								 <input type="hidden" value="<%=Cartella%>" id="cartella">
								  <input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
								   <input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
								   <input type="hidden" value="<%=ThreadParent%>" id="ThreadParent">
								   <input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
								      <input type="hidden" value="<%=Modulo%>" id="Modulo">
									   <input type="hidden" value="<%=Paragrafo%>" id="Paragrafo">
								
 
                                  <div class="control-group">
										<label for="textfield" class="control-label"><b>Topolino</b></label>
										<div class="controls">
                                            <input type="text" placeholder="Soggetto protagonista" class="input-xxlarge" name="txtTopolino" maxlength="148" id="txtTopolino"  value="<%=rsTabella("Topolino")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Formaggio</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtFormaggio" maxlength="148" id="txtFormaggio"  value="<%=rsTabella(Li)%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Fame</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  maxlength="148" name="txtFame" id="txtFame"  value="<%=rsTabella("Fame")& "(?)"%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Labirinto</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  maxlength="148" name="txtLabirinto" id="txtLabirinto"  value="<%=rsTabella("Labirinto") & "(?)"%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                         
								<input type="text" placeholder="Obiettivo" class="input-xxlarge"  name="txtStrada" id="txtStrada" maxlength="148"  value="<%=rsTabella("Strada")& "(?)"%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada OK</b></label>
										<div class="controls">
                                         
						<input type="text" placeholder="Strategia vincente" class="input-xxlarge"  name="txtStrada_ok" id="txtStrada_OK" maxlength="148"  value="<%="(?)"%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada KO</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge" maxlength="148"  name="txtStrada_ko" id="txtStrada_KO"  value="<%=rsTabella("Strada_KO")& "(?)"%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Testata</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Conseguenze della strategia perdente" class="input-xxlarge" maxlength="148"  name="txtTestata"  id="txtTestata"  value="<%=rsTabella("Testata")& "(?)"%>">
										</div>
									</div>
                                    
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Num. da 1 a 5" class="input-small"  name="txtDistanza" maxlength="148"  id="txtDistanza"  value="<%=rsTabella("Distanza")%>">
										</div>
									</div>
                                    
  <%
  
  	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&CodiceMetafora&".txt" 'per il server on line
				 url=Replace(url,"\","/")
	 		   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
			'	on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
				'sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
  %>                                  
                                    
									
								
                                 <div class="accordion" id="accordion3">
									<div class="accordion-group">      
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseMail"><center>
												
                                                <i class="icon-edit" title="Sviluppa"></i>
                                                </center>
											</a>
										</div>
										<div id="collapseMail" class="accordion-body collapse">
											<div class="accordion-inner">
 
 
										   
                     						 </div>                       
										</div>
                                     </div>  
                                     
 
                                    
                                    <div class="form-actions">
	 
	</div></p> 
	<!--<input type="button" value="clicca" onClick="<%'call ciao1()%>">-->
	 
								<div class="control-group">
										<label for="textarea" class="control-label"><b>Narr@azione</b></label>
										<div class="controls">
											<textarea maxlength="910" name="S1" id="textarea" rows="15" class="input-block-level"><%=					Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
								 
                                  
									
									 
										<button type="button" onclick="inserisci_metafore(0)" class="btn btn-primary" name="b1">Invia</button>
                                                                        
										<button type="button" class="btn btn-primary" onClick="copia_testo();" name="b1">Copia testo</button>								 									 
                                   
									
									
									</div>
								
                                
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
            
      
		</div> <!--fine main-->
        </div>
        
         
<script language="javascript">


// JavaScript Document
function inserisci_metafore(tipo){
var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  ThreadParent=document.getElementById("ThreadParent").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;		  
	   	
		  txtTopolino=document.getElementById("txtTopolino").value;
		  txtFormaggio=document.getElementById("txtFormaggio").value;
		  txtFame=document.getElementById("txtFame").value;
		  txtLabirinto=document.getElementById("txtLabirinto").value;
		  txtStrada=document.getElementById("txtStrada").value;
		  txtStrada_OK=document.getElementById("txtStrada_OK").value;
		  txtStrada_KO=document.getElementById("txtStrada_KO").value;
		  txtTestata=document.getElementById("txtTestata").value;
		  txtDistanza=document.getElementById("txtDistanza").value;
		 
		  textarea=document.getElementById("textarea").value;
		
		  dati2="&CodiceMetafora="+CodiceMetafora+"&ThreadParent="+ThreadParent+"&txtTopolino="+txtTopolino+"&txtFormaggio="+txtFormaggio+"&txtFame="+txtFame+"&txtLabirinto="+txtLabirinto+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtTestata="+txtTestata+"&txtDistanza="+txtDistanza+"&S1="+textarea;		 
		 
	
	dati="cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&Codice_Test="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo; 
      var url = "7_sviluppa_metafora_ajax.asp?"+dati+dati2;			   
	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
	  if (xhttp.readyState == 4 && xhttp.status == 200) {
		  var testo = xhttp.responseText;		
						var testo = xhttp.responseText;	
						testoJSON=JSON.parse(testo);
						stato=testoJSON["stato"];
						alert(stato);
						CodiceMetafora=testoJSON["id"];
						if (CodiceMetafora != 0)
							 window.location.href = "sintesi_metafore.asp?cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&CodiceTest="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&CodiceMetafora="+ThreadParent;
		 
					 	
			
	  }
	};
	xhttp.open("GET", url, true);
	xhttp.send();	
		
}
	
	



function copia_testo(){
	
	
	
	var testo;
	 
	testo=document.parametri.txtTopolino.value + " "+ document.parametri.txtFormaggio.value + " "+ document.parametri.txtFame.value + " "+document.parametri.txtStrada.value +  " "+document.parametri.txtStrada_ok.value + " "+document.parametri.txtStrada_ko.value + " "+document.parametri.txtTestata.value ;
	document.parametri.S1.value=testo;
	 } 
	 
</script>
			 
	</body>

 </html>

