<%@ Language=VBScript %>



<!doctype html>
<html>
<head>
   
   <title>Simula Topolino</title>   
    
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
window.open(pagina,'','width=560,height=390, left=500,top=100,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
} 
 
 

function inizio() {
     Topolino=document.parametri.txtTopolino.value.toUpperCase();
	 Formaggio=document.parametri.txtFormaggio.value.toUpperCase();
     distanza = parseInt(document.parametri.txtDistanza.value.toUpperCase());
	 Fame=document.parametri.txtFame.value.toUpperCase();
	 Labirinto=document.parametri.txtLabirinto.value.toUpperCase();
	 Strada=document.parametri.txtStrada.value.toUpperCase();
	// alert("Strada"+Strada)
	 Strada_OK=document.parametri.txtStradaOK.value.toUpperCase();
	 Strada_KO=document.parametri.txtStradaKO.value.toUpperCase();
	 Testata=document.parametri.txtTestata.value.toUpperCase();
	 
	 plurale=Topolino.search(/ e /i); //se è presente e oppure E è >0
     plurale1=Topolino.search(","); //faccio mettere ; per indicare il prurale
 
	 if ((plurale == -1) && (plurale1 == -1)){
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
	 document.parametri.Storia.value=document.parametri.Storia.value+"\n\n " + Topolino + " " + volere + " raggiungere " + Formaggio+" ?";
     document.parametri.BSx.value='No';
     document.parametri.BDx.value='Si';
	 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
}

function bottoneSx() {
	  
if (Motivato == 0 )
{
   document.parametri.Storia.value=document.parametri.Storia.value+ " NO! \n\n Mancando " + Fame + " per raggiungere "+Formaggio+" , "+Topolino +" nel contesto " + Labirinto + " non " +  raggiungere + " l'obiettivo ! ";
   document.parametri.Storia.value=document.parametri.Storia.value+"\n\n " + Topolino + " " + volere +" raggiungere " + Formaggio+" ?";
   document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
}
else
    {
    if ((contSi-contNo)==(-distanza)) { 	
    document.parametri.Storia.value=document.parametri.Storia.value +"\n\n :-(  " + Topolino + " " + allontanarsi1 + " scelto la strada chiusa  "+Strada_KO+ " " + scontrarsi + " con " + Testata+" \n ";
	 popup("../../U-ECDL/img/PaginaTopolinoTontolino.htm");
	 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;        
    }
    else {
    	contNo++;
       document.parametri.Storia.value=document.parametri.Storia.value + "\n\nATTENZIONE  " + Topolino + "  la scelta  "+ Strada_KO+" " + allontanarsi +" da  "+ Formaggio;
        document.parametri.Storia.value=document.parametri.Storia.value + "\n\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
		document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
        
    }	
    document.parametri.Storia.value=document.parametri.Storia.value + "\n\n  " + Topolino + "  quale  "+Strada+" " + scegliere +" ?  ";
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
    }	  
		
}

function bottoneDx() {
	  
		
if (Motivato==1)  {
      if ((contSi-contNo)==distanza) {
           document.parametri.Storia.value=document.parametri.Storia.value + "\n :-) COMPLIMENTI  " + Topolino + " "+ avere + " raggiunto " + Formaggio+ "!!!";
		   popup("../../U-ECDL/img/PaginaTopolinoVolpino.htm");       
		 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;

		  }
      else
      {
      contSi++;
      document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n  "+Topolino+"  la scelta  "+Strada_OK+" " +avvicinarsi +" a  "+ Formaggio+ "  " + continuare + " cosi' !  ";
      document.parametri.Storia.value=document.parametri.Storia.value + "\n\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
	  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		  if (contSi-contNo ==(distanza))
		  {
			 document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n Coraggio " + fare + " l'ultimo passo ! '";
		  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		  }
		  document.parametri.Storia.value=document.parametri.Storia.value + "\n\n "+ Topolino + "  quale  "+Strada+" " +  scegliere+" ? '";
		  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
      }
    }
    else{
	document.parametri.Storia.value=document.parametri.Storia.value + "SI!  \n\n " + Topolino + "  quale '"+Strada+" "+ scegliere +" ? '";
    document.parametri.BSx.value=Strada_KO;
	document.parametri.BDx.value=Strada_OK;
    Motivato=1;
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
	    }
}


</script>


%>
<%
  Response.Buffer = true
  On Error Resume Next  
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
						<h1> <i class="icon-comments"></i> Simula la realt&agrave; con la metafora </h1> 
                    
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
  QuerySQL="SELECT Tit, ID_Paragrafo, Cognome, CodiceMetafora, ID_Mod, Topolino, Formaggio, Fame, Labirinto, Strada, Strada_OK, Strada_KO, Testata, Distanza, In_Quiz,Posizione,Titolo, Posizione,Pi,Pf " &_
" From Elenco_Metafore_topolino " &_
" Where CodiceMetafora =" & CodiceMetafora & "" 
 ' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logSimulazione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close  
	'response.write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	 Pi=rsTabella("Pi") ' codice della metafora precedente
	 Pf=rsTabella("Pf") ' ' codice della metafora seguente
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
                   
               <form action="inserisci_metafora_topolino1.asp?daSimulazione=1&CodiceMetafora=<%=CodiceMetafora%>" name="parametri"  method="POST" class="form-vertical">
						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora N.(<%=CodiceMetafora%>)
                                <input type="submit" value="Aggiorna" name="BAggiorna" class="btn">
                                </h3>
							</div>
                            <fieldset id="Parametri">
							<div class="box-content">
							
							
     
                              	
 
                                  <div class="control-group">
										<label for="textfield" class="control-label"><b>Topolino</b></label>
										<div class="controls">
                                            <input type="text" placeholder="Soggetto protagonista" class="input-xxlarge"  name="txtTopolino"  value="<%=rsTabella("Topolino")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Formaggio</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtFormaggio"  value="<%=rsTabella("Formaggio")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Fame</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtFame"  value="<%=rsTabella("Fame")%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Labirinto</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtLabirinto"  value="<%=rsTabella("Labirinto")%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                         
								<input type="text" placeholder="Obiettivo" class="input-xxlarge"  name="txtStrada"  value="<%=rsTabella("Strada")%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada OK</b></label>
										<div class="controls">
                                         
						<input type="text" placeholder="Strategia vincente" class="input-xxlarge"  name="txtStradaOK"  value="<%=rsTabella("Strada_OK")%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada KO</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge"  name="txtStradaKO"  value="<%=rsTabella("Strada_KO")%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Testata</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Conseguenze della strategia perdente" class="input-xxlarge"  name="txtTestata"  value="<%=rsTabella("Testata")%>">
										</div>
									</div>
                                    
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Num. da 1 a 5" class="input-small"  name="txtDistanza"  value="<%=rsTabella("Distanza")%>">
										</div>
									</div>
                                </fieldset>    
                                    
                                    
									
								
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
		<center>	<span id="idInizio">						 
    <p>  <input type="button" value="INIZIO" name="BInizia" class="btn-primary" onClick="inizio()"> </p>
    </span>
    <span id="btnSxDx">	
  <p>    <input type="button" value="  " name="BSx" onClick="bottoneSx()" class="btn">
         <input type="button" value="  " name="BDx" onClick="bottoneDx()" class="btn"></p>
         </span> <!--Definisce i due bottoni del form --></center>
	</div></p> 
	<!--<input type="button" value="clicca" onClick="<%'call ciao1()%>">-->
	 
								<div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Narr@azione</b></label>
										<div class="controls">
											<textarea name="Storia" id="textarea" rows="15" class="input-block-level"><%=					Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
								 
									</div>
                                    
								<div id="collapseMail" class="accordion-body">
                                            <div class="accordion-inner">
 <center>
 <b>Simula </b><br> 
 
<span id="btnSxDx">
<input type="button" class="btn" name="indietro" value="<< Precedente " onClick="Precedente()" title="Zoom indietro">
<input type="button" class="btn" name="avanti" value="Successiva >> " onClick="Successiva()" title="Zoom avanti"> 
</span>
<input type="hidden"  name="Pi" value="<%=Pi%>">
<input type="hidden"   name="Pf" value="<%=Pf%>">
<hr>
 
										   </center>
                     						 </div>                       
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
      
      
     
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   %>
    <script language="javascript" type="text/javascript"> 
 function Successiva() {
	 //window.alert(dati.Pf.value);
//location.href="../home.asp"
//location.href=window.history.back();
if (parametri.Pf.value==0)
	{
	   alert("Non ci sono Metafore figlio");
	   return 0;
	}
 else 
	{
	   document.parametri.action = "6_simula_metafora_topolino.asp?CodiceMetafora=<%=Pf%>";
		document.parametri.submit();	
    }
 }
 
  function Precedente() {
 // window.alert(dati.Pi.value);
  if (parametri.Pi.value==0)
	{
	   
	   alert("Non ci sono Metafore genitore");
	   return 0;
	}
 else
  
	{
	   document.parametri.action = "6_simula_metafora_topolino.asp?CodiceMetafora=<%=Pi%>";
		document.parametri.submit();	
    }
//location.href="../home.asp"
//location.href=window.history.back();
 }
		
 </script>   
            
      
  <script src="../../../guida/docs/lib/bootstrap/js/bootstrap-dropdown.js"></script>
    <script src="../../../guida/docs/lib/google-code-prettify/prettify.js"></script>

    <script src="../../../guida/js/jquery.pageguide.js"></script>
    <script language="javascript">
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
   
 
     
                                
<script language="javascript" type="text/javascript" src="../jsguide/topolinosimula.js"> </script> 
							 
							 
     
      
      
		</div> <!--fine main-->
        </div>
        
         

			 
	</body>

 </html>

