<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Simula Navigazione Rete della Vita</title>   
   
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

  
<script language="JavaScript">
 var contSi,contNo,Autista,Destinazione,Carburante,Luogo,Strada,Strada_OK,Strada_KO,Cespugli,Lupo,Cestino,Motivato,Testata,v;
 var plurale,plurale1,volere,dovere,fare,avere,raggiungere,scegliere,avvicinarsi,avvicinarsi1,avvicinarsi2,allontanarsi,allontanarsi1,scontrarsi,continuare,ti_vi;

/*
function preloadImages(urls) {
  var img = new Array();
  for (var i=0; i<urls.length; i++) {
    img[i] = new Image();
    img[i].src = urls[i];
	//alert (img[i].src);
  }
}

window.onload = function() {
  var img = new Array("../U-ECDL/img/M_Navigazione/Intro.gif","../U-ECDL/img/M_Navigazione/cestino.gif", "../U-ECDL/img/M_Navigazione/manifestare1.jpg","../U-ECDL/img/M_Navigazione/no_sostegno.gif","../U-ECDL/img/M_Navigazione/lupo1.gif","../U-ECDL/img/M_Navigazione/sostegno1.gif","../U-ECDL/img/M_Navigazione/cestino.gif","../U-ECDL/img/M_Navigazione/coerenza.gif");
  preloadImages(img);
}
*/




cont=0;
immagini = new Array(10); 
immagini[0]="../../U-ECDL/img/M_Navigazione/Intro.gif";
immagini[1]="../../U-ECDL/img/M_Navigazione/cestino.gif";
immagini[2]="../../U-ECDL/img/M_Navigazione/manifestare1.jpg";
immagini[3]="../../U-ECDL/img/M_Navigazione/no_sostegno.gif";
immagini[4]="../../U-ECDL/img/M_Navigazione/lupo1.gif";
immagini[5]="../../U-ECDL/img/M_Navigazione/sostegno1.gif";
immagini[6]="../../U-ECDL/img/M_Navigazione/cestino.gif";
immagini[7]="../../U-ECDL/img/M_Navigazione/coerenza.gif";
immagini[8]="../../U-ECDL/img/M_Navigazione/explorer.jpg";
/*
immagini[0]= new Image();
immagini[1]= new Image();
immagini[2]= new Image();
immagini[3]= new Image();
immagini[4]= new Image();
immagini[5]= new Image();
immagini[6]= new Image();
immagini[7]= new Image();  
  
immagini[0].src=immagini1[0];
immagini[1].src=immagini1[1];
immagini[2].src=immagini1[2];
immagini[3].src=immagini1[3];
immagini[4].src=immagini1[4];
immagini[5].src=immagini1[5];
immagini[6].src=immagini1[6];
immagini[7].src=immagini1[7];

*/

  
function mostra_n(cont) {
  
	document["situazione"].src = immagini[cont];  
}

 

function popup_sx(pagina) 
{ 
window.open(pagina,'','width=800,height=650, left=200,top=100,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
} 
function popup_dx(pagina) 
{ 
window.open(pagina,'','width=800,height=650, left=800,top=100,resizable=yes,toolbar=no,scrollbars=no,status=no') ;
}
 
 

function inizio() {
     Autista=document.parametri.txtAutista.value.toUpperCase();
	 Destinazione=document.parametri.txtDestinazione.value.toUpperCase();
     distanza = parseInt(document.parametri.txtDistanza.value.toUpperCase());
	 Carburante=document.parametri.txtCarburante.value.toUpperCase();
	 Luogo=document.parametri.txtLuogo.value.toUpperCase();
	 Strada=document.parametri.txtStrada.value.toUpperCase();
	 Strada_OK=document.parametri.txtStrada_OK.value.toUpperCase();
	 Strada_KO=document.parametri.txtStrada_KO.value.toUpperCase();
	 Cespugli=document.parametri.txtCespugli.value.toUpperCase();
	 Lupo=document.parametri.txtLupo.value.toUpperCase();
	 Cestino=document.parametri.txtCestino.value.toUpperCase();
     contSi=0;
     contNo=0;
	 Motivato=0;Testata=0;
	 
	  plurale=Autista.search(/ e /i); //se è presente e oppure E è >0
     plurale1=Autista.search(","); //faccio mettere ; per indicare il prurale
 
	 if ((plurale == -1) && (plurale1 == -1)){
	     volere="vuoi";
		 raggiungere="raggiungerai";
		 avere="hai";
		 scegliere="scegli";
		 avvicinarsi="ti avvicina";
		 avvicinarsi1="avvicinarti";
		 avvicinarsi2="avvicinarsi";
		 allontanarsi="ti allontana";
		 allontanarsi1="ti sei allontanato troppo hai";
		 scontrarsi="e ti sei scontrato";
		 continuare="continua";
		 fare="ci sei quasi fai";  
		 dovere="devi";
		 ti_vi="ti";
	  }  
	  
	  else 
      { 
 
         volere="volete";
		 raggiungere="raggiungerete";
		 avere="avete";
		 scegliere="scegliete";
		 avvicinarsi="vi avvicina";
		 avvicinarsi2="avvicinarsi";
		 avvicinarsi1="avvicinarvi";
		 allontanarsi="vi allontana";
		 allontanarsi1="vi siete allontanati troppo avete";
		 scontrarsi="e vi siete scontrati";
		 continuare="continuate";
		 fare="ci siete quasi fate";
	     dovere="dovete";
	     ti_vi="vi";
	   } 
	 
	 
	 
     document.parametri.Storia.value="Distanza Iniziale dall'obiettivo = (" + distanza + ")"; 
	 document.parametri.Storia.value=document.parametri.Storia.value+"\n " + Autista + " "+ volere +"  raggiungere " + Destinazione+" ?";
     document.parametri.BSx.value='No';
     document.parametri.BDx.value='Si';
	 mostra_n(0);
	 
	 
}

function bottoneSx() {
	  
if (Motivato == 0 )
{
   document.parametri.Storia.value=document.parametri.Storia.value+ " NO! \n\n Mancando  " + Carburante + " per raggiungere "+Destinazione+" , "+Autista +" nel contesto " + Luogo + " non " + raggiungere +" l'obiettivo ! ";
   document.parametri.Storia.value=document.parametri.Storia.value+"\n\n " + Autista + " " + volere +" raggiungere " + Destinazione+" ?";
   mostra_n(2);
 // per posizionarsi alla fine della text area 
 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
  
}
else
    {
    if ((contSi-contNo)==(-distanza)) { 	
    document.parametri.Storia.value = document.parametri.Storia.value + "\n\n :-(  " + Autista + "  " + allontanarsi1 + " scelto la strada chiusa  " + Strada_KO + " " + scontrarsi + " con " + Lupo +" \n ";
 	document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n  " + Autista  + "  per risolvere la situazione " + dovere + "  abbandonare  " + Cestino  + " cosi' da " + avvicinarsi1 + " a " + Destinazione + "  ";  
	document.parametri.BSx.value="Trattieni il cestino";
	document.parametri.BDx.value="Lascia andare il cestino";
	mostra_n(4);
	popup_sx("../../U-ECDL/img/paginaTopolinoTontolino.htm");     
	//popup_sx("../../U-ECDL/img/paginaNavigazioneTontolina.htm");     
    Testata=1;
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;	   
		       
    }
    else {
    	contNo++;
       document.parametri.Storia.value=document.parametri.Storia.value + "\n\n  ATTENZIONE  " + Autista + "  la scelta  "+ Strada_KO+" " +allontanarsi +" dall'obiettivo! ";
	   
	   document.parametri.Storia.value=document.parametri.Storia.value + "\n\n "+ Cespugli +" "+ ti_vi + " segnalano il pericolo ! ";
	   
        document.parametri.Storia.value=document.parametri.Storia.value + "\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
		mostra_n(3);
		 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;

        
    }	
    document.parametri.Storia.value=document.parametri.Storia.value + "\n\n  " + Autista + "  quale  "+Strada+" " + scegliere +" ?  ";
    }	  
	
 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;

		
}

function bottoneDx() {
	  
		if (Motivato==1)  
{
		  if ((contSi-contNo)==distanza) {
			   document.parametri.Storia.value=document.parametri.Storia.value + "\n :-) COMPLIMENTI  " + Autista + "  " + avere +" raggiunto  "+ Destinazione+ "  !";
			   popup_dx("../../U-ECDL/img/paginaTopolinoVolpino.htm"); 
			//   popup_dx("../../U-ECDL/img/paginaNavigazioneVolpina.htm");
			         
			   mostra_n(7);
	 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		  }
		  else
		  {
		  if (Testata==1)
		  {
		     Testata=0;
			 document.parametri.BSx.value=Strada_KO;
	         document.parametri.BDx.value=Strada_OK;
			 mostra_n(6);
		  }
		  else
		  {
		  mostra_n(5);
		  contSi++;
		  document.parametri.Storia.value=document.parametri.Storia.value +  "\n\n  " + Autista + "  la scelta  " + Strada_OK + "  ti avvicina a  " + Destinazione+ "  continua cosi' !  ";
		  document.parametri.Storia.value=document.parametri.Storia.value + "\nDistanza attuale dall'obiettivo = ("+ (distanza-(contSi-contNo))+")";
		   document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
	
				  if (contSi-contNo ==(distanza))
				  {
					 document.parametri.Storia.value=document.parametri.Storia.value +  "\n Coraggio " + fare + " l'ultimo passo ! '";
				 document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
				  }
				  document.parametri.Storia.value=document.parametri.Storia.value + "\n\n " + Autista + "  quale  " + Strada + " " +  scegliere+ " ? '";
				  
		  }	
		 /* document.parametri.Storia.value=document.parametri.Storia.value + "\n\n "+ Autista + "  quale  "+Strada+"   scegli ? '";*/
		  document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;
		   }	  
 }
 
 else
    {
	 mostra_n(8);
	document.parametri.Storia.value=document.parametri.Storia.value + "SI!  \n\n " + Autista + "  quale '" + Strada + " " + scegliere + " ? '";
	document.parametri.Storia.scrollTop=document.parametri.Storia.scrollHeight;

    document.parametri.BSx.value=Strada_KO;
	document.parametri.BDx.value=Strada_OK;
    Motivato=1;
	 }
}


</script>

   
</head>

<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
		
 

 
    Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
  						
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    <%
	  CodiceMetafora = Request.QueryString("CodiceMetafora")
  'CodiceTest = Request.QueryString("CodiceTest")
  Num = Request.QueryString("Num")
  Num=Num+1
	%>
    
   
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Simul@azione </h1> 
                    
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
  
QuerySQL="SELECT Tit, ID_Paragrafo,Cognome, CodiceMetafora, ID_Mod, Autista, Destinazione, Carburante, Luogo, Strada, Strada_OK, Strada_KO, Cespugli, Lupo,Cestino,Distanza, In_Quiz,Posizione,Cartella,Pi,Pf,Data,Cartella " &_
" From Elenco_Metafore_Navigazione" &_
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
	 Codice_Test=rsTabella("ID_Paragrafo")
	 Cartella=rsTabella("Cartella")
%>   
                 
                 
      		             
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                     <form name="parametri"  method="POST" style="width:auto" class="form-vertical">
                     
 
								   
        
				           <div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora N.(<span id="codmet"><%=CodiceMetafora%></span>)
								<%if  (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin") = true)  then%>
                                <input type="button" value="Aggiorna" name="BAggiorna" onClick="aggiornaMetafora();" class="btn">
								<%end if%>
								<%%>
                                </h3>
							</div>
                            
				      <div class="box-content">
 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     
                     <fieldset id="Parametri">
          		  
           
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Autista</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" placeholder="Soggetto protagonista" class="input-xxlarge"  name="txtAutista"  id="txtAutista"  value="<%=rsTabella("Autista")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Destinazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtDestinazione" id="txtDestinazione" value="<%=rsTabella.fields("Destinazione")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Carburante</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtCarburante" id="txtCarburante" value="<%=rsTabella.fields("Carburante")%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Luogo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtLuogo" id="txtLuogo" value="<%=rsTabella.fields("Luogo")%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" placeholder="Comportamento" class="input-xxlarge"   name="txtStrada"  id="txtStrada" value="<%=rsTabella.fields("Strada")%>" >
										</div>
									</div>
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6" checked="true" >
											<input type="text" placeholder="Comportamento adeguato" class="input-xxlarge"  name="txtStrada_OK" id="txtStrada_OK" value="<%=rsTabella.fields("Strada_OK")%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" >  
											<input type="text" placeholder="Comportamento inadeguato" class="input-xxlarge"   name="txtStrada_KO" id="txtStrada_KO" value="<%=rsTabella.fields("Strada_KO")%>">
										</div>
									</div>
                                     
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cespugli</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Segnali di pericolo" class="input-xxlarge"   name="txtCespugli"  id="txtCespugli" value="<%= rsTabella.fields("Cespugli") %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Lupo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" placeholder="Conseguenze negative" class="input-xxlarge"  name="txtLupo" id="txtLupo" value="<%= rsTabella.fields("Lupo") %>" >
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cestino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" placeholder="Attaccamenti da lasciare andare" class="input-xxlarge"  name="txtCestino" id="txtCestino" value="<%= rsTabella.fields("Cestino") %>">
										</div>
									</div>
                                    
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Num. da 1 a 5" class="input-small"  name="txtDistanza" id="txtDistanza" value="<%=rsTabella.fields("Distanza")%>">
										</div>
									</div>
                   </fieldset>                 
                                    
                                    
									
                                    <center> 
                                     <b>Simula </b><br>  
                                    <span id="btnSxDx">
                                    <input type="button" class="btn" name="indietro" value="<< Precedente " onClick="Precedente()" title="Zoom indietro">
                                    <input type="button" class="btn" name="avanti" value="Successiva >> " onClick="Successiva()" title="Zoom avanti"> 
                                    </span><hr>
                                    <span id="idInizio">        
                                      <p>  <input type="button" value="INIZIO" name="BInizia" onClick="inizio()" class="btn"> </p></span>
                                      <span id="btnSxDx">
  									<p>    <input type="button" value="  " name="BSx" onClick="bottoneSx()" class="btn">  <input type="button" value="  " class="btn" name="BDx" onClick="bottoneDx()"></p>
                                    </span>
                                   
                                   <div  id="idImg" class="fileupload-new thumbnail" style="text-align:center"> 
                                    
	<img class="imground" name="situazione" src="../../U-ECDL/img/M_Navigazione/Intro.gif">
	</div><br>
    </center> 
    
   									  <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Narr@azione</b></label>
										<div class="controls">
											<textarea  rows="16" name="Storia" class="input-block-level"> </textarea> 
										</div>
									</div>
                                    
                                    <div id="collapseMail" class="accordion-body">
                                            <div class="accordion-inner">
 <center>

 

<input type="hidden" id="txtDATA"  name="txtDATA" value="<%=rsTabella("Data")%>">
<input type="hidden" id="Pi"  name="Pi" value="<%=Pi%>">
<input type="hidden" id="Pf"  name="Pf" value="<%=Pf%>">
<input type="hidden" id="CodiceMetafora"  name="Pf" value="<%=CodiceMetafora%>">
<input type="hidden" value="<%=Cartella%>" id="cartella">
<input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
<input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
<input type="hidden" value="<%=Modulo%>" id="Modulo">
<input type="hidden" value="<%=Paragrafo%>" id="Paragrafo">
<hr>
 
										   </center>
                     						 </div>                       
										</div>
                                    
                                    </form>
                    
                      
                      
               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> 
                      </div>         
			        </div>
			      </div>
			    </div>
	
                      
                      
                  <%
			  rsTabella.close : Set rsTabella = Nothing  %>
				   
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
          
          
              <script language="javascript" type="text/javascript"> 
function Successiva() {
if (document.getElementById("Pf").value==0)
	{
	   alert("Non ci sono Metafore figlio");
	   return 0;
	}
 else 
	{   
		  var url = "7_carica_metafora_json.asp?tipoMetafora=1&CodiceMetafora="+document.getElementById("Pf").value;
		   //alert(url);
		  var xhttp = new XMLHttpRequest();
		  xhttp.onreadystatechange = function() {
			if (xhttp.readyState == 4 && xhttp.status == 200) {
				var testo = xhttp.responseText;		
				var json = JSON.parse(testo);
				document.getElementById("txtAutista").value=json["soggetto"];
				document.getElementById("txtDestinazione").value=json["obiettivo"];
				document.getElementById("txtCarburante").value=json["motivazione"];
				document.getElementById("txtLuogo").value=json["ambiente"];
				document.getElementById("txtStrada").value=json["comportamento"];
				document.getElementById("txtStrada_KO").value=json["ko"];
				document.getElementById("txtStrada_OK").value=json["ok"];
				document.getElementById("txtCespugli").value=json["feedback"];
				document.getElementById("txtCestino").value=json["eccessi"];
				document.getElementById("txtLupo").value=json["conseguenze"];
				document.getElementById("Pi").value=json["pi"];
				document.getElementById("Pf").value=json["pf"];
				document.getElementById("CodiceMetafora").value=json["codicemetafora"];
				document.getElementById("codmet").innerHTML = json["codicemetafora"] ;
				<%CodiceMetafora=Pf%>
				 
				 
			}
		  };
		  xhttp.open("GET", url, true);
		  xhttp.send();
    }
 }
 
  function Precedente() {
 if (document.getElementById("Pi").value==0)
	{
	   alert("Non ci sono Metafore genitore");
	   return 0;
	}
 else 
	{   
		  var url = "7_carica_metafora_json.asp?tipoMetafora=1&CodiceMetafora="+document.getElementById("Pi").value;
		   //alert(url);
		  var xhttp = new XMLHttpRequest();
		  xhttp.onreadystatechange = function() {
			if (xhttp.readyState == 4 && xhttp.status == 200) {
				var testo = xhttp.responseText;		
				var json = JSON.parse(testo);
				document.getElementById("txtAutista").value=json["soggetto"];
				document.getElementById("txtDestinazione").value=json["obiettivo"];
				document.getElementById("txtCarburante").value=json["motivazione"];
				document.getElementById("txtLuogo").value=json["ambiente"];
				document.getElementById("txtStrada").value=json["comportamento"];
				document.getElementById("txtStrada_KO").value=json["ko"];
				document.getElementById("txtStrada_OK").value=json["ok"];
				document.getElementById("txtCespugli").value=json["feedback"];
				document.getElementById("txtCestino").value=json["eccessi"];
				document.getElementById("txtLupo").value=json["conseguenze"];
				document.getElementById("Pi").value=json["pi"];
				document.getElementById("Pf").value=json["pf"];
				document.getElementById("CodiceMetafora").value=json["codicemetafora"];
				document.getElementById("codmet").innerHTML = json["codicemetafora"] ;
				<%CodiceMetafora=Pf%>
				 
				 
			}
		  };
		  xhttp.open("GET", url, true);
		  xhttp.send();
    }
 }
 
 
 
 
 function aggiornaMetafora(){
 
  
  
//  document.parametri.action = "inserisci_metafora_patente1.asp?daSimulazione=1&CodiceMetafora=="+document.getElementById("CodiceMetafora").value;
  // document.parametri.submit();
   
  
		var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;		  
	 
			txtAutista=document.getElementById("txtAutista").value;
			txtDestinazione=document.getElementById("txtDestinazione").value;
			txtCarburante=document.getElementById("txtCarburante").value;
			txtLuogo=document.getElementById("txtLuogo").value;
			txtStrada=document.getElementById("txtStrada").value;
			txtStrada_OK=document.getElementById("txtStrada_OK").value;
			txtStrada_KO=document.getElementById("txtStrada_KO").value;
			txtCespugli=document.getElementById("txtCespugli").value;
			txtLupo=document.getElementById("txtLupo").value;
			txtCestino=document.getElementById("txtCestino").value;
			txtDistanza=document.getElementById("txtDistanza").value;
		//	txtData=document.getElementById("txtDATA").value;
		//	textarea=document.getElementById("textarea").value; 
			//segnalata=document.getElementById("cb1").checked; 
			//voto=document.getElementById("txtVAl").value;
			dati2="&txtAutista="+txtAutista+"&txtDestinazione="+txtDestinazione+"&txtCarburante="+txtCarburante+"&txtLuogo="+txtLuogo+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtCespugli="+txtCespugli+"&txtCestino="+txtCestino+"&txtLupo="+txtLupo+"&txtDistanza="+txtDistanza+"&daSimulazione=1";
 
	
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
   
 
     
                                
<script language="javascript" type="text/javascript" src="../jsguide/navigazionesimula.js"> </script> 
							 
							       
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         

			 
	</body>

 </html>

