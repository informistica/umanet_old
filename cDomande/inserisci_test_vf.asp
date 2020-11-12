<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
      <meta charset="UTF-8">  
   <title>Inserisci test Vero Falso</title>   
     <meta https-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	  <meta charset="UTF-8">
	
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
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	 <script src="../../js/eak_app_dem.min.js"></script>

	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

  

 <script type="text/javascript" src="../js/controlla_checkbox.js"></script>
 <script language="javascript" type="text/javascript" >
 function validate2() {

 if (frmDocument.txtDomanda.value=="")
	{
	   alert("Non hai inserito la Domanda.");
	   frmDocument.txtDomanda.setfocus();
	   return 0;
	}
/*else
 if (frmDocument.txtR1.value=="")
	{
	   alert("Non hai inserito la risposta1.");
	   frmDocument.txtR1.setfocus();
	   return 0;
	}
 else
  if (frmDocument.txtR2.value=="")
	{
	   alert("Non hai inserito la risposta2.");
	   frmDocument.txtR2.setfocus();
	   return 0;
	}else
	 if (frmDocument.txtR3.value=="")
	{
	   alert("Non hai inserito la risposta3.");
	   frmDocument.txtR3.setfocus();
	   return 0;
	}else
	 if (frmDocument.txtR4.value=="")
	{
	   alert("Non hai inserito la risposta4.");
	   frmDocument.txtR4.setfocus();
	   return 0;
	}else*/
	 else
	 if (frmDocument.txtRE.value=="")
	{
	   alert("Non hai inserito il numero della risposta esatta.");
	   frmDocument.txtRE.setfocus();
	   return 0;
	}else
	 if (frmDocument.S1.value=="")
	{
	   alert("Non hai inserito la spiegazione.");
	   frmDocument.S1.setfocus();
	   return 0;
	}
	else
	
	{
	    document.frmDocument.action = "inserisci_test_vf1.asp?Tipo=<%=Tipo%>&by_UPLOAD=<%=by_UPLOAD%>&by_UECDL=<%=by_UECDL%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Multiple=<%=Multiple%>&AggRisDomanda=1&Img=0&Id_Domanda=<%=Id_Domanda%>&contDomande=<%=contDomande%>";
		document.frmDocument.submit();
		
	 
    }
	
}
 </script>
 <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione ? scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
 <script language="javascript" type="text/javascript"> 
function showText4() {window.alert("Non adesso grazie! Troppo tardi o troppo presto !")
location.href="../home.asp"
 
 }
 </script>
   
</head>

 <!-- #include file = "../service/formatta_caratteri_speciali.asp" -->
       
<%Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function
if not(strcomp(Scadenza,"gg/mm/aaaa")=0 or (Scadenza="")) then
' se non ? impostata la scadenza la pongo uguale ad oggi per evitare errori
      Scadenza=Cdate(Request.QueryString("Scadenza"))
   else
      Scadenza=gira_data()
end if


      Data = gira_data()
	   if Datediff("d",Scadenza,Data)>0 then  
		' response.write(Scadenza & " (1) " & Data & " (2) " & Datediff("d",Scadenza,Data) )%> 
        <BODY onLoad="showText4();"> </BODY>
       <%end if%>
  
  
    
    <%' per il controllo della validit? della sessione, se ? scaduta -> nuovo login
	if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  

  <% end if %>
	<div id="navigation">
     
        <% 
	 Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa , 2 vero/falso
  Img=Request.QueryString("Img") ' tipo di domanda 1 se anche Img=1 inserisco l'immagine come quesito 
  by_UPLOAD=Request.QueryString("by_UPLOAD") ' se sono stato chiamato dopo un Upload di una immagine
  Id_Domanda=Request.QueryString("Id_Domanda")	' id della domanda inserita serve per mostrarne le immagini								
  contDomande=Request.QueryString("contDomande") ' incremento per il nome delle immagini multiple per la stessa frase
 
  
  Multiple=Request.QueryString("Multiple") ' vale 1 se devo gestire l'inserimento delle ripsoste multiple
  'Request.Cookies("Dati")("CodiceTest")= Codice_Test
  
  Codice_Test=Request.QueryString("CodiceTest")
  'response.write(Codice_Test)
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito")
  Cartella=Request.QueryString("Cartella")
  CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  predomanda = Request.QueryString("predomanda") 
  ID_Predomanda=Request.QueryString("ID_Predomanda") 
  'response.write("ID_PRE:"&ID_Predomanda)
  id_classe = Request.QueryString("id_classe") 
  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = Request.QueryString("Num")
  Num=Num+1
  by_UECDL=Request.QueryString("by_UECDL")
  Scadenza=Request.QueryString("Scadenza")  
  R1=Request.QueryString("R1")	
  
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <!-- #include file = "../service/formatta_caratteri_speciali.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserisci test V/F </h1> 
                    
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%> : <%=Paragrafo%>
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
                     
                      <% QuerySQL="Select In_Quiz from Allievi where CodiceAllievo='" & CodiceAllievo &"';"
  'response.write(QuerySQL)
  
  '	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2012-2013_2/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\anno_2012-2013_2\logIVF.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 
  In_Quiz=rsTabella(0)
  if session("Admin")=true then
  '??????? inquiz tra un inserimento e il successivo
  end if
  ' aggiunto nella query il Tipo=2 per le multiple
  if CodiceSottopar<>"" then
  QuerySQL="Select Domande.*,Allievi.Cognome, Allievi.Nome " &_
	  " FROM Domande INNER JOIN Allievi ON Domande.Id_Stud = Allievi.CodiceAllievo " &_
	  " where Id_Arg='" & CodiceTest &"' and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "'  and Allievi.In_Quiz=" & In_Quiz & " and VF=1 order by CodiceDomanda,Data,Quesito;"
  
  else
		QuerySQL="Select Domande.*,Allievi.Cognome, Allievi.Nome " &_
	  " FROM Domande INNER JOIN Allievi ON Domande.Id_Stud = Allievi.CodiceAllievo " &_
	  " where Id_Arg='" & CodiceTest &"' and Allievi.In_Quiz=" & In_Quiz & " and VF=1 order by CodiceDomanda,Data,Quesito;"
  
  end if
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  'response.write(QuerySQL)
%>
  
<!--<div id="container">-->
<form method="POST" class='form-vertical'>
 	<div class="control-group">
	<label for="select" class="control-label">Domande gi&agrave; presenti<small> (evita duplicati)</small></label>
										<div class="controls">
      <select name="txtData">
	  <% while not rsTabella.eof %>
		 <option value="<%=rsTabella.fields("Quesito")%>"><%=rsTabella.fields("Quesito") & " ("& rsTabella.fields("Data") &" - "& rsTabella.fields("Cognome") & " " & left(rsTabella.fields("Nome"),1)&".)"%> </option>
         
		 
		 <% rsTabella.movenext()
		    wend%>
	</select>
										</div>
									</div>
 
 <!-- <input type="submit" value="Visualizza" name="B1">  -->
  
</form>

<% 


if strcomp(Img,"1")=0 then ' devo mettere il form per caricare l'immagine, come in 2insersici_frase %>
	 <form class='form-vertical' name="frmDocument" METHOD="Post" ENCTYPE="multipart/form-data">
	<% else%>
    
	<form class='form-vertical' method="POST" name="dati" action="inserisci_test_vf1.asp?by_UECDL=<%=by_UECDL%>&predomanda=<%=predomanda%>&ID_Predomanda=<%=ID_Predomanda%>&Multiple=<%=Multiple%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>" >  
<% end if %>

	
	 
      <% if Tipo="0" then ' cio? domanda con testo semplice
	'response.write(left (Quesito,len(Quesito)))
		 %>
	  <p><input type="text" class="input-xxlarge" name="txtDomanda"  size="135" maxlength="250" value="<%=response.write(left (Quesito,len(Quesito)))%>" ><b> 
		Domanda <a title="Testo lungo per la domanda" href="inserisci_test_vf.asp?Num=<%=Num-1%>&Multiple=<%=Multiple%>&Tipo=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">(+)</a>
		<a title="Immagine come domanda" href="inserisci_test_vf.asp?Num=<%=Num-1%>&Multiple=<%=Multiple%>&Tipo=1&Img=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">(*)</a>
		 
		</b></p> 
	  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
  <% else ' domanda con testoplus o immagine%> 
  <p>
  
 <% if strcomp(by_UPLOAD,"1")=0 then ' se ho gia fatto upload ho anche inseriro domanda quindi disabilito per evitare modifiche che non verrebbero aggiornate
' inoltre prelevo i dati della domanda appena inserita
 			QuerySQL="Select * from Domande where CodiceDomanda="& clng(Id_Domanda)&";"
			 Set rs1 = ConnessioneDB.Execute(QuerySQL) 
			 Quesito=rs1("Quesito")
'			 R1=rs1("Risposta1")
			' R2=rs1("Risposta2")
'			 R3=rs1("Risposta3")
'			 R4=rs1("Risposta4")
			 RE=rs1("RispostaEsatta")
			' devo mostrare l'immagine caricata

%>
   <input type="text" class="input-xxlarge" name="txtDomanda"  size="135" disabled="true" maxlength="250" value="<%=response.write(left (Quesito,len(Quesito)))%>">&nbsp;<b>Domanda<a href="inserisci_test_vf.asp?Multiple=<%=Multiple%>&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">(-)</a></b>
<%else%>
   <input class="input-xxlarge" type="text" name="txtDomanda"  size="135"  maxlength="250" value="<%=response.write(left (Quesito,len(Quesito)))%>">&nbsp;<b>Domanda<a href="inserisci_test_vf.asp?Multiple=<%=Multiple%>&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">(-)</a></b>
<%end if%>
  
  </p>
         <% if strcomp(by_UPLOAD,"1")=0 then
		 ' se provengo da upload ho gia inserito la domanda e ora la devo ripescare per mostrare il form configurato altrimenti sarebbe vuoto
			
			 
			QuerySQL="Select * from Domande_Img where Id_Domanda="& clng(Id_Domanda)&";"
			url= "../Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Domande/Img" ' vuole il percorso relativo della cartella
			url=Replace(url,"\","/")   
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)   
			%><div class="immagini" style="overflow:auto;"><%   
			do while not rsTabella.eof
				   'response.write(url&"/"& rsTabella("Url")&"<b	
				   urlimg=url&"/"& rsTabella("Url") ' aggiungo al percorso il nome del file
				   urldelete=Server.MapPath(homesito)&"/Materie/"&Session("ID_Materia")&"/"&Cartella&"/"&Modulo&"_Domande/Img/"&rsTabella("Url")  ' per cancellare l'immagine.jpg
				   urldelete=Replace(urldelete,"\","/")
				  
				   'response.write("urlimg="&urlimg)%>
				   <p align="center">
				   <img src="<%=urlimg%>" border="1"> <br>
				  <% response.write(rsTabella("Nome"))%><br>
				  <a href="../service/cancella_immagine.asp?by_Domande=1&urldb=<%=rsTabella("Url")%>&urlimg=<%=urldelete%>&CodiceAllievo=<%=Session("CodiceAllievo")%>"><img src="../../img/elimina_small.jpg" width="10" height="10" title="Elimina" onClick="return window.confirm('Vuoi veramente cancellare questa immagine?');"></a></p>
				 <% rsTabella.movenext
		   loop%>
           </div>
		  <br><br><%
		  ' leggo il file di testo per la spiegazione
		  url1=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&Id_Domanda&".txt"
		  url1=Replace(url1,"\","/")
			
			Set objFSO = CreateObject("Scripting.FileSystemObject")  
'			url2="C:\Inetpub\umanetroot\anno_2012-2013\logFile.txt"
'			Set objCreatedFile = objFSO.CreateTextFile(url2, True)
'			objCreatedFile.WriteLine(url1)
'			objCreatedFile.Close 
			
			Set objTextFile = objFSO.OpenTextFile(url1, 1) ' 1= ForReading
			sReadAll = objTextFile.ReadAll
			objTextFile.Close
			' sReadAll="CAZZO"
			' metto il form per aggiunger altre immagini
			 %>
			  <div class="contenuti_sint" style="width:300px">
	  
                 <br>
                 Immagine : <INPUT TYPE="file"  name="flname"  ><BR><br>
                 Nome Img : <input type="text" name="imgname"><br>
                <p> <input type="button"  class="btn-primary" name="btnUpload" value="Carica" onClick="return validate2(1);"></p>
           
               </div> 
			 
         
         
         <%else ' mostro il bottone per il caricamento%>
         <% if strcomp(Img,"1")=0 then ' devo mettere il form per caricare l'immagine, come in 2insersici_frase %>
         	  
            
              <div class="contenuti_sint" style="width:300px">
	  
                 <br>
                 Immagine : <INPUT TYPE="file"  name="flname"  ><BR><br>
                 Nome Img : <input type="text" name="imgname"><br>
                <p> <input type="button" name="btnUpload" class="btn-primary" value="Carica" onClick="return validate2();"></p>
           
               </div>
			<%else ' metto la textarea%>
				<% if strcomp(by_UPLOAD,"1")=0 then ' se ho gia fatto upload ho anche inseriro domanda quindi disabilito per evitare modifiche che non verrebbero aggiornate
                %>
                   <p><textarea class="input-block-level" rows="6"  name="txtDomandaplus" disabled="true" cols="96"></textarea> </p>
                <%else%>
                   <p><textarea class="input-block-level" rows="6"  name="txtDomandaplus" cols="96"></textarea> </p>
                <%end if%>
       
          <%end if%>
        <%end if' if strcomp(by_UPLOAD,"1")=0 then%> 
  <% end if%>
  
 			 <% if strcomp(by_UPLOAD,"1")=0 then ' se ho gia fatto upload ho anche inseriro domanda quindi disabilito per evitare modifiche che non verrebbero aggiornate
                %>
                  <p><input type="text" class="input-xxlarge" name="txtR1"  disabled="true" value="<%=R1%>"   size="135" maxlength="150"><b> 
                    Risposta 1</b></p>
                  <p>
                    <input type="text" class="input-xxlarge" name="txtR2" disabled="true" value="<%=R2%>"    size="135" maxlength="150"><b> 
                    Risposta 2 </b></p>
                  <p>
                    <input type="text" class="input-xxlarge" name="txtR3" disabled="true" value="<%=R3%>"  size="135" maxlength="150"><b> 
                    Risposta 3 </b></p>
                  <p><input type="text" class="input-xxlarge" name="txtR4" disabled="true" value="<%=R4%>"    size="135" maxlength="150"><b> 
                    Risposta 4 </b></p>
                  <p><input type="text" class="input-xxlarge" name="txtRE" disabled="true" value="<%=RE%>"  size="1"><b> 
                    Risposta Esatta </b></p>
                    <b> Spiegazione OK</b> <br><br>
                   <div style="border:solid #CCF; background: #CCC; width:800px; height:auto; padding:10px;">
                    <% response.write(sReadAll)%>
              
               </div>
             
                <%else%>
                   <p><b> Vero  <input TYPE="RADIO"  name="VF" value=1> <br>
                   <p><b> Falso  <input TYPE="RADIO"  name="VF" value=0> <br>
              
                    <p><b> 
                    Spiegazione </b></p>
                   <p><textarea class="input-block-level"  rows="6"  name="S1" cols="96"></textarea> </p>
                <%end if%>
				
   
   <b>Lingua</b>
	 <select name="lingua">
		  <option value="it">Italiano</option>
		  <option value="en">Inglese</option>
 
		</select> 
    
     <%if (session("Admin")=true) then %>
 <p>  
      <b>Inserisci in </b> 
	 <%'visualizzo le checkbox per la scelta del QUiz in cui inserire
	  			    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") & "';" 
					Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
					' se raggiungo il limite ricomncio
					 
					max_in_quiz=clng(rsTabella1("Max_In_Quiz"))%>
					Tutti <input type="RADIO" checked="true" name="inQuiz" value="-1">  <br>	
					<% for i=1 to max_in_quiz %>
						 
                      <%=i%>    <input TYPE="RADIO"  name="inQuiz" value="<%=i%>"> 			 
					<% next %>
		 
	<% else %>
    
	
     <b>Inserisci in </b> 
	
    <%   QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") & "';" 
					Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
					' se raggiungo il limite ricomncio
					 
					max_in_quiz=clng(rsTabella1("Max_In_Quiz"))
	
	   for i=1 to max_in_quiz %>
						 <% if i=In_Quiz then %>
                      <%=i%>    <input TYPE="RADIO" checked="true"  name="inQuiz" value="<%=i%>"> 	
                         <%else%>
                          <%=i%>    <input TYPE="RADIO"   name="inQuiz" value="<%=i%>"> 
                         <% end if%>		 
					<% next %>
    
    
	
	<% end if %>
    
    
    
     <% if strcomp(by_UPLOAD,"1")=0 then ' da verificare !!!!! inserisci_test %>
       <p><a href="../cClasse/scegli_azione_test.asp?Num=<%=Num+1%>&Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>">
       <input type="button" value="Termina" class="btn">
       </a>   
         </p>      
  <%else%>
  <p><input type="submit" value="Invia" name="B1" class="btn"> 
   </p> 
  <%end if%>
                      
              
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

