<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Valutazioni domande</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
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
	<!-- jQuery UI -->
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.draggable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
   
    <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
     
	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       
    
       
<!--Controllo accesso quaderno e sessione scaduta con redirect ad index.html-->
       <script src="../js/privacy.js"></script>
       
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
    
  <script type="text/javascript">
function checkTutti() {
	numcb=0;
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		if (elements[i].type == 'checkbox')
		    {
		     elements[i].checked = true;
			 numcb=numcb+1;
			}
		}
	}
	document.dati.txtNUMREC.value=numcb;
}
function uncheckTutti() {
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'checkbox')
		elements[i].checked = false;
		}
	 
	}
	document.dati.txtNUMREC.value=0;
	
}
function aggiorna(nome) {
	 
		with (document.dati) { 
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMREC.value=parseInt(txtNUMREC.value)+1;
		 else
		    txtNUMREC.value=parseInt(txtNUMREC.value)-1;
	    }	
}
function aggiorna2(nome) {
	 
		with (document.dati) { 
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)+1;
			
		 else
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)-1;
	    }	
}

 </script>
 
 <script type="text/javascript">
//assegna la valutazione solo se il record è selezionato per la valutazione
function valutaTutti(voto) {
	var stringa,stringa2;
	var voto=document.dati.txtVoto.value;
	
	numcb=1;
	 
		for (var i=0; i < document.dati.elements.length; i++) {
			stringa=document.dati.elements[i].name;
			stringa2='txtVAl'+numcb;
			
		if (stringa.search(stringa2) == 0)
		     {
			if (document.dati.elements["cbVal"+numcb].checked == true) document.dati.elements[i].value = voto;
			numcb=numcb+1;
			 
		 	}
		}
}
 
  function selezionatutti(id) {
	//per modificare tutte le date di un form impostandole uguale al valore della textbox passata per parametro
    //document.dati.date3.value="11/11/1111";
	// document.dati.txtScadenza1.value="19/11/2010";
	
    var el = document.getElementById(id);
    var idtext=1;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=0; i < elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'txtDataDomanda'+idtext)
		    {
		    elements[i].value = el.value; 
			idtext=idtext+1;
			}
	 }
	 return true;
    }
 }
 
 </script> 

     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<%
  Response.Buffer = true
' Abilita la gestione degli errori
On Error Resume Next 
 
 if session("CodiceAllievo")="" then
    stringa="vuoto"
	
 else
   stringa="pieno"
 end if



 
  
 
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
 
  Dim objFSO, objTextFile 
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<% 
  Codice_Test=Request.QueryString("CodiceTest")
  'CodiceDomanda=Request.QueryString("CodiceDomanda")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
 ' TitoloParagrafo = Replace(TitoloParagrafo, Chr(44), "")
  DATA=Request.QueryString("DATA")
  Modulo=Request.QueryString("Modulo")
  ID_MOD=Request.QueryString("ID_MOD")
   ID_PAR=Request.QueryString("ID_PAR")
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update
  Gruppi=Request.QueryString("Gruppi")
  Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  if left(Cartella,1)<>"" then
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if
  
  BoxApro=Request.QueryString("BoxApro")
 
 function ReplaceCar(sInput)
dim sAns
   
  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
' 
 ' sAns=  Replace(sInput,"&igrave;","i")
'  sAns=  Replace(sAns,"&egrave;","e'")
'  sAns=  Replace(sAns,"&ugrave;","u'")
'  sAns=  Replace(sAns,"?","&ograve;")
'  sAns=  Replace(sAns,"&agrave;","a'")
' 
 
ReplaceCar = sAns
'ReplaceCar = sInput

end function
 
 
'if MO<>"" then 
' Modulo=MO
'end if  
'
Segnalate=request.QueryString("Segnalate")
if (Segnalate<>"") then  ' se sono stata chiamata dalla pagina studente_domande, valuterò solo le domande di quello studente
     if (Nulle<>"") then ' se devo mostrare sollo quelle con voto=0
	          if (Tutte<>"") then
			      QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where Voto=0 and Segnalata=1 and ID_Mod<>'6C';"
		        else
				     QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_MOD='"& ID_MOD &"' and Voto=0 and Segnalata=1;"
				end if 
	else	        
            if (Data<>"") then ' se devo mostrare sollo quelle dopo una certa data
	            if (Tutte<>"") then
		             QuerySQL="SELECT MODULO_PARAGRAFO_DOMANDE1.*, MODULO_PARAGRAFO_DOMANDE1.Data FROM MODULO_PARAGRAFO_DOMANDE1 WHERE MODULO_PARAGRAFO_DOMANDE1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"  and Segnalata=1 and ID_Mod<>'6C';"
			   else
			         QuerySQL="SELECT MODULO_PARAGRAFO_DOMANDE1.*, MODULO_PARAGRAFO_DOMANDE1.Data FROM MODULO_PARAGRAFO_DOMANDE1 WHERE MODULO_PARAGRAFO_DOMANDE1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"  and ID_MOD='"& ID_MOD  &"' and Segnalata=1;"
			   end if
	        else
			   if (Tutte<>"") then
	                QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where  and Segnalata=1 and ID_Mod<>'6C';"
			 	else
				     QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"& ID_PAR  &"' and Segnalata=1;"
				end if	
	        end if
		  end if 
else
    if (Gruppi<>"") and (Nulle<>"") then ' mostro le domande per gruppo solo quelle con voto =0 
  'response.write("QUI")
	QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &" and Voto=0;"
    else
	   if (Gruppi<>"") then
	       QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &";"
	   else
          if (Nulle<>"") then
		        QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"& Paragrafo &"' and Voto=0"
	  	 else	        
             if (Data<>"") then
	     
		       QuerySQL="SELECT MODULO_PARAGRAFO_DOMANDE1.*, MODULO_PARAGRAFO_DOMANDE1.Data FROM MODULO_PARAGRAFO_DOMANDE1 WHERE MODULO_PARAGRAFO_DOMANDE1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &" AND ID_Paragrafo='"& Paragrafo &"';"
	        else
	          QuerySQL="SELECT * FROM MODULO_PARAGRAFO_DOMANDE1 Where ID_Paragrafo='"& Paragrafo &"'"
	        end if
		  end if 
	  end if  

    end if 
	
end if 

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\Anno_2012-2013_2\logFile1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				
'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_Paragrafo='"& Paragrafo &"'"
Set rsTabellaNew = ConnessioneDB.Execute(QuerySQL)	

'QueryPrima=	QuerySQL
QueryPrima=	QuerySQL

'per il copia incolla ed il privato
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	Privato=rsTabellaCI.fields("Privato") 
	rsTabellaCI.close
 
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
  


Set objFSO = CreateObject("Scripting.FileSystemObject")
%>
   

<%  if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
		   <% if (CIAbilitato=0) then ' disabilito copia incolla%>
        <body class='theme-<%=session("stile")%>'  oncontextmenu="return false" ondragstart="return false" onselectstart="return false">  
        <%else%>
         <body class='theme-<%=session("stile")%>'>
         
        <%end if%>
  <% end if %>





	<div id="navigation">
     
   
	
		 
        <!-- #include file = "../var_globali.inc" --> 
 		 
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
 <%
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
 %>   
    
    
	<div class="container-fluid" id="content">
   
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-question-sign"></i>Valutazioni domande </h1> 
                    <%'response.write(QueryPrima)%>
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
							<a href="../cClasse/home_app.asp?id_classe=<%=session("id_classe")%>">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Valutazioni</a>
                            
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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%> : <%=TitoloParagrafo%></h3>
			          </div>
				      <div class="box-content">
                      
<% if rsTabellaNew.eof then%>
    <span class="alert-error">Non ci sono compiti da valutare
    </span><br><br><b>
    <a href="javascript:history.back()">	Indietro </a></b>
    <%
	else%>
 			 		 
	 <%if (session("Admin")=true) then %>
	<form  method="POST" class="form-vertical"  action="../inserisci_valutazioni.asp?BoxApro=<%=BoxApro%>&Nulle=1&Tutte=<%=Tutte%>&Gruppi=<%=Gruppi%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>"><b>Seleziona domande</b><br><br>
	<b>Da valutare </b>
	 <input type="submit" value="Voto=0" name="B1" class="btn"> </p> 
	</form> 

<form method="POST" class="form-vertical"action="../inserisci_valutazioni.asp?BoxApro=<%=BoxApro%>&Gruppi=<%=Gruppi%>&Tutte=<%=Tutte%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
<i class="icon-calendar"></i>  <b>Data da:</b>
<%if data<>"" then%>
<input type="text" name="txtDATA" value="<%=Data%>" class="input-small datepick">
<% else%>
<input type="text" name="txtDATA" value="gg/mm/aaaa" class="input-small datepick">
<% end if%>
  
 <input type="submit" value="Invia" name="B1" class="btn"> </p> 
 
</form>
</div>
<%end if %>
  




 <div class="immagini" style="height:auto; width:auto; border:none;" > 
  <form name="dati" class="form-vertical" method="POST"  action="inserisci_valutazioni1.asp?BoxApro=<%=BoxApro%>&NumRec=<%=i%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>" > 
 
<!----><p align="center">
 
	
	
	<%
	TitoloParagrafo1=TitoloParagrafo
	i=1
	'response.write(QuerySql) 
	
	' apro il file di testo che conterrà gli url delle domande da modificare 
	
	
    do while not rsTabellaNew.eof
%>  <!-- <p><hr style="width:80%" align="center"><br>-->
<!--<div class="hr"><hr /></div><br>-->

			<% if StrComp(TitoloParagrafo1, rsTabellaNew("Tit")) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")
                   
                   else 
                    'i=0 
                       TitoloParagrafo1= rsTabellaNew("Tit")
					    %><span class="alert-info">
                      <%Response.write (TitoloParagrafo1) %>   <!-- stampa il titolo-->
                         </span>
					 
			      
                <%end if %>  	


 
<fieldset><legend><h4>
			 <b><%=ucase(rsTabellaNew(2))%></b> &nbsp; &nbsp; &nbsp;</h4>
            <small><%=rsTabellaNew("Cognome")%> &nbsp;<%=left(rsTabellaNew("Nome"),1)&"."%></small> </legend>
             
             
             <div class="control-group">
				 
				  <div class="controls">
	
    			
             
             
             
            
               <input class="input-xxlarge" type="hidden" name="txtDomanda<%=i%>"  value="<%=rsTabellaNew.Fields("Quesito")%>" tabindex="<%=(7*i)+1%>"   maxlength="250">
          <INPUT TYPE="HIDDEN" NAME="txtCodiceAllievo<%=i%>" VALUE="<%=rsTabellaNew("CodiceAllievo")%>">
          
          <%if rsTabellaNew.Fields("Tipo")=1 then
	ID=rsTabellaNew.Fields("CodiceDomanda") ' per la funzione domandaplus
	 %>
	   <br>
	   <textarea rows="3" name="TestoDomandaPlus<%=i%>" value="ciao" cols="96"><%=Response.write(domandaplus())%> </textarea><br>
		
	<% end if%>
    
  <%if not(rsTabellaNew.Fields("VF")=1) then ' non è una domanda vero falso %>
	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  <p><input type="text" class="input-xxlarge" name="txtR1<%=i%>" value="<%=rsTabellaNew.Fields("Risposta1")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150"><b> 
	Risposta 1</b></p> 
  <p>
	<input type="text" class="input-xxlarge" name="txtR2<%=i%>" value="<%=rsTabellaNew.Fields("Risposta2")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150"><b> 
	Risposta 2 </b></p>
  <p>
	<input type="text" class="input-xxlarge" name="txtR3<%=i%>" value="<%=rsTabellaNew.Fields("Risposta3")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"><b> 
	Risposta 3 </b></p>
  <p><input type="text" class="input-xxlarge" name="txtR4<%=i%>" value="<%=rsTabellaNew.Fields("Risposta4")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150"><b> 
	Risposta 4 </b></p>
 
  <p><input type="text" class="input-mini" name="txtRE<%=i%>" value="<%=rsTabellaNew.Fields("RispostaEsatta")%>" tabindex="<%=(7*i)+5%>" size="2"><b> 
	Risposta Esatta </b></p>
    
     <%else ' è vero falso%>
          
          
            
            <br><br>
             <% if (rsTabellaNew.Fields("RispostaEsatta")=1)  then  %>
                                            
											 <INPUT TYPE="RADIO" name="txtRE<%=i%>" checked="true" value="1">Vero 
                                             <INPUT TYPE="RADIO" name="txtRE<%=i%>" value="0">Falso 	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtRE<%=i%>" value="1">Vero   
                                             <INPUT TYPE="RADIO" name="txtRE<%=i%>"   checked="true" value="0">Falso  
                                           
										<% end if %>
    
             
	 <%end if%>
	
	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  
	 <% 
	    Paragrafo=rsTabellaNew(0)
		
		Modulo=rsTabellaNew.fields("ID_Mod")
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceDomanda")&".txt"
   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url=Replace(url,"\","/")
 ' Response.write(url)
'response.write(Server.MapPath(homesito))

	          ' url1="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFile.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.WriteLine(Modulo)
'				objCreatedFile.WriteLine(Paragrafo)
'			 
'				objCreatedFile.Close
'
'response.write(url)

    urliFrame="https://www.umanet.net"&homesito&"/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceDomanda")&".txt"
	
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	sReadAll = objTextFile.ReadAll
	'sReadAll=url
	'response.write(sReadAll)
	objTextFile.Close	%>
	<b>Spiegazione</b>
	<input  type="text" name="url<%=i%>" value="<%=url%>" size="0" class="hidden">
	<p>
	
   <% lunghezza=1+round((len(sReadAll))/40)%>
    <%' if CIAbilitato=0 then   ' se lo impedisco metto la textarea altrimenti iframe %>
  
	<textarea class="input-block-level" rows="<%=1+round((len(sReadAll))/60)%>"   tabindex="<%=(2*i)+1%>" name="txtSpiegazione<%=i%>"><%
			 ' if clng(rsTabellaNew(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			' finito=false
			 'sReadAll1=sReadAll
'			 inizio=1
'			' do 
'			  
'			  posApice=instr(inizio,sReadAll1,"'")
'			  if posApice=0 then
'			      finito=true
'			   else
'			     if  (strcomp (Mid(sReadAll1,posApice-1,1),"e")<>0) then ' se il carattere prcedente  ' non è una e allora sostituisco con "
'			         sReadAll1=replace(sReadAll1,"'",chr(34)&";",posApice)
'				 end if 
'				 inizio=posApice+1
'			   end if	  
'			 ' response.write("Apice "& posApice)
'			    
'			 
'			' loop until finito=true 
			 sReadAll1=sReadAll
		 	' sReadAll1=replace(sReadAll1,"';",chr(34)&";") 
			' sReadAll1=replace(sReadAll1,"<'","<"&chr(34)) 
			 ' sReadAll1= replace(sReadAll,"e"&chr(34),"e'")
'			 sReadAll1= replace(sReadAll,"o"&chr(34),"o'")
'			 sReadAll1= replace(sReadAll,"a"&chr(34),"a'")
'			 sReadAll1= replace(sReadAll,"u"&chr(34),"u'")
			' sReadAll1= replace(sReadAll,"'",chr(34))
			 Response.write(ReplaceCar(sReadAll1))%> </textarea></p>
         <%'else%>    
             
            <center>
             <%'response.write(lunghezza)
			 
			 %>
           <!--   <iframe style="" src="<%=urliFrame%>"  width= 800px  height=<%=lunghezza*20%>px ></iframe>
              -->
          <%'end if%>  
 </center>
 <%'inserisco le eventuali immagini
if rsTabellaNew("Img")=1 then%>
 
 <%     QuerySQL1="Select * from Domande_Img where Id_Domanda="& rsTabellaNew("CodiceDomanda")&";"
	   url= "../Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Domande/Img" ' vuole il percorso relativo della cartella
       url=Replace(url,"\","/")
	   
	   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)        
   	   do while not rsTabella1.eof
	   'response.write(url&"/"& rsTabellaNew("Url")&"<br>")
	   
	   urlimg=url&"/"& rsTabella1("Url") ' aggiungo al percorso il nome del file
	   urldelete=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Domande/Img/"&rsTabella1("Url")  ' per cancellare l'immagine.jpg
	   urldelete=Replace(urldelete,"\","/")
	  
	   'response.write("urlimg="&urlimg)%>
       
      <% nome=right(rsTabella1("Nome"),len(rsTabella1("Nome"))- instr(ucase(rsTabella1("Nome")),"C:\FAKEPATH\"))
		 nome=left(nome,len(nome)-4)%>
    
       <p align="center">
       <img src="<%=urlimg%>" border="1"> <br>
      <%' response.write(nome)%><br>
      <a href="../service/cancella_immagine.asp?urldb=<%=rsTabella1("Url")%>&urlimg=<%=urldelete%>&CodiceAllievo=<%=Session("CodiceAllievo")%>"><img src="../../img/elimina_small.jpg" width="10" height="10" title="Elimina" onClick="return window.confirm('Vuoi veramente cancellare questa immagine?');"></a></p>
    <%rsTabella1.movenext
	   loop
%> <%
end if
%>

  <b>Codice Domanda </b> <input class="input-mini"  type="text" name="txtCodiceDomanda<%=i%>"  tabindex="<%=(2*i)%>" value="<%=rsTabellaNew.Fields("CodiceDomanda")%>" >
             <b>Data </b> <input type="text" name="txtDataDomanda<%=i%>" class="input-small"  value="<%=rsTabellaNew.Fields("Data")%>" size="8" maxlength="250"> 
             <b>Ora </b> 
            <b> <input type="text" class="input-mini" name="txtOraDomanda<%=i%>"  value="<%=rsTabellaNew.Fields("Ora")%>" size="5" maxlength="250"> <br><br>
  
<%if (session("Admin")=true) then %>
 <p><input class="input-mini" type="text" name="txtVAl<%=i%>" value="<%=rsTabellaNew.Fields("Voto")%>" size="1"  ><b> 
	Valutazione </b>   
   <br>
    
     
       <span title="Feedback all'autore"><b>Segnalata</b></span> 
											 
                                             <% if (rsTabellaNew.Fields("Segnalata")=1)  then  %>
                                            
											 <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"   checked="true" value="0">No  
                                           
										<% end if %><br>
      <p><input type="checkbox"  name="cb<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b> 
	Seleziona per la stampa </b><br>
      <p><input type="checkbox"  name="cbVal<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna2('cbVal<%=i%>');">  <b> 
	Seleziona per la valutazione </b><br>
	<!-- <input type="text" name="txtINQUIZ<%=i%>" value="<%=rsTabellaNew.Fields("In_Quiz")%>" size="1" ><b> In Quiz </b></p>
  <!--Definisce i due bottoni del form -->
  
  
  <% else 
   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
  <p><input type="text" disabled="disabled" name="txtVAl<%=i%>" value="<%=rsTabellaNew.Fields("Voto")%>" size="1"><b> 
	Valutazione </b><br>
   <p><input type="checkbox"  name="cb<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b> 
	Seleziona per la stampa </b><br>
    
	<!-- <input type="text" disabled="disabled" name="txtINQUIZ<%=i%>" value="<%=rsTabellaNew.Fields("In_Quiz")%>" size="1"><b> In Quiz </b> -->
	</p>
    <!--<p><input type="submit" value="Invia" name="B1"> </p>  Definisce i due bottoni del form -->
	 
<% end if 
end if 


    i=i+1
	'response.write(i)
	%> <br><%
    rsTabellaNew.movenext
loop
%>
	</div>
				</div>
             
</legend>
 
 
   
 

 <hr>


<img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa();">
&nbsp; 
<b>Stampa <input class="input-mini" type="text" name="txtNUMREC" value="<%=i-1%>">Domande</b></p> 
 
<input class="btn" type="button" value="Seleziona tutti" onClick="checkTutti()">
<input  class="btn" type="button" value="Deseleziona tutti" onClick="uncheckTutti()"><br><hr>
 <% if Session("Admin")=true then%>
<b>Voto</b><input class="input-mini" type="text"   name="txtVoto">
<input type="button" class="btn" value="Valuta tutti" onClick="valutaTutti()">
<input type="text" name="txtNUMVAL" value="<%=i-1%>" size="1" class="input-mini"><br>
 
 <hr>
  <i class="icon-calendar"></i> <b>Consegnati</b>  
    <input type="text"  name="date3" id="datepicker" class="input-small datepick" /></p>
  <input class="btn" type="button" value="Tutti" title="Attribuisci a tutti la stessa data di consegna" onClick="selezionatutti('datepicker')">
    
  <hr>
<input type="submit" value="Invia" name="B1" class="btn-primary"> </p> 
<%end if%>
 
</form> 
	 
				 
<%end if ' if not rsTabellaNew.eof %>			 
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        

			 
	</body>
    <% else%> 
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if %>
   
 <script language="javascript" type="text/javascript"> 
function stampa() {
    document.dati.action = "7_stampa_schede_domande.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
}
 </script>

 </html>

