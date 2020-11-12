<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Valutazioni nodi</title>   
   
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
  On Error Resume Next  
   





 
  
 
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
 
  Dim objFSO, objTextFile 
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<% 
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceAllievo=Request.QueryString("cod")
  'cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceFrase=Request.QueryString("CodiceFrase")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("CodiceTest")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update
  
  '-----
  Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("ID_MOD")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  if left(Cartella,1)<>"" then ' DA SISTEMARE NELLE QUERY PER I GRUPPI !!!!!!!!!!!!!
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if
  
  ID_Prenodo=Request.QueryString("ID_Prenodo")
  NodoScelto=Request.QueryString("NodoScelto")
  
  BoxApro=Request.QueryString("BoxApro")
  
  
  'per selezionare il periodo della 
DataClaq=Session("DataClaq")
DataClaq2 =Session("DataClaq2")' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
 
  
  '----
 FraseScelta=Request.QueryString("FraseScelta")  
  
 
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
 if (NodoScelto<>"") then ' se sono chiamata da 2scegli_valutazioni_frasi visualizzo la stessa frase per tutti gli studenti
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where  ID_Prenodo="&ID_Prenodo&" and ID_Mod<>'6C' AND Cartella = '" & Cartella &"' order by Data asc, Ora asc;"  
					
else
		if (CodiceAllievo<>"") then  ' se sono stata chiamata dalla pagina studente_domande, valuterò solo le domande di quello studente
			 if (Nulle<>"") then ' se devo mostrare sollo quelle con voto=0
					  if (Tutte<>"") then
						  QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where Voto=0 and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"
						else
							 QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_MOD='"& ID_MOD &"' and Voto=0 and CodiceAllievo='"&CodiceAllievo&"';"
						end if 
			else	        
					if (Data<>"") then ' se devo mostrare sollo quelle dopo una certa data
						if (Tutte<>"") then
							 QuerySQL="SELECT MODULO_PARAGRAFO_NODI1.*, MODULO_PARAGRAFO_NODI1.Data FROM MODULO_PARAGRAFO_NODI1 WHERE MODULO_PARAGRAFO_NODI1.Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"#  and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"
					   else
							 QuerySQL="SELECT MODULO_PARAGRAFO_NODI1.*, MODULO_PARAGRAFO_NODI1.Data FROM MODULO_PARAGRAFO_NODI1 WHERE MODULO_PARAGRAFO_NODI1.Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"#  and ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"';"
					   end if
					else
					   if (Tutte<>"") then
							QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where  CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C';"
						else
							 QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"';"
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
						QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_Paragrafo='"& Paragrafo &"' and Voto=0"
				 else	        
					 if (Data<>"") then
				 
					   QuerySQL="SELECT MODULO_PARAGRAFO_NODI1.*, MODULO_PARAGRAFO_NODI1.Data FROM MODULO_PARAGRAFO_NODI1 WHERE MODULO_PARAGRAFO_NODI1.Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# AND ID_Paragrafo='"& Paragrafo &"';"
					else
					  QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_Paragrafo='"& Paragrafo &"';"
					end if
				  end if 
			  end if  
		
			end if 
			
		end if 
               
end if		

'response.write(QuerySQL)
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
        <body class='theme-<%=session("stile")%>'  oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="Effect.toggle('dAttività','appear');Effect.toggle('dAvvisi','appear'); return false;">  
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
						<h1> <i class="glyphicon-snowflake"></i> Valutazioni nodi</h1> 
                    
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
                      
 
 			 		 
	 <%if (session("Admin")=true) then %>
	<form  method="POST" class="form-vertical"  action="1inserisci_valutazioni_nodi.asp?Nulle=1&Tutte=<%=Tutte%>&Gruppi=<%=Gruppi%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>"><b>Seleziona nodi</b><br><br>
	<b>Da valutare </b>
	 <input type="submit" value="Voto=0" name="B1" class="btn"> </p> 
	</form> 

<form method="POST" class="form-vertical"action="1inserisci_valutazioni_nodi.asp?Gruppi=<%=Gruppi%>&Tutte=<%=Tutte%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
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
  <form name="dati" class="form-vertical" method="POST"   action="1inserisci_valutazioni_nodi1.asp?NumRec=<%=i%>&TitoloParagrafo=<%=TitoloParagrafo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&BoxApro=<%=BoxApro%>">
 
<!----><p align="center">
 
	
	
	<%
	TitoloParagrafo1=TitoloParagrafo
	i=1
	'response.write(QuerySql) 
	
	' apro il file di testo che conterrà gli url delle domande da modificare 
    do while not rsTabellaNew.eof
%>  <!-- <p><hr style="width:80%" align="center"><br>-->
<!--<div class="hr"><hr /></div><br>-->

			<% if StrComp(TitoloParagrafo1, rsTabellaNew("TitPar")) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")
                   
                   else 
                    'i=0 
                       TitoloParagrafo1= rsTabellaNew("TitPar")
					    %>
                        <font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (TitoloParagrafo1) %></b></font> <!-- stampa il titolo-->
					 
			      
                <%end if %>  	


 
<fieldset><legend><h4>
			 <b><%=UCASE(rsTabellaNew(2))%>&nbsp;<%=left(UCASE(rsTabellaNew("Nome")),1)&"."%></b> &nbsp; &nbsp; &nbsp;</h4></legend>
             
             
             <div class="control-group">
				 
				  <div class="controls">
	
    			
    
      <input type="text" class="input-mini" name="txtCodiceNodo<%=i%>"  tabindex="<%=(7*i)%>" value="<%=rsTabellaNew.Fields("CodiceNodo")%>" size="10" maxlength="250">
              <b>Codice Nodo </b> 
			  <input type="text" name="txtDATA<%=i%>" value="<%=rsTabellaNew.Fields("Data")%>" size="8" maxlength="250">
              <b>Data</b>
			  <input type="text" class="input-small" name="txtOraNodo<%=i%>" value="<%=rsTabellaNew.Fields("Ora")%>" size="6" maxlength="250">
              <b>Ora</b> 
			  
			  <br>
            <p><input type="text" class="input-xxlarge" name="txtChi<%=i%>"  value="<%=rsTabellaNew.Fields("Chi")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250"><b>Chi <br>
	 

	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  <p><input type="text" class="input-xxlarge" name="txtR1Cosa<%=i%>" value="<%=rsTabellaNew.Fields("Cosa")%>" tabindex="<%=(7*i)+1%>"   maxlength="150"><b> 
	Cosa</b></p> 
  <p>
	<input type="text" class="input-xxlarge" name="txtR1Dove<%=i%>" value="<%=rsTabellaNew.Fields("Dove")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150"><b> 
	Dove </b></p>
  <p>
	<input type="text" class="input-xxlarge" name="txtR1Quando<%=i%>" value="<%=rsTabellaNew.Fields("Quando")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"><b> 
	Quando </b></p>
  <p><input type="text" class="input-xxlarge" name="txtR1Come<%=i%>" value="<%=rsTabellaNew.Fields("Come")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150"><b> 
	Come </b></p>
  <p><input type="text" class="input-xxlarge" name="txtR1Perche<%=i%>" value="<%=rsTabellaNew.Fields("Perche")%>" tabindex="<%=(7*i)+5%>" size="135"><b> 
	Perch&egrave; </b></p>
	<p><input type="text" class="input-xxlarge" name="txtR1Quindi<%=i%>" value="<%=rsTabellaNew.Fields("Quindi")%>" tabindex="<%=(7*i)+6%>" size="135"><b> 
	Quindi </b></p>
	  
          <INPUT TYPE="HIDDEN" NAME="txtCodiceAllievo<%=i%>" VALUE="<%=rsTabellaNew("CodiceAllievo")%>">
	
    
    
    
    
    
	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  
	 <% 
	    Paragrafo=rsTabellaNew(0)
		
		Modulo=rsTabellaNew.fields("ID_Mod")
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabellaNew.Fields("TitoloParagrafo")&"_"&rsTabellaNew.Fields("CodiceNodo")&".txt"
   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url=Replace(url,"\","/")
  'Response.write(url)
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

    urliFrame="https://www.umanet.net"&homesito&"/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceNodo")&".txt"
	
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	sReadAll = objTextFile.ReadAll
	'sReadAll=url
	'response.write(sReadAll)
	objTextFile.Close	%>
	<b>Sintesi</b>
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
 
 <%     QuerySQL1="Select * from Nodi_Img where Id_Nodo="& rsTabellaNew("CodiceFrase")&";"
	   url= "../Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Nodi/Img" ' vuole il percorso relativo della cartella
       url=Replace(url,"\","/")
	   
	   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)        
   	   do while not rsTabella1.eof
	   'response.write(url&"/"& rsTabellaNew("Url")&"<br>")
	   
	   urlimg=url&"/"& rsTabella1("Url") ' aggiungo al percorso il nome del file
	   urldelete=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Nodi/Img/"&rsTabella1("Url")  ' per cancellare l'immagine.jpg
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
	 <input type="submit" value="Invia" name="B1" class="btn-primary"> </p> 
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
<b>Stampa <input class="input-mini" type="text" name="txtNUMREC" value="<%=i-1%>">Frasi</b></p> 
 
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
    document.dati.action = "7_stampa_schede_frasi.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&QuerySQL=<%=QueryPrima%>";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
}
 </script>

 </html>

