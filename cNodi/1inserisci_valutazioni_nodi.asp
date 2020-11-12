 <!-- mostra tutti i nodi per la modifica studente o valutazione admin-->
 <%@ Language=VBScript %>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<style>
<!--
 li.MsoNormal
	{mso-style-parent:"";
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	margin-left:0cm; margin-right:0cm; margin-top:0cm}
-->
</style>
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Valuta o Modifica Nodi Studente</title>
 
 <script type="text/javascript" src="../calendar/calendar.js"></script>
<script type="text/javascript" src="../calendar/calendar-it.js"></script>
<script type="text/javascript" src="../calendar/calendario.js"></script>
 <script type="text/javascript" src="../js/selezionatutti.js"></script>

<script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

location.href="studente_domande.asp"
//location.href=window.history.back();
 }
 </script>
 <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 
 
 //assegna la valutazione solo se il record è selezionato per la valutazione
function valutaTutti(voto) {
	 
	
	var stringa,stringa2;
	//window.alert("ciao2");
	//window.alert(document.dati.txtVoto.value);
	var voto=document.dati.txtVoto.value;
	//window.alert(voto);
	//window.alert("ciao3");
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
    
    with (document.dati) {
	for (var i=0; i < elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'txtDATA'+idtext)
		    {
		    elements[i].value = el.value; 
			idtext=idtext+1;
			}
	 }
	 return true;
    }
 }
 
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
</head>

<%
  Response.Buffer = true
 ' On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>

  
  
 

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO,i
  Dim ConnessioneDB,rsTabella, QuerySQL,CodiceTest,StringaConnessione
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<%  
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  'CodiceAllievo=Request.QueryString("cod")
  'cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  'CodiceDomanda=Request.QueryString("CodiceDomanda")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  'response.write("TP:"&TitoloParagrafo)
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update

  '-----
  Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("ID_MOD")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  if left(Cartella,1)<>"" then
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if
  
  '----
  
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
	          QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_Paragrafo='"& Paragrafo &"'"
	        end if
		  end if 
	  end if  

    end if 
	
end if 
               
	QueryPrima=QuerySQL	
	
	 
'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_NODI1 Where ID_Paragrafo='"& Paragrafo &"'"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)				
   
Set objFSO = CreateObject("Scripting.FileSystemObject")

' verifico se è privato 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaP = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabellaP.fields("Privato") 
	rsTabellaP.close
	
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine

%>
   
 

<div id="container">
 
	<div class="contenuti_login" style="width: 100%; height: auto;">
<!---------->
<%if (session("Admin")=true) then %>
	<form method="POST"  action="1inserisci_valutazioni_nodi.asp?Nulle=1&Tutte=<%=Tutte%>&Gruppi=<%=Gruppi%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
	  <b>Seleziona Nodi </b><br>
	  <br>
	<b>Da valutare </b>
	 <input type="submit" value="Voto=0" name="B1"> </p> 
	</form> 

<form method="POST"  action="1inserisci_valutazioni_nodi.asp?Gruppi=<%=Gruppi%>&Tutte=<%=Tutte%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
<b>Data da:</b>
<%if data<>"" then%>
<input type="text" name="txtDATA" value="<%=Data%>" size="10">
<% else%>
<input type="text" name="txtDATA" value="gg/mm/aaaa" size="10">
<% end if%>
 <input type="submit" value="Invia" name="B1"> </p> 
</form>
</div>
<%end if %>
  
  <form method="POST" name="dati" action="1inserisci_valutazioni_nodi1.asp?NumRec=<%=i%>&TitoloParagrafo=<%=TitoloParagrafo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
  <div id="bloc_destra_cont"> 
 <br><br>
  <font size="4" color="#FF0000"><b>Valuta o modifica </b></font><strong><font color="#FF0000" size="4">i Nodi </font></strong><b><font color=#FF0000 size="4"> :</font></b>
  <br>
  <p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Capitolo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (TitoloParagrafo) %></b></font> <!-- stampa il titolo del test -->
	
    <p>
	
	<div class="contenuti_login" style="width: 1093px; height: auto;" >	
	<%
	i=0
	TitoloParagrafo1=TitoloParagrafo
	'response.write(QuerySql) 
    do while not rsTabella.eof
	
	 if StrComp(TitoloParagrafo1, rsTabella("Titolo")) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")
                   
                   else 
                    'i=0 
                       TitoloParagrafo1= rsTabella("Titolo")
					 %>
					 <center><font color=#0066FF face ="Verdana" size="3"><hr>
                    <b>Paragrafo : <%Response.write (TitoloParagrafo1) %></b></font> 
                    </center><!-- stampa il titolo del test -->
	
                <%end if %>  
	
   <p><hr><br> 
 
			<tr><td><b><%=rsTabella(2)%>&nbsp;&nbsp;&nbsp;</b></td></tr>
              <input type="text" name="txtCodiceNodo<%=i%>"  tabindex="<%=(7*i)%>" value="<%=rsTabella.Fields("CodiceNodo")%>" size="10" maxlength="250">
              <b>Codice Nodo </b> 
			  <input type="text" name="txtDATA<%=i%>" value="<%=rsTabella.Fields("Data")%>" size="8" maxlength="250">
              <b>Data</b>
			  <input type="text" name="txtOraNodo<%=i%>" value="<%=rsTabella.Fields("Ora")%>" size="6" maxlength="250">
              <b>Ora</b> 
			  
			  <br>
            <p><input type="text" name="txtChi<%=i%>"  value="<%=rsTabella.Fields("Chi")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="250"><b>Chi <br>
	 

	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  <p><input type="text" name="txtR1Cosa<%=i%>" value="<%=rsTabella.Fields("Cosa")%>" tabindex="<%=(7*i)+1%>" size="135" maxlength="150"><b> 
	Cosa</b></p> 
  <p>
	<input type="text" name="txtR1Dove<%=i%>" value="<%=rsTabella.Fields("Dove")%>" tabindex="<%=(7*i)+2%>" size="135" maxlength="150"><b> 
	Dove </b></p>
  <p>
	<input type="text" name="txtR1Quando<%=i%>" value="<%=rsTabella.Fields("Quando")%>" tabindex="<%=(7*i)+3%>" size="135" maxlength="150"><b> 
	Quando </b></p>
  <p><input type="text" name="txtR1Come<%=i%>" value="<%=rsTabella.Fields("Come")%>" tabindex="<%=(7*i)+4%>" size="135" maxlength="150"><b> 
	Come </b></p>
  <p><input type="text" name="txtR1Perche<%=i%>" value="<%=rsTabella.Fields("Perche")%>" tabindex="<%=(7*i)+5%>" size="135"><b> 
	Perchè </b></p>
	<p><input type="text" name="txtR1Quindi<%=i%>" value="<%=rsTabella.Fields("Quindi")%>" tabindex="<%=(7*i)+6%>" size="135"><b> 
	Quindi </b></p>
	 <% 
	    Paragrafo=rsTabella(0)
		Modulo=rsTabella.fields("ID_Mod")
	    url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&rsTabella.Fields("CodiceNodo")&".txt"
    
    url=Replace(url,"\","/")
 


				
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	sReadAll = objTextFile.ReadAll
	'sReadAll=url
	'response.write(url)
	objTextFile.Close	%>
	<b>Spiegazione</b><p><textarea rows="7" name="S1<%=i%>" tabindex="<%=(7*i)+7%>" value="ciao" cols="116"><%=Response.write(sReadAll)%> </textarea></p>
 



<%if (session("Admin")=true) then %>
 
 
  <p><input type="text" name="txtVAl<%=i%>" value="<%=rsTabella.Fields("Voto")%>" size="1"  ><b> 
	Valutazione </b>   
    <input type="text" name="txtSegnalata<%=i%>" value="<%=rsTabella.Fields("Segnalata")%>" size="1"  ><b> 
	Segnalata </b> </p><p>
     <p><input type="checkbox"  name="cb<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b> 
	Seleziona per la stampa </b><br>
      <p><input type="checkbox"  name="cbVal<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna2('cbVal<%=i%>');">  <b> 
	Seleziona per la valutazione </b><br>
	<p> <input type="text" name="txtINQUIZ<%=i%>" value="<%=rsTabella.Fields("In_Quiz")%>" size="1" ><b> In Quiz </b></p>
  <!--Definisce i due bottoni del form -->
<% else 
   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
 <p><input type="text" disabled="disabled" name="txtVAl<%=i%>" value="<%=rsTabella.Fields("Voto")%>" size="1"><b> 
	Valutazione </b></p><p>
	 <input type="text" disabled="disabled" name="txtINQUIZ<%=i%>" value="<%=rsTabella.Fields("In_Quiz")%>" size="1"><b> In Quiz </b>
	</p>
  <!-- <p><input type="submit" value="Invia" name="B1"> </p> <!--Definisce i due bottoni del form --> 
<% end if 
end if 


    i=i+1
	'response.write(i)
    rsTabella.movenext
loop
%>
<hr><br>
 <input type="text" name="txtNUMREC" value="<%=i%>" size="1">
 
  <% if Session("Admin")=true then%>
<p>
 <div class="immagini">
   <b>Consegnati</b> <input type="text" name="date3" id="sel3" size="10" value="gg/mm/aaaa"> 
    <input type="reset" value=" ... " title="Clicca e seleziona data di consegna nel calendario a fondo pagina" onClick="return showCalendar('sel3', '%d/%m/%Y');">
    <input type="button" value="Tutti " title="Attribuisci a tutti la stessa data di consegna" onClick="selezionatutti('sel3');">
   
   <br><br><b>Voto</b><input type="text"   name="txtVoto" size="1">
<input type="button" value="Valuta tutti" onClick="valutaTutti()" >
<input type="text" name="txtNUMVAL" value="<%=i%>" size="1" style="border:none">
   
   </div>
   
<%end if%>
 
 
<p><input type="submit" value="Invia" name="B1"> </p> 
</form> <!-- Chiude l'interfaccia -->
<p><a target="_new" href="7_stampa_schede_nodi.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&QuerySQL=<%=QueryPrima%>"><img src="../../img/printer.jpg" alt="Stampa questa scheda"></a></p>
 <!--#include file="../include/tornaquaderno.html" -->
   
</div>
</body>
<% else%> 
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if %>
</html>