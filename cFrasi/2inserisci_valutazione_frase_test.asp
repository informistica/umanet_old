<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Valutazioni frasi</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<meta charset="utf-8">

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
	<!--
      <script type="text/javascript" src="../js/utility.js"></script>

	  -->
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

    <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
 <script src="../../js/datapicker_it.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />



<!--Controllo accesso quaderno e sessione scaduta con redirect ad index.html-->
       <script src="../js/privacy.js"></script>

	   <!-- <script src="https://cdn.ckeditor.com/ckeditor5/16.0.0/inline/ckeditor.js"></script>-->

		<script src="ckeditor/ckeditor.js"></script>

<script language="javascript" type="text/javascript">
function showText3() {window.alert("Il compito è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"

 }
    </script>

	<% x = Request.ServerVariables("HTTP_REFERER")
if x = "" then %>
<script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

 //alert("<%=x%>");
 location.href="../../../../index.html";

//location.href=window.history.back();
 }
 </script>

 <% else %>
 <script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

 location.href="<%=x%>";

//location.href=window.history.back();
 }
 </script><% end if%>




<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<%
  Response.Buffer = true
  'On Error Resume Next









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
  Cartella=Request.QueryString("classe")
  umanet=Request.QueryString("umanet")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update
  ID_MOD=Request.QueryString("MO")
  i=1


   tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 tDom=request.querystring("tDom")
 tFra=request.querystring("tFra")
 tNod=request.querystring("tNod")
  'per selezionare il periodo della



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
 sAns = Replace(sAns,"\","\\")
 sAns = Replace(sAns,"""", "'")

' if DateDiff("d", cDate(rsTabellaNew.Fields("Data")), cDate("28/12/2017")) > 0 then
 sAns = Replace(sAns,VBCrlf,"&#13;&#10;")
 'sAns = DateDiff("d", cDate(rsTabellaNew.Fields("Data")), cDate("01/01/2018"))
 'end if

ReplaceCar = sAns
'ReplaceCar = sInput

end function

Function quanteVolte(str1)
  Dim strArray
  strArray = Split(str1, VBCrlf)
  quanteVolte = UBound(strArray)
End Function


'if MO<>"" then
' Modulo=MO
'end if
'



'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\Anno_2012-2013_2\logFile1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close


'per il copia incolla ed il privato
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

 'response.write QuerySQL
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)

	CIAbilitato=rsTabellaCI("CIAbilitato")
	Privato=rsTabellaCI("Privato")

	rsTabellaCI.close

	QuerySQL = "Select CIAbilitato from Allievi where CodiceAllievo = '"&Session("CodiceAllievo")&"'"

	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato2=rsTabellaCI("CIAbilitato")
	rsTabellaCI.close

	'response.write QuerySQL
	if CIAbilitato2 = 1 then
	CIAbilitato = 1
	end if

if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
  ' response.write ("CodiceFrase="& clng(CodiceFrase))   ' clng da errore visto che abbimo superato maxint=32767 frasi nel db
QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where CodiceFrase="& clng(CodiceFrase) &" and CodiceAllievo='"&CodiceAllievo &"'"
'response.write QuerySQL
Set rsTabellaNew = ConnessioneDB.Execute(QuerySQL)

if  rsTabellaNew.eof then

   session("fraseinesistente") = true
   'response.Redirect request.serverVariables("HTTP_REFERER")
   Response.Redirect "../cMessaggi/centro_messaggi.asp"

   %>  <BODY onLoad="showText2();">

   </BODY> <%
END IF


Set objFSO = CreateObject("Scripting.FileSystemObject")
%>


<%  if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
		  	<%if (CIAbilitato=0)  and Session("Admin")=False then  %>
        <!--<body  oncontextmenu="return false" ondragstart="return false" onselectstart="return false" > -->
        <!--<body>-->
        <body   class='theme-<%=session("stile")%>' >

                <%else%>
        <body   class='theme-<%=session("stile")%>' >
        <%end if%>
  <% end if %>


	<script src="parametri_iniz.js"></script>


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
						<h1> <i class="icon-comments"></i>Valuta e modifica frasi</h1>

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
							<a href="../home_app.asp?id_classe=<%=session("id_classe")%>">Libro</a>
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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%>: <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">









 <div class="immagini" style="height:auto; width:auto; border:none;" >
  <form id="frmDocument" name="dati" class='form-vertical form-bordered form-striped' method="POST"  action="2inserisci_valutazione_frase1.asp?umanet=<%=umanet%>&CodiceAllievo=<%=CodiceAllievo%>&CodiceTest=<%=Codice_Test%>&Cartella=<%=Cartella%>&Modulo=<%=ID_MOD%>&Paragrafo=<%=Paragrafo%>&tCap=<%=tCap%>&tSot=<%=tSot%><%=p%>&tFra=<%=tFra%>&CodiceFrase=<%=rsTabellaNew.Fields("CodiceFrase")%>&Chi=<%=rsTabellaNew.Fields("Chi")%>&Capitolo=<%=Capitolo%>" >

<!----><p align="center">



	<%
	TitoloParagrafo1=TitoloParagrafo



%>

     <% if session("Modificata")=true then
    session("Modificata")=false
  %>
    <span class="alert-success">Modifica effettuata<br>
     <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6>
     </span>
<% end if%>


<fieldset><legend><h4>
			 <b><%=UCASE(rsTabellaNew(2))%>&nbsp;<%=left(UCASE(rsTabellaNew("Nome")),1)&"."%></b> &nbsp; &nbsp; &nbsp;</h4></legend>


             <div class="control-group">

				  <div class="controls">





              <b>Codice Frase </b> <input class="input-mini" disabled type="text" name="txtCodiceDomanda"  value="<%=rsTabellaNew.Fields("CodiceFrase")%>" >&nbsp; &nbsp;
             <b>Data </b> <input type="text" disabled name="txtDataDomanda" id="datepicker" class="input-small datepick"  value="<%=rsTabellaNew.Fields("Data")%>" size="8" maxlength="250"> &nbsp; &nbsp;
             <b>Ora </b>
            <b> <input type="text" disabled class="input-mini" name="txtOraDomanda"  value="<%=left(rsTabellaNew.Fields("Ora"), 5)%>" size="5" maxlength="250"> <br><br>
             <b>Chi </b>   <input class="input-xxlarge" disabled type="text" name="txtDomanda"  value="<%=rsTabellaNew.Fields("Chi")%>" tabindex="<%=(7*i)+1%>"   maxlength="250">
          <INPUT TYPE="HIDDEN" NAME="txtCodiceAllievo" VALUE="<%=rsTabellaNew("CodiceAllievo")%>">
		  <INPUT TYPE="HIDDEN" id ="isImg" NAME="Img" VALUE="<%=rsTabellaNew("Img")%>">

	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->

	 <%
if rsTabellaNew("ID_Sottoparagrafo")<>"" then
 urlRisorsa=rsTabellaNew("URL")
else
   urlRisorsa=rsTabellaNew("URL_O")
end if

 

	    Paragrafo=rsTabellaNew(0)

		Modulo=rsTabellaNew.fields("ID_Mod")
		'Cartella=rsTabellaNew.fields("Cartella")
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceFrase")&".txt"
	    url=Replace(url,"\","/")
	  
	   'urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/Risorse/Mod_" &Modulo&"/"&urlRisorsa 
	   'urlRis=Replace(urlRis,"\","/")
		'response.write urlRis

   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    
	
    url_feedback=left(url,instr(url,".")-1)
    url_feedback=url_feedback&"_feedback.txt"

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


    urliFrame="https://www.umanetexpo.net"&homesito&"/Db"&Session("DB")&"/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceFrase")&".txt"

	'response.write(url)
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)

	sReadAll1=""

	' Use different methods to read contents of file.
	sReadAll1 = objTextFile.ReadAll

	if sReadAll1 = "" then
		sReadAll1 = "File spiegazione mancante. Elimina e reinserisci la frase nel tuo quaderno."
		dis = true
	else
		sReadAll1=sReadAll1
		dis = false
	end if

	if instr(sReadAll1,"<script>")<>0 then
	   sReadAll1=Replace(sReadAll1,"<script>","")
	   sReadAll1=Replace(sReadAll1,"</script>","")
	end if
	sReadAll=sReadAll1

	sReadAll1 = ltrim(sReadAll1)
	sReadAll1 = rtrim(sReadAll1)
	sReadAll = ltrim(sReadAll)
	sReadAll = rtrim(sReadAll)
	'sReadAll1=url
	'response.write(sReadAll)
	objTextFile.Close


  if (rsTabellaNew.Fields("Segnalata")=1) or (rsTabellaNew.Fields("Segnalata")=2)  then
     f="<span style='color:red'>Segnalata</span>"
     'response.write("<br>Segnalata:"&rsTabellaNew.Fields("CodiceFrase"))
    ' leggo feedback
    if objFSO.FileExists(url_feedback) then
    'Response.write("<br>"&url_feedback)
      Set objTextFile = objFSO.OpenTextFile(url_feedback, ForReading)
      feedback="" 'pulisco feedback -> altrimenti rimane la vecchia feedback
      feedback = objTextFile.ReadAll

      'feedback=url_feedback
      objTextFile.Close
    end if
  else
   f="<span>Segnalata</span>"
   url_feedback=""
   feedback=""

  end if


  	%>
	<b>Frase</b>
	<input  type="text" name="url" value="<%=url%>" size="0" class="hidden">
	<p>

   <% lunghezza=1+round((len(sReadAll))/40)%>
    <%' if CIAbilitato=0 then   ' se lo impedisco metto la textarea altrimenti iframe
	righe=1+round((len(sReadAll))/100)
	if righe <3 then
	  righe=3
	 end if

	 righe = quanteVolte(sReadAll)+2
	 
  sReadAll1=sReadAll%>

<div id="txtS1">
    <%=sReadAll%>
</div>
<textarea name="txtSpiegazione"  id="txtSpiegazione" style="display:none;">

</textarea>
<script>
   
</script>


  
 




 </center>
 <%'inserisco le eventuali immagini
 img=0
if rsTabellaNew("Img")=1 then
img=1
%>


		<br><h4>Modifica link immagini o documenti</h4>
							Inserisci documenti tramite link di condivisione di <a target="_blank" href="https://drive.google.com/">google drive</a> <br><br>

							Inserisci immagini tramite <a target="_blank" href="https://postimages.org/it/">questo servizio di hosting esterno</a>.&nbsp;<font color="#000000"><br><b>N.B</b>&nbsp;</font>Ridimensiona l'immagine (es.640x480) ed incolla (ctrl+v) il <b>Collegamento Diretto</b> (secondo della lista), inoltre <b>non inserire caratteri speciali nel nome del file</b>.<br><br>

							<div class="controls">
								<input name="txtImg1" id="textfield1" placeholder="Incolla collegamento diretto" class="input-xxlarge" style="width:100%" type="text">
							</div>
							 <div class="controls">
								<input name="txtImg2" id="textfield2" placeholder="Incolla collegamento diretto" class="input-xxlarge" style="width:100%" type="text">
							</div>
							 <div class="controls">
								<input name="txtImg3" id="textfield3" placeholder="Incolla collegamento diretto" class="input-xxlarge" style="width:100%" type="text">
							</div>

			<br>
			<br>

 <%     QuerySQL1="Select * from Frasi_Img where Id_Frase="& rsTabellaNew("CodiceFrase")&";"
	   url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Frasi/Img" ' vuole il percorso relativo della cartella
       url=Replace(url,"\","/")
	  ' response.write(url&"<br>")
	   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)

	   id_txtImg = 1

   	   do while not rsTabella1.eof
	   pagina=0 '1 quando devo inserire url a pagina .html o .php anzichè immagine
	   'response.write(url&"/"& rsTabellaNew("Url")&"<br>")

	   urlimg=url&"/"& rsTabella1("Url") ' aggiungo al percorso il nome del file
	   urldelete=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Frasi/Img/"&rsTabella1("Url")  ' per cancellare l'immagine.jpg
	   urldelete=Replace(urldelete,"\","/")

		response.write("<script>document.getElementById('textfield"&id_txtImg&"').value = '"&rsTabella1("Url")&"'</script>")

	  ' response.write("urlimg0="& instr("ciaohttp","https"))


	   %>

      <% nome=right(rsTabella1("Nome"),len(rsTabella1("Nome"))- instr(ucase(rsTabella1("Nome")),"C:\FAKEPATH\"))
		 nome=left(nome,len(nome)-4)%>

       <p align="center">
     <%  gdoc="false"
	     if ((instr(rsTabella1("Url"),"docs.google.com")<>0) or (instr(rsTabella1("Url"),"drive.google.com")<>0) or (instr(rsTabella1("Url"),"colab.research.google.com")<>0) )  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
         gdoc="true"
		 response.write("<a href='"& rsTabella1("Url") &"' target='_blank'>apri url google drive</a>")
		  %>
	 <% else%>
	   <% if ((instr(rsTabella1("Url"),"tp://")<>0) or (instr(rsTabella1("Url"),"tps://")<>0)) and ((instr(rsTabella1("Url"),".jpg")<>0) or (instr(rsTabella1("Url"),".jpeg")<>0) or (instr(rsTabella1("Url"),".png")<>0))  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
         'response.write(rsTabella1("Url"))
		 'response.write "ciao2"
		 %>
		 <a href="<%=rsTabella1("Url")%>" target="_blank"><img src="<%=rsTabella1("Url")%>" border="1"></a> <br>
       <%else
			if ((instr(rsTabella1("Url"),"tp://")<>0) or (instr(rsTabella1("Url"),"tps://")<>0)) and ((instr(rsTabella1("Url"),".htm")<>0) or (instr(rsTabella1("Url"),".html")<>0) or (instr(rsTabella1("Url"),".php")<>0) )  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
				pagina=1
				%>  <a href="<%=rsTabella1("Url")%>" target="_blank"><%=rsTabella1("Url")%></a> <br><%
			else
	   'response.write "ciao"
	  ' response.write("urlimg1="& urlimg)
	  %>

       <img src="<%=urlimg%>" border="1"> <br>
           <%end if%>
	    <%end if%>
     <br>

	 <%end if%>

	 <% if pagina=0 and gdoc="false" then%>
      <a href="../service/cancella_immagine.asp?urldb=<%=rsTabella1("Url")%>&urlimg=<%=urldelete%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceFrase=<%=rsTabellaNew.Fields("CodiceFrase")%>"><img src="../../img/elimina_small.jpg" width="10" height="10" title="Elimina" onClick="return window.confirm('Vuoi veramente cancellare questa immagine?');"></a></p>
     <%end if%>


	<%
		id_txtImg = id_txtImg+1
	   rsTabella1.movenext
	   loop
%>









<%
end if ' prima della fine dell'if andrebbe inserita la possibilità di cambiare il link alle immagini!!
%>



<%if (session("Admin")=true) then %>
 <p><br><input class="input-mini" type="text" name="txtVAL" value="<%=rsTabellaNew.Fields("Voto")%>" size="1"  ><b>
	Valutazione </b> 
  <%end if%>
  <br>
   <b>Risorse</b><a  title="Riguarda la spiegazione" href="<%=urlRisorsa%>" target="_blank">&nbsp;<i class="icon-cloud"></i>&nbsp;&nbsp;&nbsp; </a>
   <br>
   <%
if (rsTabellaNew.Fields("Segnalata")=1)  then
       f="<span style='color:red'>Segnalata</span>"
     Else 
	    if (rsTabellaNew.Fields("Segnalata")=2)  then
           f="<span style='color:green'>Segnalata</span>"
		else
		   f="<span>Segnalata</span>"
		end if
     end if
%>
       <span title="Feedback all'autore"><b><%=f%></b></span>

                                           


										    <% if (rsTabellaNew.Fields("Segnalata")<>0)  then%>
										        	 <INPUT TYPE="RADIO" name="txtSegnalata"  id="txtSegnalata" checked="true" value="1" onclick="segno_segnalazione(0,1);">Si
													<INPUT TYPE="RADIO" name="txtSegnalata"  value="0"  onclick="segno_segnalazione(1,1);">No
													<% else %>
													<INPUT TYPE="RADIO" name="txtSegnalata"  id="txtSegnalata" value="1" onclick="segno_segnalazione(0,1);">Si
													<INPUT TYPE="RADIO" name="txtSegnalata"   checked="true" value="0" onclick="segno_segnalazione(1,1);">No
																	<% end if %><br>
												<div id="divSegno"  style="display:block"><b>Segno</b>
												<% if (rsTabellaNew.Fields("Segnalata")=2)  then%>
													<INPUT TYPE="RADIO" name="txtSegno" checked="true" value="1">(+)
													<INPUT TYPE="RADIO" name="txtSegno"  value="0">(-)
												<% else %>
													<INPUT TYPE="RADIO" name="txtSegno" value="1">(+)
													<INPUT TYPE="RADIO" name="txtSegno"   checked="true" value="0">(-)
											<% end if %>
<p>
 
 <br><textarea class="input-block-level" rows="2" cols="40" name="txtSegnalazione"><%=feedback%></textarea>
 </p>





	  <br>
	</div>
				</div>

 





 <br>


<!-- *** Ripristinare ?
<img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa();">
&nbsp;
-->

 
<input type="button" onclick="inviaDati()" id="btnImg" value="Invia" name="B1" class="btn"> </p>
 



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
/*
 $("#btnImg").click(function(){

	getParametri(1);

 });
*/

  

CKEDITOR.replace('txtS1' );
CKEDITOR.instances.txtS1.on('paste', function(evt) {
evt.cancel();
});	

 function segno_segnalazione(s)
{
if (s==0)
 document.getElementById("divSegno").style.display='block';
 else
   document.getElementById("divSegno").style.display='none';
}


 function checkImg(){
    img1 = document.getElementById("textfield1").value.trim();
	img2 = document.getElementById("textfield2").value.trim();
	img3 = document.getElementById("textfield3").value.trim();

	if(img1.search("http") == -1 && img2.search("http") == -1 && img3.search("http") == -1){
		alert("Devi inserire almeno un url con protocollo http/https");
	}else if(img1.search("http") == -1 && img1 != ""){
		alert("L'immagine deve essere con protocollo http/https");
	}else if(img2.search("http") == -1 && img2 != ""){
		alert("L'immagine deve essere con protocollo http/https");
	}else if(img3.search("http") == -1 && img3 != ""){
		alert("L'immagine deve essere con protocollo http/https");
	}else{
		 
		document.dati.action = document.dati.action+"&Img=1";
		document.getElementById("frmDocument").submit();
	}

	}

function stampa() {
    document.dati.action = "7_stampa_schede_frasi_elenco_una.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&QuerySQL=<%=QueryPrima%>";
		//document.dati.action = "../../home.asp"
		document.dati.submit();
}



function inviaDati() {
	
	   var testo= String(CKEDITOR.instances.txtS1.getData());
	  
		 
		//$('#frmDocument').attr('action', url);
		
		//testo=encodeURIComponent(testo);
		document.getElementById("txtSpiegazione").value=testo;
		let Img = document.getElementById("isImg").value;
		if (Img==1) checkImg();
		document.getElementById("frmDocument").submit(); 
		
		// alert("ciao");
	//	$('#frmDocument').attr('submit');

		// alert(testo);
	 	// params="testo="+testo;
		// xhttp.open('POST', url) 
		// xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')		
		// xhttp.send(params);				 
	//CKEDITOR.instances.editor1.setData(testo);										
	}

 </script>


 </html>
