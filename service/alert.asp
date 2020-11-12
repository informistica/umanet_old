<html>
<head>
	<script type="text/javascript" src="../AlertMessage/js/ResponseManager.js"></script>
	<script type="text/javascript" src="../AlertMessage/js/prototype.js"></script>
	<script type="text/javascript" src="../AlertMessage/js/scriptaculous.js"></script>
	<link rel='stylesheet' type='text/css' href='../AlertMessage/style/style.css' />
	<link rel="stylesheet" type="text/css" href="../../stile.css">
</head>
<body>
	<script type="text/javascript">
	ResponseManager.Startup();
		 
	ResponseManager.setBoundsDiv("400","180","18");
	</script>
	Premi sulla scritta per vedere l'effetto.. 
	<ul>
		<li onClick="ResponseManager.PrintAlertDiv(1,'sono il messaggio OK');" style="cursor:pointer;">Messaggio OK</li>
		<li onClick="ResponseManager.PrintAlertDiv(0,'sono il messaggio KO');" style="cursor:pointer;">Messaggio ERROR</li>
	</ul>
	

	
	<%@ Language=VBScript %>
	<% response.write("Ciao")%>
 
   <% Response.Buffer=True 
      Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome,Modulo
    
 
  Cartella=Request.QueryString("Cartella")
  DataTest = Request.Cookies("Dati")("DataTest")
  CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Nome=Request.Cookies("Dati")("Nome")
  Cognome=Request.Cookies("Dati")("Cognome")
  
  CodiceTest = Request.QueryString("CodiceTest") 
  Response.Cookies("Dati")("CodiceTest")=CodiceTest
  Capitolo=Request.QueryString("Capitolo") 
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
 

%>
 <div id="container">
<div id="bloc_destra_cont">
<div id="bloc_sinistra_login">
<div class="contenuti_login" >


<b>
	       <br>
	   <p align="center">    <font color=#FF0000 face ="Verdana" size="3"><b>Scegli l'esercitazione sul test </font> <br> <font color=black face ="Verdana" size="3"> <%Response.write (Paragrafo) %></b></font> <br><p>
  <div class="immagini">  
   <FIELDSET>
   <LEGEND><B>QUIZ A RISPOSTA SINGOLA </B></LEGEND>
            <h4><a href="../cDomande/inserisci_test.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Inserisci</a></h4> <!-- se il login è corretto richima la pagina per inserire le domande del test -->
	  <h4><a href="../cDomande/esegui_test.asp?Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>">Esegui</a></h4> 
	 
	   <h4><a href="../cDomande/esegui_test.asp?Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>">Verifica</a></h4> 
	  
  <h4><a href="../cClasse/studente_quiz.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modifica</a></h4>
  
  
  
  
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND><B>Admin</B></LEGEND>
      <h4><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a></h4>
	  
	  <h4><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modifica </a></h4>
	  </FIELDSET>
      <% end if
  %> 
</FIELDSET>



</div>


<div class="immagini">  
   <FIELDSET>
   <LEGEND><B>QUIZ IMMAGINI  </B></LEGEND>
             

	  <h4><a href="../cDomande/esegui_test_immagini.asp?CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>">Esegui</a></h4> 
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND><B>Admin</B></LEGEND>
      <h4><a href="../cAdmin/mescola_test_img.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a></h4>
      <% end if
  %></FIELDSET>
</FIELDSET>
</div>








</div>
</div>
  </div>
  </div>
	  
 </body>
</html>