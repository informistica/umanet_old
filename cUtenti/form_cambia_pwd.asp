<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%
  
  CodiceAllievo=Request.QueryString("CodiceAllievo")
   id_classe=Request.QueryString("id_classe")
   classe=Request.QueryString("classe")
   Session("Cartella")=classe
  divid=Session("divid")
  
  'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    
%>   
  <!-- #include file = "../service/controllo_sessione.asp" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../var_globali.inc" -->
 
 
 <%
' VERIFICHIAMO SE L'UTENTE E' IDENTIFICATO (LOGGATO)

IF Session("Loggato") = True then

QuerySQL="Select * from Allievi where CodiceAllievo='"&CodiceAllievo& "';" 
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

End IF


Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
 
 

%>

  
 
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
<title>Dati personali</title>
<script src="../lib/prototype.js" type="text/javascript"></script> 
<script src="../src/scriptaculous.js" type="text/javascript"></script> 
<script src="../src/unittest.js" type="text/javascript"></script>
<script>
function PopUpWindow(w,h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

	window.open('../upload_resize/ex2_imgprofilo.asp','../upload_resize/ex2_imgprofilo.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl)
}
// -->
 

	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>

 <script language="javascript" type="text/javascript" >
 function validate2() {
	 
	 if (frmDocument.txtCognome.value=="")
	{
	   alert("Non hai inserito il cognome");
	   frmDocument.txtCognome.setfocus();
	   return 0;
	}
 else
 if (frmDocument.txtNome.value=="")
	{
	   alert("Non hai inserito il nome");
	   frmDocument.txtNome.setfocus();
	   return 0;
	}
 else
 if (frmDocument.txtCodiceAllievo.value=="")
	{
	   alert("Non hai inserito il vecchio username ");
	   frmDocument.txtCodiceAllievo.setfocus();
	   return 0;
	}
	else
	if (frmDocument.txtPwdAllievo.value=="")
	{
	   alert("Non hai la vecchia password");
	   frmDocument.txtPwdAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewCodiceAllievo.value=="")
	{
	   alert("Non hai inserito il nuovo username ");
	   frmDocument.txtNewCodiceAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd.value=="")
	{
	   alert("Non hai inserito la nuova password");
	   frmDocument.txtNewPwd.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd1.value=="")
	{
	   alert("Non hai confermato la nuova password");
	   frmDocument.txtNewPwd1.setfocus();
	   return 0;
	}else
	{
		
		 document.frmDocument.action = "modifica_pwd.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>";  
		
	   
		document.frmDocument.submit();
		
	 
    }
	
}


function validate3() {
	 
	 if (frmDocument1.txtNewEm.value=="")
	{
	   alert("Non hai inserito la nuova email");
	   frmDocument1.txtNewEm.setfocus();
	   return 0;
	}
 else
 if (frmDocument1.txtNewEm1.value=="")
	{
	   alert("Non hai confermato la nuova email");
	   frmDocument1.txtNewEm1.setfocus();
	   return 0;
	}
 else
 if (frmDocument1.txtNewEm1.value !=  frmDocument1.txtNewEm.value)
	{
	   alert("Le due email non corrispondono ");
	   frmDocument1.txtEm1.setfocus();
	   return 0;
	}
	else
	{
		
		 document.frmDocument1.action = "modifica_contatti.asp";  
		
	   
		document.frmDocument1.submit();
		
	 
    }
	
}
 </script>
  
  
 
  
</head>
<body bgcolor="#FFFFFF">
<div id="container">

<center>
 


<div id="bloc_sinistra">
	<div id="bloc_sinistra_int" style="margin-top:35px;margin-left:-5px;">
		
   <div style="margin-left:-5px;">	
            
                    <div id="logo_space">
                        <div class="menu_title"><div id="home_page">
                            <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx">
                            </div></div>
                            <div class="menu_cont_one"><div id="comune"><b>
                                <a href="../../home.asp"><font color=#000000>HOME PAGE</font></a></b></div></div>
                            <div class="menu_cont_two">
                                <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx">
                      </div>
      </div>
                        <div id="logo_space1">
                        <p align="center">
                        <%
    
    ' HO CREATO UNA CLASSE
    ' per sapere se sono in esecuzione sul server o in locale, serve per distinguere gli url per le risorse
     ' pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
    '  if (left(pathEnd1,10)="c:\inetpub") then
    '     locale=1
    '  else
    '     locale=0
    '  end if 	
    ' 	
                
                     %>
                             <a href="https://www.umanet.net/informistica/UWWW/Benvenuto.html" target="_blank"><img src="../../img/umanet2.png" width="90%" ></a></div>
                        
                        
                          
                <%
                    
    
           
                     
                       
                       'Apertura della connessione al database
                       Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>   
                       <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
                        
                        <%
                            id_classe=Session("Id_Classe")
                            QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
                            
                            
                    '    dim objFSO,objCreatedFile
'    				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'    				Dim sRead, sReadLine, sReadAll, objTextFile
'    				Set objFSO = CreateObject("Scripting.FileSystemObject")
'    				url="C:\Inetpub\umanetroot\anno_2012-2013\log.txt"
'    				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'    				objCreatedFile.WriteLine(QuerySQL) 
'    				objCreatedFile.Close
                        
                        
                    
                        Set rsTabella = ConnessioneDB.Execute(QuerySQL)
                        'divid=request.querystring("divid")
                        cartella=rsTabella.fields("Cartella")%>
                        
                            <div class="menu_sinistra">
                                    
                                  <div class="menu_title"><div id="<%=divid%>"><%=rsTabella.fields("Classe")%></div>
                              </div>
                                    <div class="menu_cont_one">
                                       <a href="../lavagna/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Lavagna&nbsp;</a>
                                    </div>
                                    <div class="menu_cont_two"  >
                                        <a   href="../cClasse/home_app.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>">Apprendimento</a>
                                    </div>	
                                        
                                    <div class="menu_cont_one"  >
                                        <a href="../../home_ver.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>&cartella=<%=rsTabella.fields("cartella")%>">Verifica</a> 
                                    </div>	
                                    <div class="menu_cont_two"  >
                                        <a href="../forum/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Forum&nbsp;</a> 
                                    </div>	
                                    <div class="menu_cont_one"  >
                                        <a href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">&nbsp;Chat</a>
                              </div>
                                        
                                         <div class="menu_cont_two"  >
                                        <a class="menu_selected" href="../cClasse/studente_domande.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>">Classe</a></div>
      </div>	
                             
                            </p>
                            
                            
                            
                            
                            <div class="menu_sinistra">
                            <div class="menu_title"><div id="quintacom">U-ECDL</div></div>
                            <div class="menu_cont_one">
                                <a href="../cClasse/home_uecdl_app.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>">Apprendimento</a></div>
                            <div class="menu_cont_two">
                                <a href="../../U-ECDL/home_uecdl_ver.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>">Verifica</a></div>	
                    </div>
                    <div class="menu_sinistra">
                            <div class="menu_title"><div id="quarta">GESTIONE</div></div>
                            <div class="menu_cont_one">
                                <a href="../service/logout.asp">Logout</a></div>
                                
                            <%if (session("Admin")=true) then %>
                            <div class="menu_cont_two">
                            <a href="../cClasse/studente_domande_gruppi.asp">Gruppi</a>
                            </div>
                            
                        
                             <div class="menu_cont_one">
                            <a href="../cAdmin/admin.asp?Id_Classe=<%=id_classe%>&divid=<%=divid%>">Admin</a>
                            </div>
                             
                            <%end if %>
                    </div>
                    
                        <%
                        rsTabella.Close()
                        Set rsTabella = nothing
                    
                    
                       QuerySQL="Select * from Allievi where CodiceAllievo='"& CodiceAllievo&"';"
	'response.write(QuerySQL)  
	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)       
                        
                        %>									
    
                         
  </div>
</div>
        </div>
        </div>
</div>
 

<form method="POST" action="modifica_profilo.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
  
  <div id="bloc_destra_cont">
   
   
  <b><span class="sottotitoloquaderno" style="font-size:18px; font-weight:100">Dati personali</b></font></b><p>
	<div id="bloc_sinistra_login">
<div class="contenuti_login" style="width:auto;">	
<br>	
  </p>
  
  <form method="POST" action="modifica_contatti.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>" name="frmDocument1" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <center>

  <%
  
  
    url= "../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
      
	   url=Replace(url,"\","/")
	   
	  
  if strcomp(rsTabella("Url_img")&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è = a ""
  
  	urlimg=url&"/"&"profilo_vuoto_thumb.png" 
	  
	%>	
    <fieldset style="width:15%"><img class="imground" src="<%=urlimg%>" ><br><b><%=Cognome & " " & Nome %></b></fieldset><br>
    <% 'response.write(urlimg)%>
   
  <%else%>
  <% 
    urlimg=url&"/"& rsTabella("Url_img") ' aggiungo al percorso il nome del file
  
 'response.write(urlimg)%>
    <fieldset style="width:24%"> <img class="imground" src="<%=urlimg%>" ><br><br> <b><%=rsTabella("Cognome") & " " & rsTabella("Nome") %></b></fieldset> <br>
 </center>   
<%end if %>
  
  
  
      <fieldset><legend> <a href="#" onClick="Effect.toggle('profilo','appear'); return false;">Modifica Profilo</a> </legend>
<div id="profilo" style="display:none;"><div style="width:570px;padding:10px;"> 
<p> 
    <span class="sottotitoloquaderno" style="font-size:14px; font-weight:100">Modifca Profilo</b></font></b><p></p></span>
 
  <%
  
      
	 

 
  
   if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then 
	%>
    <span class="sottotitolo"><a href="#" onClick="javascript:PopUpWindow(550,200);return false;"> +Foto</a>
</span>  
   <%end if%> <br>
    <p><b> Mi piace <br>
	</b><input type="text" name="mipiace" value="<%=rsTabella1("Mipiace")%>" size="80" maxlength="100"></p>   	
   <p><b> Non mi piace<br>
	</b><input type="text" name="nonmipiace" value="<%=rsTabella1("Nonmipiace")%>" size="80" maxlength="100"></p> 
 <br><b>Descriviti</b><br>
  <textarea rows="6" cols="50" name="descriviti"> 
  <%=response.write(rsTabella1("Descriviti"))%> 
   </textarea>
 </p>
 
   
 
    
    <%
	if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then 
	%>
     <!-- crea la variabile di tipo inputbox avente un certo nome -->
   <fieldset><legend> Sei veramente tu ? </legend>
    <p><input type="text" name="username" value="" size="40"><b> Username
	</b></p> 
    <p><input type="password" name="password" value="" size="40"><b> Password
	</b></p> 
  </fieldset>
    <p><input type="submit" value="Invia" name="B1"> 
   
</form> <!-- Chiude l'interfaccia -->

</div></div> <!-- Chiudetendina -->
<br><br>

 
<form method="POST" action="modifica_pwd.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>" name="frmDocument" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 
<fieldset><legend> <a href="#" onClick="Effect.toggle('login','appear'); return false;">Modifica dati Login</a> </legend>
<div id="login" style="display:none;"><div style="padding:10px;"> 
<p> 
 

 
    <div class="contenuti_login" >	
    
  </p>
  <span class="sottotitoloquaderno" style="font-size:14px; font-weight:100">Login</b></font></b><p></p></span>
   
  <p><input type="text" name="txtCognome" value="<%=rsTabella1("Cognome")%>" size="20"><b> Cognome 
	</b></p> 
    <p><input type="text" name="txtNome" value="<%=rsTabella1("Nome")%>" size="20"><b> Nome 
	</b></p>   	
   <p><input type="text" name="txtCodiceAllievo" value="<%=rsTabella1("CodiceAllievo")%>" size="20"><b> Username 
	</b></p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
  <p>
   <% if session("Admin")=true then%>
    <input type="text" name="txtPwdAllievo" value="<%=rsTabella1("Password")%>" size="20">
   <% else%>
      <input type="password" name="txtPwdAllievo" value="" size="20">
   <%end if%>
  <b> Vecchia Password </b><b></b></p> 
	 
	 <p><input type="text" name="txtNewCodiceAllievo"     size="20">
  <b> 
	Nuovo Username &nbsp;&nbsp; </b></p> 
	
     <p><input type="password" name="txtNewPwd"     size="20">
  <b> 
	Nuova Password &nbsp;&nbsp; </b></p>
	<p><input type="password" name="txtNewPwd1"     size="20">
  <b> 
	Conferma Nuova Password &nbsp;&nbsp; </b></p>
   <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then 
	%>
    <p><input type="button" value="Invia" name="B2" onClick="return validate2();"> 
    <%end if%>
</form> <!-- Chiude l'interfaccia -->
	
 <%end if%>
</div>

</div></div> <!-- effeto tendina-->
</fieldset>
<br>


<form method="POST" action="modifica_contatti.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>" name="frmDocument1" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 
<fieldset><legend> <a href="#" onClick="Effect.toggle('contatti','appear'); return false;">Modifica Contatti</a> </legend>
<div id="contatti" style="display:none;"><div style="width:auto;padding:10px;"> 
<p> 
 

 
    <div class="contenuti_login" >	
    
  </p>
  <span class="sottotitoloquaderno" style="font-size:14px; font-weight:100">Modifica i dati di contatto</b></font></b><p></p></span>
    <input type="hidden" name="txtCodiceAllievo" value="<%=rsTabella1("CodiceAllievo")%>" size="20">
   <% if strcomp(rsTabella("Email")&"","")=0 then %>
  
  <p><input type="text" name="txtEm" value="Nessuna" size="40"><b> Email 
	</b></p> 
   <%else%>
   <input type="text" name="txtEm" value="<%=rsTabella1("Email")%>" size="40"><b> Email 
   <%end if%>
  <b> 
	 
     <p><input type="text" name="txtNewEm"     size="40">
  <b> 
	Nuova Email &nbsp;&nbsp; </b></p>
	<p><input type="text" name="txtNewEm1"     size="40">
  <b> 
	Conferma Nuova Email &nbsp;&nbsp; </b></p>
   <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then 
	%>
    <p><input type="button" value="Invia" name="B2" onClick="return validate3();"> 
    <%end if%>
</form> <!-- Chiude l'interfaccia -->
	
 
</div>

</div></div> <!-- effeto tendina-->
</fieldset>
<br>

 
   
  
	<% if session("Admin")=true then%>
    
    <fieldset><legend> <a href="#" onClick="Effect.toggle('impostazioni','appear'); return false;">Impostazioni</a> </legend>
<div id="impostazioni" style="display:none;"><div style="width:auto;padding:10px;"> 
<p> 
 
    <fieldset><Legend>Eccezioni alle scadenze</Legend>
      <div class="contenuti_login" >	
        <form action="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&Id_Stud=<%=CodiceAllievo%>" method="post">
         <input type="submit" value="Aggiungi">
        </form>
        
       </fieldset>
       </div><br>
    
        <form action="../cClasse/cancella_studente.asp?CodiceAllievo=<%=CodiceAllievo%>" method="post">
         <input type="submit" value="Cancella questo utente" onClick="return window.confirm('ATTENZIONE Vuoi veramente cancellare questo utente e tutti i suoi dati?');" >
        </form>
    <% end if%>
  </div></div>





</div>
</div>
</div>

</div>
</body>
</html>