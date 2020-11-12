<%@ Language=VBScript %>
<% 
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 divid=Session("divid")
  
  id_classe=Session("id_classe")
 
 %>
<!--#include file = "stringa_connessione.inc"--> 

<script src="../SpryAssets/SpryTabbedPanels.js" type="text/javascript"></script>
<link href="../SpryAssets/SpryTabbedPanels.css" rel="stylesheet" type="text/css" />
 <script type="text/javascript">
 
function addsmile(codice) {
	 
		with (document.frmMessage) { 
		 
		 
		  messaggio.value= messaggio.value + codice;
		 
	    }	
}
 

 </script>
 

<%
 
  Dim ID_Chat
  Dim ConnessioneDB,rsTabella,rsTabella0, QuerySQL
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  
   Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
   %>
    <!-- #include file = "stringa_connessione_forum.inc" -->
     <!-- #include file = "../var_globali.inc" -->
   
  
<center>
 
 
<!--#include file="functions/functions_chat.asp"-->

 </center>

  <% 
function inHTML(sReadAll)
   sReadAll=replace(sReadAll,"[color","<font color")
   sReadAll=replace(sReadAll,"[/color]","</font>")  
   sReadAll=replace(sReadAll,"[i]","<i>")
   sReadAll=replace(sReadAll,"[/i]","</i>")
   sReadAll=replace(sReadAll,"[b]","<b>")
   sReadAll=replace(sReadAll,"[/b]","</b>")
   sReadAll=replace(sReadAll,"]",">")
   inHTML=FormatMessage(sReadAll) 
end function



function durata(h,m,s)
 if h>0 then
    durata= h&"h"
 end if
  if m>0 then
    if h>0 then
    	durata=durata&":" & m &"min"
	else
	   durata= m &"min"
	end if
 end if 
 if s>0 then
    if m>0 then
    	durata=m&"min:" & round(m/s) &"sec"
	else
	   durata= s &"sec"
	end if
 end if 
 

end function

   ID_Chat=Request.QueryString("ID_Chat")
  
   QuerySQL="Select * from CHAT_SESSION where ID_Chat=" & ID_Chat &""
   Set rsTabella0 = ConnessioneDB1.Execute(QuerySQL)   
   ore=DateDiff("h",rsTabella0("Inizio"),rsTabella0("Fine")) 
   minuti=DateDiff("n",rsTabella0("Inizio"),rsTabella0("Fine")) 
   secondi=DateDiff("s",rsTabella0("Inizio"),rsTabella0("Fine")) 
 
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & rsTabella0("cartella") & "/Chatlog/" & rsTabella0("Nome")  
   url=Replace(url,"\","/")
   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
   sReadAll = objTextFile.ReadAll
   sReadAll1=sReadAll
   'response.write(url)
	'response.write(inHTML(sReadAll))
	objTextFile.Close
	'registrataresponse.write(url)
	daShowChat2=1 'serve per include che deve selezionare in base all'inclusione da qui o da nuovo messaggio
		%>
 
   
<html>
<head>
 <script src="../lib/prototype.js" type="text/javascript"></script> 
  <script src="../src/scriptaculous.js" type="text/javascript"></script> 
  <script src="../src/unittest.js" type="text/javascript"></script> 
  
   

<link rel="stylesheet" type="text/css" href="../../stile.css">
 
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Chat</title>
 
  

</head>

 
 <body>
 

<div id="bloc_sinistra">
		<div id="bloc_sinistra_int">
			<div id="bloc_sinistra_cont">
			  <div id="logo_space">
                <div class="menu_title">
                  <div id="home_page"> <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx"> </div>
                </div>
			    <div class="menu_cont_one">
			      <div id="comune"><b> <a href="../home.asp"><font color=#000000>HOME PAGE</font></a></b></div>
		        </div>
			    <div class="menu_cont_two"> <img class="imground_sx" src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%"><img src="../../img/code.gif"  width="25%" class="imground_dx"> </div>
		      </div>
			  <div id="logo_space1">
					<p align="center">
					<img src="../../img/umanet2.png" width="90%" >
			  </div>
				
				<%QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					'divid=request.querystring("divid")%>
					
					
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
									<a   href="../../home_ver.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>&cartella=<%=rsTabella.fields("cartella")%>">Verifica</a> 
								</div>	
                                <div class="menu_cont_two"  >
									<a  href="../forum/default.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Forum&nbsp;</a> 
								</div>	
                                <div class="menu_cont_one"  >
									<a class="menu_selected" href="showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">&nbsp;Chat</a>
                                    </div>
                               		
                                     <div class="menu_cont_two"  >
									<a href="../cClasse/studente_domande.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>">Classe</a></div>
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
					 				
				
			</div>
			
			</div>
			</div>

	




<div id="bloc_destra">
		<div id="bloc_destra_int">
			<div id="bloc_destra_cont">
  
	
    <br>
			
	<div class="contenuti_forum" style="width: 700px; height: auto" >	
	 	  
	<center><b><h2>Chat</h2></b></center><br>
    <TABLE WIDTH = 75%  id="zebra" border=1 align="center" bordercolor=pink>
    <thead>
    <TR >
    <Th><B><FONT COLOR = "RED">Titolo</FONT></B></Th>
   
    <Th ALIGN = CENTER><B><FONT COLOR = "RED">Durata</FONT></B></Th>
    <tr style="border-bottom:inset;"><td><b><%=rsTabella0("Titolo")%></b></td>
        
        <td width="20%"><%=durata(ore,minuti,secondi)%></td>   
    </tr>
    <tr style="border-bottom:inset;"><td colspan="2">
	<p><%=Response.write(FormatMessage(sReadAll))%> </p>
    </td>
    </tr></table>
    
   <br><br><center>
  <%if session("Admin")=true then%>
  <a title="Modifica testo" href="#" onClick="Effect.toggle('d2','BLIND'); return false;">Modifica</a> 
   <%end if%> 
    <div id="d2" style="display:none;"><div style="background-color:#ffffff;width:500px;border:1px solid white;padding:10px;"> 
    <form name="frmMessage" action="aggiorna_chat.asp?ID_Chat=<%=ID_Chat%>&nome=<%=rsTabella0("Nome")%>" METHOD = "POST">
     
    <br><b>Titolo : <br></b><br>
    <input type="text" name="txtTitolo" size="40" value="<%=rsTabella0("Titolo")%>"><br>
    <br><b>Messaggio :</b> <br></div>
    <textarea name="messaggio" cols="60" rows="15"><%=sReadAll1%></textarea>
    <br> 
    <p>
 
<center>    
<a href="#" onClick="Effect.toggle('dEmo','BLIND'); return false;">
<img title='Inserisci emoticons' src="smilies/icon-smilie.gif" align="absmiddle" style="border-color:blue">
</a> 
</center>
<center>
<div id="dEmo" style="display:none;">
<div style="background-color:#ffffff;width:400px;border:1px solid blue;padding:5px;align=center;"> 

<!--#include file = "include/Tabbed_Panels.inc"-->

</div></div> 
</center>
¨<br> <center>
       <input type="submit" value="Aggiorna"><br><br><hr style="width:35%"> </center>
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>
     </div></div> 
  
  
 <!-- Chiude l'interfaccia -->
 
   

 
</body>
 
</html>