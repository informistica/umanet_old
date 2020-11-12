<%@ Language=VBScript %>
<!-- #include file = "../extra/test_server.asp" -->
  <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim esecuzione, messaggio
   set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito
    
   
  'response.write("CI SONO")
	
       %> 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non sei autorizzato !")

location.href="../../home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"
//location.href=window.history.back();
 }
 </script>
</head>


   
 <%  Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")%>
	<!--#include file = "controllo_sessione.asp"-->
     <!-- #include file = "stringa_connessione_forum.inc" -->
     <!-- #include file = "../var_globali.inc" -->
    

<%  
                           

'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<%if true then%>
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">
 
<%

 dim objFSO,objCreatedFile
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim sRead, sReadLine, sReadAll, objTextFile
 Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="C:\Inetpub\umanetroot\anno_2012-2013\logAggiorna.txt"
 
 
 ' eseguo per le smile
  
  'url1=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/include/replace_codici.asp"  
  url1=Server.MapPath(homesito)& "/script/ChatRoom/functions/replace_codici_img.asp" 
  url1=Replace(url1,"\","/")
  
  'QuerySQL="Select * from TUTTESMILES2 where CodiceAllievo='" & Session("CodiceAllievo") &"' order by Posizione,Pos;"
  QuerySQL="Select * from TUTTESMILES2  order by Posizione,Pos;"
  	
   
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logGenera.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(url1)
'				objCreatedFile.Close

' apro il file replace_codicei_img per aggiungere  
Set fso = CreateObject("Scripting.FileSystemObject")
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				Set objCreatedFile = fso.OpenTextFile(url1, ForAppending,True)
		        
 
Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
rsTabellaS.movefirst
objCreatedFile.WriteLine("<%")
do while not rsTabellaS.eof 
' distinguo se eseguo in locale e quale tipo diimg devo inserire, per quelle diverse da smiles devo aggiungere un percorso diverso
   if esecuzione.locale = 1 then   
        
	 	 messaggio="strMessage = Replace(strMessage," &chr(34)& rsTabellaS("Codice")&chr(34)& ","&chr(34)&"<a target=_blank href='"&rsTabellaS("Href_L")&"'><img title='"&rsTabellaS("Title")&"' src=../../" &Session("cartella")&"/img_social/img/"&rsTabellaS("Url")&" align=absmiddle></a>"&chr(34)&")" 	
   	    
   else 
	 
	 'new messaggio="strMessage = Replace(strMessage," &chr(34)& rsTabellaS("Codice")&chr(34)& ","&chr(34)&"<a target=_blank href='"&rsTabellaS("Href_O")&"'><img title='"&rsTabellaS("Title")&"' src=../../" &Session("cartella")&"/img_social/img/"&rsTabellaS("Url")&" align=absmiddle></a>"&chr(34)&")" 	
   	    
	'orig messaggio="strMessage = Replace(strMessage," &chr(34)& rsTabellaS("Codice")&chr(34)& ","&chr(34)&"<img src=../../" &Session("cartella")&"/img_social/img/"&rsTabellaS("Url")&" align=absmiddle></a>"&chr(34)&")" 
	 
	 
	 	 messaggio="strMessage = Replace(strMessage," &chr(34)& rsTabellaS("Codice")&chr(34)& ","&chr(34)&"<a target=_blank href='"&rsTabellaS("Href_O")&"'><img class='imground_shadow'  title='"&rsTabellaS("Title")&"'  src=../../" &Session("cartella")&"/img_social/img/"&rsTabellaS("Url")&" align=absmiddle></a>"&chr(34)&")" 
		   
	    
	end if  
    objCreatedFile.WriteLine(messaggio)
	response.write("<br>scrivo:"&messaggio)
  	rsTabellaS.movenext
   loop	
   fine= "" & chr(37) &">"  ' sever per mettere % > la chiusura di vbscript
objCreatedFile.WriteLine(fine)

objCreatedFile.Close 

   
   ' ho scritto la prima parte dei replace, cioè quelli relativi agli smile

' ADESSO DEVO GENERARE IL file tabbed_panel.inc
 

 Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="C:\Inetpub\umanetroot\anno_2012-2013\logAggiorna.txt"
 
 
 ' eseguo per le smile
   url1=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/include/Tabbed_Panels.inc"  
   
  url1=Replace(url1,"\","/")
  
 ' QuerySQL="Select * from CAT_SOCIAL where CodiceAllievo='"&Session("CodiceAllievo") &"' order by Posizione;"
 QuerySQL="Select * from CARTELLESMILES where Cartella='"&Session("cartella") &"' order by Posizione;"
  
  'objFSO.deletefile url1
 
' ho cancellato la vecchia versione
' creo la nuova

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logGenera1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
    If objFSO.FileExists(url1) then
       objFSO.deletefile url1
    Else
        Response.Write "Il file da cancellare NON esiste"
    End If

Set objCreatedFile = objFSO.CreateTextFile(url1, True)
Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
rsTabellaS.movefirst

testoInclude="<div id='TabbedPanels2' class='TabbedPanels'>"
objCreatedFile.WriteLine(testoInclude)
testoInclude=" <ul class='TabbedPanelsTabGroup'> "
objCreatedFile.WriteLine(testoInclude)
 
testoInclude=""

do while not rsTabellaS.eof 

     testoInclude="   <li class='TabbedPanelsTab' tabindex='0'>" & rsTabellaS("Testo") & "</li>"
	 	 	 
    objCreatedFile.WriteLine(testoInclude)
  	rsTabellaS.movenext
   loop	
testoInclude=" </ul> "
objCreatedFile.WriteLine(testoInclude)
testoInclude="<div class='TabbedPanelsContentGroup'>" 
objCreatedFile.WriteLine(testoInclude)
response.write(testoInclude)

rsTabellaS.movefirst 
do while not rsTabellaS.eof 

     
	 testoInclude=" <div class='TabbedPanelsContent'><!--"&chr(35) &"include file ='"&rsTabellaS("Cartella_Cat") & ".inc'--></div>"
	 	 	 
    objCreatedFile.WriteLine(testoInclude)
  	rsTabellaS.movenext
   loop	
  testoInclude="</div>"
  objCreatedFile.WriteLine(testoInclude)
  testoInclude="</div>"
  objCreatedFile.WriteLine(testoInclude)
  
 testoInclude="<script type='text/javascript'>"
 objCreatedFile.WriteLine(testoInclude)
 testoInclude="var TabbedPanels2 = new Spry.Widget.TabbedPanels('TabbedPanels2');"
 objCreatedFile.WriteLine(testoInclude)
 testoInclude="</script>"
 objCreatedFile.WriteLine(testoInclude)
 response.write(testoInclude)
objCreatedFile.Close 

'ADESSO DEVO CREARE TANTI .inc 	PER QUANTE SONO LE CATEGORIE CON I javascript per aggiungere il codice

   QuerySQL="Select * from TUTTESMILES2 where Cartella='"&Session("cartella") &"' order by Posizione" 
  Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
  rsTabellaS.movefirst
  	
  
  
 
  i=0
  do while not rsTabellaS.eof 
  if i=0 then ' per il primo di ogni cat creo il file in cui scriver
    '  url1=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/include/" & rsTabellaS("Cartella_Cat") &".inc"
	url1=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/include/" & rsTabellaS("Cartella_Cat") &".inc"
  url1=Replace(url1,"\","/")
     ' objFSO.deletefile url1
      Set objFSO1 = CreateObject("Scripting.FileSystemObject")
	  If objFSO1.FileExists(url1) then
       objFSO1.deletefile url1
    Else
        Response.Write "Il file da cancellare NON esiste"
    End If
	  Set objCreatedFile = objFSO1.CreateTextFile(url1, True)
    '  response.write(url1)
	  response.write(url1)
	  testo=rsTabellaS("Testo") ' serve per capire quando cambia categoria per generare nuovo file
	  
  end if

	' messaggio="<img width='50' height='50'  src=../img_social/" &rsTabellaS("Cartella_Cat")&"/" & rsTabellaS("Url")&" align=absmiddle  title='" & rsTabellaS("Codice") &"' onclick='javascript:addsmile(" &chr(34)& rsTabellaS("Codice") &chr(34)&")'> "
	
	messaggio="<"&chr(37) & "if daShowChat2=0 then" &chr(37)&">" 
    objCreatedFile.WriteLine(messaggio) 
	messaggio="<a href=Javascript:postmessage.AddSmileyIcon('" & rsTabellaS("Codice") &"');><img width='50' height='50' src='../../" & Session("cartella")&"/img_social/thumb/"& rsTabellaS("Url")&"'></a>" 
    objCreatedFile.WriteLine(messaggio) 
    messaggio="<"&chr(37) & "else" &chr(37)&">" 
    objCreatedFile.WriteLine(messaggio) 
	
	 messaggio="<a href=Javascript:addsmile('" &rsTabellaS("Codice") &"');><img border=0 width='50' height='50'  title='"&rsTabellaS("Codice")&"' src=./../" & Session("cartella")&"/img_social/thumb/"& rsTabellaS("Url")&"></a>" 
	 objCreatedFile.WriteLine(messaggio)
	 messaggio="<"&chr(37) & "end if" &chr(37)&">" 
    objCreatedFile.WriteLine(messaggio) 
	
	
	
     
	'response.write(messaggio)
  	rsTabellaS.movenext
	i=i+1
	if not rsTabellaS.eof then
		if strcomp (testo,rsTabellaS("Testo"))=0 then
		 ' objCreatedFile.Close 
	      ' response.write("UGUALE A 0" & Testo & "="& rsTabellaS("Testo") )
		 
	     else
		    i=0
			testo=rsTabellaS("Testo")
			'response.write("DIVERSO DA 0 CAMBIO")
		end if
	end if
	
   loop	

objCreatedFile.close
 Set objFSO = Nothing
  Set objFSO1 = Nothing








				
'				
		 
 
On Error Resume Next
If Err.Number = 0 Then
		Response.Write "(2) Creazione dei file di inclusione riuscita! "
		%><br><br>
		<a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
		<%
		'Response.Redirect "showChat2.asp?ID_Chat="&ID_Chat
Else
		Response.Write Err.Description 
		Err.Number = 0
End If 
%>
	<center><br><br><font size="3">
 
</center>
<!--#include file = "database_cleanup.inc"-->
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
	</body>
	</html>
	