<%@ Language=VBScript %>

  <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
    
       %> 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi modificare i dati degli altri studenti!")

location.href="ShowMessage.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>"
//location.href=window.history.back();
 }
 </script>
</head>
<%

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>
   

<!-- #include File="resizecheck.asp" -->
     

   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<!-- #include file = "../var_globali.inc" -->
<!-- #include file = "../service/controllo_sessione.asp" -->
  

<%  
 
 
'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
 <% if true then %>
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">
 
<%
'Id_Cat=Request("txtCategoria")
nomeimg=Request("nomeimg")
if nomeimg="" then
nomeimg="Senza nome "
end if
linkto=Request("linkto")
if linkto<>"" then
isHTTP=left(linkto,7)
  if strcomp("https://",isHTTP)<>0 then
      linkto = "https://"& linkto 
  end if
else
  linkto="#"
end if
'title=Request("txtTitle")
'descrizione=ltrim(Request("txtDescrizione"))
 
imgPathDir=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_"&Session("social")&"/img" 
imgPathDir=Replace(imgPathDir,"\","/")
thumbPathDir=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_"&Session("social")&"/thumb" 
thumbPathDir=Replace(thumbPathDir,"\","/")
 
'se uso la libreria asp.net al posto di Session("FileName") devo usare un file in cui scrive upload.aspx
If Not CheckResizeLib Then
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sReadAll
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objCreatedFile = objFSO.OpenTextFile(server.MapPath("FileName.txt"), ForReading)
	sReadAll = objCreatedFile.ReadLine
	objCreatedFile.Close
	imgPath=imgPathDir &"/" & sReadAll
	thumbPath=thumbPathDir &"/" & sReadAll
else
   imgPath=imgPathDir &"/" & Session("FileName")
   thumbPath=thumbPathDir &"/" & Session("FileName")
End if
imgPath=Replace(imgPath,"\","/")
thumbPath=Replace(thumbPath,"\","/")

 
   QuerySQL="  INSERT INTO IMG_FORUM (CodiceAllievo)  SELECT '" & Session("CodiceAllievo")  & "';"   
   ConnessioneDB.Execute (QuerySQL) 
   response.write(QuerySQL)
  ' QuerySQL="select max (ID_Smile), Pos from IMG_FORUM group by Pos;"
  QuerySQL="select max (ID_Smile)  from IMG_FORUM ;"
   response.write(QuerySQL)
   set rsTabella=ConnessioneDB.execute(QuerySQL)
   MAXID=rsTabella(0)
   MAXPOS=1
   'MAXPOS=rsTabella(1)
   ' nome del file di destinazione 
  
   url=MAXID&".jpg"
  ' codice=":;"&MAXID
   QuerySQL ="UPDATE IMG_FORUM SET Url = '" & url & "', Pos = " & MAXPOS+1 & ",  Nome = '" & nomeimg & "', Href_O = '" & linkto &"' WHERE ID_Smile =" &MAXID&";"
     response.write(QuerySQL &"<br>")
   ConnessioneDB.execute(QuerySQL)

'messaggio=FormatMessage(messaggio)
' rinomino i file dell'immagine grande copiando e cancellando	
' response.write(imgPath)
'imgPath="../../3PC/img_social/img/imgdomanda.jpg"
destinazione=imgPathDir&"/"&MAXID&".jpg"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url=server.MapPath("../../3PC/img_social/img/LogFileName.txt")
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(imgPath & "<br>" & destinazione)
'				objCreatedFile.Close

Set fso = CreateObject("Scripting.FileSystemObject")
set OggFile = fso.GetFile (imgPath)
response.Write(imgPath)
OggFile.Copy destinazione,true
OggFile.Delete


'response.write(filePath)
' ora per il thumb
'imgPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/thumb"  
'imgPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/thumb" 
'imgPath=Replace(imgPath,"\","/")
'filePath=imgPath &"/" & Session("FileName")
destinazione=thumbPathDir&"/"&MAXID&".jpg"
Set fso = CreateObject("Scripting.FileSystemObject")
response.Write(imgPath)
set OggFile = fso.GetFile (thumbPath)
OggFile.Copy destinazione,true
OggFile.Delete
' salvo il nome dell'immagine che poi inserisco nella forum_messages per collegare il post all'immagine
Session("Caricata")=true
Session("NomeImgForum")=MAXID&".jpg"

''txtPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/img" 
''txtPath=Replace(txtPath,"\","/")
'' fle di testo con la descrizione della immagine
'txtPath=imgPathDir
'txtPath=txtPath &"/" &MAXID &".txt"
'
'Set objCreatedFile = fso.CreateTextFile(txtPath, True)
'' Write a line with a newline character.
'objCreatedFile.WriteLine(descrizione)
''Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
'objCreatedFile.Close




	 

' 
 
'response.write(url)
 
On Error Resume Next
If Err.Number = 0 Then
        session("uploaded")=true
			if Request.ServerVariables("HTTP_REFERER") <>"" then 
					  response.Redirect request.serverVariables("HTTP_REFERER")
		end if
		Response.Write "<br>Caricamento avvenuto! : " & Session("FileName") &" in " & MAXID&".jpg"
		'Response.Redirect "ShowMessage.asp?ID="&ID
Else
		Response.Write Err.Description 
		Err.Number = 0
End If 
%>
	<center><br><br><font size="3">
 
</center>
<!--#i<!--i nclude file = "database_cleanup.inc"-->
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
</body>
	</html>
	