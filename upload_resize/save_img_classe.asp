<%@ Language=VBScript %>

  <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
    
       %> 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
  
</head>
<%

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<!--#include file = "../service/controllo_sessione.asp"-->
    <!-- #include File="resizecheck.asp" -->
     
<!-- #include file = "../var_globali.inc" -->
  

<%  

 
'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
 
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">
 
<%
nomeimg=Session("CodiceAllievo") &".jpg"
 
imgPathDir=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/img" 
imgPathDir=Replace(imgPathDir,"\","/")
thumbPathDir=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/thumb" 
thumbPathDir=Replace(thumbPathDir,"\","/")
img=request.QueryString("img")

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
	'imgPath=imgPathDir  
   'thumbPath=thumbPathDir  
    
 
				 
			
 	
  
				
 
   QuerySQL="  UPDATE Classi SET  Url_img= '" &  img & "' where ID_Classe='"&Session("Id_Classe") &"';"   
  ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logProfilo.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
   ConnessioneDB.Execute (QuerySQL) 
   ' response.write(QuerySQL &"<br>"&img)
   
'Rinomino e cancello le immagini ed i thumb
destinazione=imgPath 
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url=server.MapPath("../../3PC/img_social/img/LogFileName.txt")
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(imgPath & "<br>" & destinazione)
'				objCreatedFile.Close
'Set fso = CreateObject("Scripting.FileSystemObject")
'set  fso.MoveFile (old_path,new_path da provare

'Set fso = CreateObject("Scripting.FileSystemObject")
'set OggFile = fso.GetFile (imgPath)
'OggFile.Copy destinazione,true
'OggFile.Delete
' 
'destinazione=thumbPath 
'Set fso = CreateObject("Scripting.FileSystemObject")
'response.Write(imgPath)
'set OggFile = fso.GetFile (thumbPath)
'OggFile.Copy destinazione,true
'OggFile.Delete
	 

 
 
'response.write(url)
 
On Error Resume Next
If Err.Number = 0 Then
        session("uploaded")=true
			if Request.ServerVariables("HTTP_REFERER") <>"" then 
					  response.Redirect request.serverVariables("HTTP_REFERER")
		end if
		Response.Write "<br>Caricamento avvenuto! : " & Session("FileName")  
		'Response.Redirect "ShowMessage.asp?ID="&ID
Else
		Response.Write Err.Description 
		Err.Number = 0
End If 
%>
	<center><br><br> 
 
</center>
<!--#i<!--i nclude file = "database_cleanup.inc"-->
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	 
	</body>
	</html>
	