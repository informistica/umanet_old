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

 
'Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
' 
''Create the FSO.
'url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/db_" & Session("Id_Classe") & ".txt" 
'url=Replace(url,"\","/")
''response.write("ulr spiegazio="&url)
'
'Set objTextFile = objFSO.OpenTextFile(url, ForReading)
'Session("DBCopiatestonline") = objTextFile.ReadAll
'objCreatedFile.Close
'Set OggFile = objFSO.GetFile (url)
'OggFile.Delete 
'Session("DBCopiatestonline") = Request.QueryString("db")
 
'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
 
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">
 
<%

'stronzo=request.queryString("db")
'response.Write("DBstronzo="&stronzo)
'response.write("DB1="&session("DBCopiatestonline1"))
session("DBCopiatestonline")=session("DBCopiatestonline1")


 
    
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   
   
%>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	
	<% 

	if session("Registrati")=true then
else
	' già provato a togliere ma non va dice semore sessione scaduta di ritorno da ulpoad imgprofilo
%>
<!-- #include file = "../service/controllo_sessione.asp" -->
  
<% end if%>
    
    
    <!-- #include File="resizecheck.asp" -->
     
<!-- #include file = "../var_globali.inc" -->
  


<%

nomeimg=Session("CodiceAllievo") &".jpg"
 
imgPathDir=Server.MapPath(homesito)& "/DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/img" 
imgPathDir=Replace(imgPathDir,"\","/")
thumbPathDir=Server.MapPath(homesito)& "/DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/thumb" 
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
    
 
				 
			
 	
  
				
 
   QuerySQL="  UPDATE Allievi SET  Url_img= '" &  img & "' where CodiceAllievo='"&Session("CodiceAllievo") &"';"   
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
	 
%>
<!-- <script language="javascript">
		opener.location.reload();
		window.close();
	</script>-->
<% 
'response.write(url)
 
On Error Resume Next
If Err.Number = 0 Then
        session("uploaded")=true
		' lo commento perchè altrimenti ottengo sessione scaduta al rientro dall'upload
			'if Request.ServerVariables("HTTP_REFERER") <>"" then 
				'	  response.Redirect request.serverVariables("HTTP_REFERER")
		'end if
		Response.Write "<br><font color='green'>Caricamento avvenuto!</font> <br><b><b>Clicca su Home e rientra nella tua classe </b>  " & Session("FileName")  
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
	