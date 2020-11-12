<% 
  Dim objMessage
  Dim rsConfigFile
  Dim txtFileData
	
  Set objMessage=Server.CreateObject("SCRIPTING.FileSystemObject")
  Set rsConfigFile=objMessage.OpenTextFile (Server.MapPath("config.txt"))

  while not rsConfigFile.AtEndOfStream
	txtFileData=rsConfigFile.ReadAll
  wend
  'Response.write "txtFileData="&txtFileData&"<br>"
	
  dim LinkMessage
  LinkMessage = split(txtFileData,"=")

  'SET THE PATH OF WHERE TO STORE THE FILES HERE
  Dim path
  path = LinkMessage(1)
  'response.write("PATH="&path&"<br>")
  dim upd_doc
  dim flnam,pth,NewFileName
  dim flObj,upl
  dim biData
  
  'Now whatever the admin uploads, we have a preset file name for each of
  'them. We now upload this to the specific directory.
  
  set upl = new BinForm
  upl.Read()
  upl.Extensions=Array(".htm")
 ' Response.write "server.mappath(path)="&server.mappath(path)&"<br>"
 ' Response.write "path="&path&"<br>"
  upl.Directory=server.mappath(path)
  upl.Create=false
  upl.OverWrite=true
 
  set flObj=server.createobject("Scripting.FileSystemObject")
 'response.write("FileName="&upl.form("flname"))
   if len(upl.form("flname")) > 0 then

	 if flObj.FileExists(server.mappath("path"&flnam)) then
		flObj.DeleteFile(server.mappath("path"&flnam))
	 end if
	 upl.Save("flname")
     if upl.Success then 
         Response.Write "<p>&nbsp;<p>Documento caricato correttamente."
		 ' se sono stato chiamato da carica risorsa modulo modifica_modulo.asp aggiorno il db
		
     else
         Call ShowErr()
     end If 

     set upl=nothing
     set flObj=nothing

     sub ShowErr()
		 'Response.Write "There was an error uploading this file. Please try again.<p>Possible causes are:<p>"
'    	 response.write "<li>File size exceeds 150kb. Please ensure the file you are uploading is lesser than 150 kb.</li></ul><p>"
'		 response.write "Possible server causes:<p>"
'		 response.write "<ul><li>The directory to which the file is being uploaded does not exist - /includes/sitedocs</li>"
'		 response.write "<li>The posted data was not read due to problem with form posting</li></ul><p>"
'		 response.write "<input type=""button"" name=""btnBack"" value=""Back"" onClick=""javascript:location.href='default.asp';"">"
       response.write " Errore di caricamento "
	 end sub
%>