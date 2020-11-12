<%
Session.Abandon
          Response.Cookies("Dati")("Loggato")= ""
		  Response.Cookies("Dati")("Cognome")= ""
		  Response.Cookies("Dati")("Nome")=""
		  Response.Cookies("Dati")("CodiceAllievo")= ""
		  Response.Cookies("Dati")("Username")=""
		  Response.Cookies("Dati")("DataTest")= "" 
		  Response.Cookies("Dati")("Id_Classe")=""
		  Response.Cookies("Dati")("cartella")=""
		  Response.Cookies("Dati")("Cartella")=""
		  Response.Cookies("Dati")("CartellaAdmin")= ""
	      Response.Cookies("Dati")("In_Quiz")= ""
	      Response.Cookies("Dati")("CodAdmin")= ""
		  ' impostate in home.asp
		  
     Response.Cookies("Dati")("Materia")= ""
	 Response.Cookies("Dati")("ID_Materia")= ""
	 Response.Cookies("Dati")("ID_Matsint")= ""  ' mi serve la chiave numerica per il DBMatprof per recuperare la login dell'admin
	 Response.Cookies("Dati")("idxMat")= ""
	 
	 Response.Cookies("Dati")("DBCopiatestonline")= ""
	 Response.Cookies("Dati")("DBForum")= ""
	 Response.Cookies("Dati")("DBLavagna")= ""
	 Response.Cookies("Dati")("DBDiario")= ""
	 Response.Cookies("Dati")("DBDesideri")= ""
	 
	 	'  url=Server.MapPath(homesito)& "/Materie/Materia_"&Session("idxMat")&"/"&Session("Cartella") &"/Profili/sessioni/"&CodiceAllievo&".txt" 
		'  url=Replace(url,"\","/")
	 
		  %><br><%
	 
		 
	
	'Dim objFSO,objCreatedFile
'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
'	Dim sRead, sReadLine, sReadAll, objTextFile
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	 
'	'Create the FSO.
'	' salvo la sessione su file anzichè su cookie
'	Set objFSO = CreateObject("Scripting.FileSystemObject") 	 
'	Set objCreatedFile = objFSO.DeleteFile url 
''
pageset = Request.Cookies("Dati")("DB")
'Response.Cookies("Dati")("DB")=""

doc=Request.Cookies("Dati")("DOC")
'Response.Cookies("Dati")("DOC")=""
if (strcomp(pageset,"2")=0) and (strcomp(doc,"1")=0) then
   Response.Redirect("../../home.asp")
else
  if (strcomp(pageset,"2")=0) then
    Response.Redirect("../../home.asp")
  end if 
 Response.Redirect("../../home.asp")
end if
%>

