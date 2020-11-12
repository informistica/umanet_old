<%@ Language=VBScript %>
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../stile.css">
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
 <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("Cancellazione modulo effettuata ")
//location.href="../Home.asp"
//location.href=window.history.back();
 }
 </script>
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella,rsTabella1, QuerySQL,StringaConnessione,URL,RecSet
    Dim cartelle(4)
   cartelle(0)="_Domande"
   cartelle(1)="_Frasi"
   cartelle(2)="_Nodi"
   cartelle(3)="_Spiegazioni"
   
   Id_Mod=Request.QueryString("Id_Mod")
   Id_Classe=request.querystring("Id_Classe")
   Classe=Request.QueryString("Classe")
   
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
 ' mi servirà per cancellare la cartella risorse                             

url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))
url=Replace(url,"\","/")
'response.write(url)
 
 
 
 ' prima devo cancellare tutti i compiti associati al modulo
 
 ' prima devo cancellare tutti i paragrafi associati al modulo, poi cancello il modulo
 QuerySql="SELECT Paragrafi.ID_Paragrafo, Moduli.ID_Mod " &_
" FROM Paragrafi, Moduli, Classi_Moduli_Paragrafi " &_
" WHERE  Classi_Moduli_Paragrafi.Id_Modulo=Moduli.ID_Mod and Classi_Moduli_Paragrafi.Id_Paragrafo=Paragrafi.ID_Paragrafo " &_
" And Moduli.ID_Mod='" & Id_Mod&"';"
set rsTabella1=ConnessioneDB.Execute(QuerySQL)

while not(rsTabella1.eof) 
 QuerySQL ="DELETE   FROM preFrasi WHERE Id_Paragrafo ='" &rsTabella1("ID_Paragrafo")&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preNodi WHERE Id_Paragrafo ='" &rsTabella1("ID_Paragrafo")&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preDomande WHERE Id_Paragrafo ='" &rsTabella1("ID_Paragrafo")&"';"
	 ConnessioneDB.Execute(QuerySQL)
 
 
     QuerySQL ="DELETE   FROM Paragrafi WHERE Id_Paragrafo ='" &rsTabella1("ID_Paragrafo")&"';"
	 rsTabella1.movenext()
   '  response.write(QuerySQL&"<br>")
	 ConnessioneDB.Execute(QuerySQL)
wend
     QuerySQL ="DELETE   FROM Moduli WHERE ID_Mod ='" &Id_Mod&"';"
'response.write(QuerySQL)
	 ConnessioneDB.Execute(QuerySQL)
     Set fso = CreateObject("Scripting.FileSystemObject")   
	 for i=0 to 3 
		url=Server.MapPath(homesito)&"/"&Classe&"/"&Id_Mod&cartelle(i) 
		url=Replace(url,"\","/")
		if fso.FolderExists (url) then
			' response.Write( "La cartella " & url & " esiste già.<br>")
			'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\anno_2012-2013\logTxMod2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(url&"/Img")
'				objCreatedFile.Close      
			 
		    fso.DeleteFolder (url&"/Img")
			fso.DeleteFolder (url) 
			
			'response.Write( "La cartella " & url&"/Img" & " è stata eliminata.<br>") 
		end if
    next 
	  ' dopo le 4 cartelle cqancello anche quella del modulo
     url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
	url=Replace(url,"\","/")
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists (url) then
    fso.DeleteFolder(url)
'response.write("La cartella è stata cancellata : <br> "& url )
'se esiste la cartella risorse la cancello
else

   response.write("La cartella non esiste : <br> "& url )
   end if

'response.write(url)
On Error Resume Next
If Err.Number = 0 Then
Response.Write "Cancellazione del Modulo avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If


 if Request.ServerVariables("HTTP_REFERER") <>"" then %>
  <BODY onLoad="showText2();"> 
<%		response.Redirect request.serverVariables("HTTP_REFERER") 
end if 
%>						 
					 
 
	</body>
	</html>
	