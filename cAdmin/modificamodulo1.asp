 <%@ Language=VBScript %>
 <%
  Response.Buffer=True 
   Dim ConnessioneDB, rsTabella, QuerySQL  
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   
 %>

  <!--#include file="Upload/upload.asp"-->
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
<% 
  
  
  ID_Mod=Request.QueryString("ID_Mod")
  Classe=Request.QueryString("Classe")
  divid=Request.QueryString("divid")
  Id_Classe=Request.QueryString("Id_Classe")
  urlrisorsa=Request.QueryString("URLRISORSA")
  posMod=Request.QueryString("posMod")
  posizioneMod=Request.form("txtposMod")
  byUmanet=Request.QueryString("byUmanet")
  
 
 if (urlrisorsa <>"") or (posMod<>"") then
   if (urlrisorsa <>"") then
    urldb=Request("txtURL_OL")
     QuerySQL="  UPDATE Moduli SET URL = '" & urldb & "',URL_OL='" &urldb&"'" &_
    " WHERE ID_Mod='" & ID_Mod & "';"
    
       ' response.write(QuerySQL)
		   ConnessioneDB.Execute QuerySQL 
     end if
	 if (posMod <>"") then
    
	if byUmanet<>"" then
	
	QuerySql="SELECT MODULI_CLASSE_UMANET.Titolo, MODULI_CLASSE_UMANET.Posizione, MODULI_CLASSE_UMANET.Id_Classe,MODULI_CLASSE_UMANET.ID_Mod" &_
		" FROM MODULI_CLASSE_UMANET WHERE (((MODULI_CLASSE_UMANET.Id_Classe)='"&Id_Classe&"')) ORDER BY MODULI_CLASSE_UMANET.Posizione;"
	else
	
  	QuerySql="SELECT MODULI_CLASSE.Titolo, MODULI_CLASSE.Posizione, MODULI_CLASSE.Id_Classe,MODULI_CLASSE.ID_Mod" &_
		" FROM MODULI_CLASSE WHERE (((MODULI_CLASSE.Id_Classe)='"&Id_Classe&"')) ORDER BY MODULI_CLASSE.Posizione;"
		
	end if	
		'response.write(QuerySql & " " &Paragrafo)
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		i=0
		 do while not rsTabella.eof
		  visibile=Request("txtVisibile"&i)
		  if visibile="" then 
			 visibile=1
		  end if
		 
			   QuerySQL="  UPDATE Moduli SET  Posizione = " & Request.form("txtPosMod"&i)  & ", Visibile = " & visibile  & " WHERE ID_Mod='" & rsTabella("ID_MOD")& "';"
				'response.Write(QuerySQL&"<BR>")
				ConnessioneDB.Execute QuerySQL 
				 
				i=i+1
			rsTabella.movenext
		loop 
     end if
	 
	     
		 response.Redirect "modificamodulo.asp?Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&byUmanet="&byUmanet
  
 else
  
  
  Response.Expires=0
  Response.Buffer = TRUE
  Response.Clear
  byteCount = Request.TotalBytes
  RequestBin = Request.BinaryRead(byteCount)
  Dim UploadRequest
  Set UploadRequest = CreateObject("Scripting.Dictionary")
  BuildUploadRequest  RequestBin
  contentType = UploadRequest.Item("blob").Item("ContentType")
  filepathname = UploadRequest.Item("blob").Item("FileName")
  filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
  value = UploadRequest.Item("blob").Item("Value")

  'Create FileSytemObject Component
  Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

  'Create and Write to a File
  pathEnd  = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
  pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
   Classe=Request.QueryString("Classe")
  'Classe=Session("Classe")
   'Id_Mod=Session("Id_Mod") dovo lu uso ?
  Id_Mod=Request.QueryString("Id_Mod")
  divid=Request.QueryString("divid")
  Id_Classe=Request.QueryString("Id_Classe")
  
  'Id_Par=Request.Form("txtId_Par")
  Id_Par=UploadRequest.Item("txtId_Par").Item("Value")
  url=Server.MapPath(homesito)&"/"&Classe&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
  url=Replace(url,"\","/")
      
   response.write("<br>Path info"& url &"/"& filename)
   ' response.write("<br>Path end1"& pathEnd1)
   ' response.write("<br>Test:"& left(pathEnd1,10))  
	          
				'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url2="C:\Inetpub\umanetroot\anno_2012-2013_ma\Upload.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.Close
%>			
 
 		  <form name="frm_upload"> 
		  File <input type="text" name="txtFile" style="border:none"><br>
          Caricato <input type="text" name="txtUP" value="3" size="1" style="border:none"> 
		 <!--<img src="black1.jpg" width="8" height="11" name="bar">-->
		 
		 <%
			 Set MyFile = ScriptObject.CreateTextFile(url &"/"& filename)
	  
	  
	 
			For i = 1 to LenB(value)
					MyFile.Write chr(AscB(MidB(value,i,1)))
					percent=round((i/LenB(value))*100)%>
					  
				   <script type="text/javascript">
					document.frm_upload.txtFile.value="Upload di <%=filename%>"
					document.frm_upload.txtUP.value="<%=percent%>%"
					/* Visualizzo il contatore del % di caricamento*/
					</script>
				  <% 
			 
			 Next
			 MyFile.Close%>
      </form>
		
		<p align="center"><font face="Verdana" size="2">
		  File "<b><%=filename%></b>" ricevuto con successo</font>
		<p align="center"><font face="Verdana" size="2"><a href="../admin/form_upload.asp">torna</a></font></p>
		</p>
  
	  <% 'urldb=Classe&"/Risorse/"&filename
	  urldb=filename
     QuerySQL="  UPDATE Paragrafi SET URL_L = '" & urldb & "',URL_O='" &urldb&"'" &_
    " WHERE Id_Paragrafo='" & Id_Par& "';"
    
       ' response.write(QuerySQL)
		   ConnessioneDB.Execute QuerySQL 
         
		 response.Redirect "modificamodulo.asp?Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&byUmanet="&byUmanet
  
 end if 
  %>
 
<html>
<head>
</head>
<body bgcolor="#FFCB8C">

<p align="center"><font face="Verdana" size="2">
  File "<b><%=filename%></b>" ricevuto con successo</font>
<p align="center"><font face="Verdana" size="2">
 <%if session("DB")=1 then %>
<a href="../../home.asp">Home Page</a></font></p>
<%else%>
<a href="../../home.asp">Home Page</a></font></p>
<%end if%>
</p>
</body>
</html>