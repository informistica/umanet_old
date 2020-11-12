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
 
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
<!-- #include file = "../service/controllo_sessione.asp" --> 
<!-- #include file = "../var_globali.inc" --> 

<% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella,rsTabella1,rsTabella2, QuerySQL,StringaConnessione,URL,RecSet,k
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
  
	Dim Drive, OggFile, origine, destinazione
	Set fso = CreateObject("Scripting.FileSystemObject")
   

'on error resume next


   'Url=Session("UrlDB")
   Url1=Request.QueryString("Url")
   ID_Mod_old=Request.QueryString("ID_ModSel")
   homesito_origine=Request.QueryString("homesito_origine")
  
 
ID_Mod_new=Request.Form("txtID_Mod") 
' per stabilire se trasferire o condividere modulo
condividi=request.QueryString("condividi")
 
Id_Classe=Session("Id_Classe")
Id_ClasseOld=Session("ID_ClaSel")
divid=Session("divid")
Classe=request.QueryString("Cartella") 
byUmanet=request.QueryString("byUmanet") ' vale 1 se sono chiamato per trasferire un modulo umanet
   
   'Apertura della connessione al database  
    
	session("DB2")=1
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     <% 'connessione al db sorgente da cui importare
	
	'  Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") 
 ' lo commento ma quando lavorer� su due db servira ripristinare 
	'  ConnessioneDB2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
       '       "DBQ=" & Server.MapPath(Url1)  
		   
' stringa di connessione al db sorgente, sessione impostata in seleziona_origine.asp



 
 if (left(pathEnd1,10)="c:\inetpub") then
     locale=1
  else
     locale=0
  end if 
  
  
  if condividi<>"" then
  ' inseriisci on Classi_Moduli_paragrafi
  ID_Mod_new=ID_Mod_old
 ' per condividere aggiungo solo questi record e creo le cartele vuote
			 QuerySQL=" SELECT * from Classi_Moduli_Paragrafi where Id_Classe='"&Id_ClasseOld&"' and Id_Modulo='"&ID_Mod_new&"';"
			  response.write(QuerySQL)
			 set rsTabellaCMP=ConnessioneDB.Execute (QuerySQL)
			   while not rsTabellaCMP.eof
				   QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (Id_Classe,Id_Modulo,Id_Paragrafo,Id_As)  SELECT '" & Id_Classe & "','" & rsTabellaCMP("Id_Modulo") & "','" & rsTabellaCMP("Id_Paragrafo") & "'," & 1 &";"
				   response.write(QuerySQL)
		          ConnessioneDB.Execute QuerySQL 
				 rsTabellaCMP.movenext
				 
			    wend
  
  else ' trasferisco
  
  ' per stabilire la posizione del nuovo modulo, che per ora va in coda a tutti maxPos+1
      
	 if byUmanet<>"" then ' seleziono moduli umanet
	    QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&Session("Cartella")&"';"
	else
	  QuerySQL="SELECT max(posizione) FROM Moduli where Cartella='"&Session("Cartella")&"';"
	
     end if
	
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	 
	  if isnull(rsTabella1(0)) then
	     maxPos=0
	  else	 
		  maxPos=rsTabella1(0)
	  end if 
	  posizione=maxPos+1
	 'posizione=rsTabella1("Posizione")
	  
	   QuerySQL="SELECT * FROM INPORT;"
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  Inport=rsTabella1(0) ' se vale 0 inporto da vecchai struttura altrimenti nuovi url
	  
	  
	  QuerySQL="SELECT * FROM Moduli WHERE ID_Mod = '" & ID_Mod_old & "' order by Posizione;"
	 ' set rsTabella1=ConnessioneDB0.Execute(QuerySQL)  'da ripristinare quando importero in due db
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  
	  'Gestire anche la posizione in cui deve essere inserito il modulo, in coda o in mezzo a scelta ?
	 ' posizione_old=rsTabella1("Posizione")
	posizione_old=right(ID_Mod_old,len(ID_Mod_old)-instr(ID_Mod_old,"_"))
	 QuerySQL="SELECT * FROM CARTELLA_MODULO_PARAGRAFI WHERE ID_Mod = '" & ID_Mod_old & "' order by Posizione;"
	 
	'  Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logTxMod0.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				
	 'set rsTabella=ConnessioneDB0.Execute(QuerySQL)  
	 set rsTabella=ConnessioneDB.Execute(QuerySQL)  
	 cartella=rsTabella("Cartella")
	 
	 
	 ' devo inserire il nuovo modulo
	 
	 
	if byUmanet<>"" then
	' vedo la prima posizione del _
	' poi seleziono dal secondo in poi per ricavare la posizione del modulo
	primo=instr(Id_Mod_new,"_")
	
	 QuerySQL="  INSERT INTO Moduli (Id_Mod,Titolo,Posizione,Cartella,URL,URL_OL,Visibile)  SELECT '" & ID_Mod_new & "','" & rsTabella1("Titolo") & "', '" & right(Id_Mod_new,len(Id_Mod_new)-instr(primo+1,Id_Mod_new,"_"))& "', '" & Classe  & "', '" & rsTabella1("URL") & "', '" & rsTabella1("URL_OL") & "',1;"		
 
    else
	 QuerySQL="  INSERT INTO Moduli (Id_Mod,Titolo,Posizione,Cartella,URL,URL_OL,Visibile)  SELECT '" & ID_Mod_new & "','" & rsTabella1("Titolo") & "', '" & right(Id_Mod_new,len(Id_Mod_new)-instr(Id_Mod_new,"_"))& "', '" & Classe  & "', '" & rsTabella1("URL") & "', '" & rsTabella1("URL_OL") & "',1;"
	end if  
	 
	  ConnessioneDB.Execute QuerySQL
 response.write("<br>"&QuerySQL)
	  i=1
	    
	    do while not rsTabella.eof 
		   
		  Id_ParagrafoNew=ID_Mod_new&"_"&i
		 ' Id_ParagrafoOld=ID_Mod_old&"_"&i' ecco l'inghippo che sfasava le domande tra i paragrafi, quando l'id non corrisponde alla posizione si sballava l'ordinamento
		  Id_ParagrafoOld=rsTabella("ID_Paragrafo") 		  
		 
		   ' inserisco i nuovi paragrafi
		   QuerySQL="  INSERT INTO Paragrafi (Id_Paragrafo,Titolo,Posizione,URL_L,URL_O)  SELECT '" & Id_ParagrafoNew & "','" & rsTabella("Titolo") & "', '" & rsTabella("Posizione") & "', '" & rsTabella("URL_L") & "', '" & rsTabella("URL_O") & "';"
          
		  ConnessioneDB.Execute QuerySQL 
		  
		 
		  ' verifico se esistono sotto paragrafi e nel caso trasferisco anche quello
		  '
		      QuerySQL=" SELECT * from ParagrafiSottoparagrafi where Id_Paragrafo='"& Id_ParagrafoOld&"';"
			 set rsTabellaSotPar=ConnessioneDB.Execute (QuerySQL)
			'set rsTabellaSotPar=ConnessioneDB0.Execute (QuerySQL)
			 j=1
			 k=1
			 while not rsTabellaSotPar.eof
			      'per ogni sottoparagrafo devo creare 
				
				'   Id_SottoParagrafoOld=Id_ParagrafoOld&"_"&j
				 Id_SottoParagrafoOld=rsTabellaSotPar("Id_Sottoparagrafo")
				    ' inserisco l'associazione tra  paragrafi e sottoparagrafi
		    
			' leggo il sottoparagrafo e lo inserisco
			 QuerySQL=" SELECT * from Sottoparagrafi where Id_Sottoparagrafo='"&Id_SottoParagrafoOld&"';"
			 set rsTabellaSotPar1=ConnessioneDB.Execute (QuerySQL)
			
			   while not rsTabellaSotPar1.eof
			        Id_SottoParagrafoNew=Id_ParagrafoNew&"_"&j
				   QuerySQL="  INSERT INTO Sottoparagrafi (Id_Sottoparagrafo,Titolo,Posizione,Url)  SELECT '" & Id_SottoParagrafoNew & "','" & rsTabellaSotPar1("Titolo") & "','" & rsTabellaSotPar1("Posizione") & "','" & rsTabellaSotPar1("Url") &"';"
		  		   ConnessioneDB.Execute QuerySQL 
				   response.write("<br>"&QuerySQL)
				 rsTabellaSotPar1.movenext
				 j=j+1
			    wend
				   QuerySQL="  INSERT INTO ParagrafiSottoparagrafi (Id_Paragrafo, Id_Sottoparagrafo)  SELECT '" & Id_ParagrafoNew  & "','" & Id_SottoParagrafoNew & "';"
		   		 ConnessioneDB.Execute QuerySQL 
				    response.write("<br>"&QuerySQL)
				 rsTabellaSotPar.movenext
				 k=k+1
			 wend
		  
		  
		   ' inserisco l'associazione tra classe moduli paragrafi
		    QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod_New & "', '" & Id_ParagrafoNew & "';"
		    ConnessioneDB.Execute QuerySQL 
			' trasferisco i compiti del  paragrafo del modulo
			 response.write("<br>"&QuerySQL)
			
			
			QuerySQL="Select * from preFrasi where Id_Mod='"&ID_Mod_old & "' and Id_Paragrafo='"& Id_ParagrafoOld&"';"
			set rsTabella2=ConnessioneDB.Execute(QuerySQL)  
			  response.write("<br>"&QuerySQL)
			  Set objFSO = CreateObject("Scripting.FileSystemObject")
			  folderdestinazione=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/"&ID_Mod_new&"_Esercizi" 
			  folderdestinazione=Replace(folderdestinazione,"\","/")
			  if objFSO.FolderExists (folderdestinazione) then
				 response.Write( "<br>La cartella " & folderdestinazione & " esiste gi�.<br>")
			  else
		  		fso.CreateFolder (folderdestinazione) 
				response.Write( "<br>La cartella " &folderdestinazione & " � stata creata.<br>") 
		  	end if
			  do while not rsTabella2.eof 
			      
				 scadenza=rsTabella2("Scadenza")
				  'response.Write("<br>"&scadenza)
				 if isnull(scadenza) then
				    scadenza=fine_anno
				 end if
				 img=rsTabella2("Img")
				  if isnull(img) then
				    img=0
				 end if 
				 estesa=0
				 if rsTabella2("Estesa")=true  then
				    estesa=1
			     end if 
				 Id_Prefrase=rsTabella2("Id_prefrase")

			     QuerySQL="  INSERT INTO preFrasi (Id_Mod,Id_Paragrafo,Quesito,Posizione,Scadenza,Img,Id_Sottoparagrafo,Estesa)  SELECT '" & Id_Mod_New & "','" & Id_ParagrafoNew & "', '" & rsTabella2("Quesito") & "', '" & rsTabella2("Posizione") & "', '" & scadenza & "', '" & img& "', '" &replace(rsTabella2("Id_Sottoparagrafo"),Id_ParagrafoOld,Id_ParagrafoNew) & "','"&estesa&"';"     
		 '**** 
		'
'					url="C:\Inetpub\umanetroot\anno_2012-2013\logTx_"&i&".txt"
'					Set objCreatedFile = objFSO.CreateTextFile(url, True)
'					objCreatedFile.WriteLine(QuerySQL)
'					objCreatedFile.Close
			'   objCreatedFile.WriteLine(QuerySQL)
				 response.Write("<br>"&QuerySQL)
			  ConnessioneDB.Execute QuerySQL 
				 if estesa=1 then 'copio file esercizi
					QuerySql="select max(Id_prefrase) from preFrasi"
					set rsTab=ConnessioneDB.execute(QuerySql)
					MaxIdPrefrase=rsTab(0)
					folderorigine=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" &cartella&"/"&ID_Mod_old&"_Esercizi"	
					fileorigine=folderorigine&"/"&Id_ParagrafoOld&"_"&Id_Prefrase&".txt"
					fileorigine=Replace(fileorigine,"\","/")
					filedestinazione=folderdestinazione&"/"&Id_ParagrafoNew&"_"&MaxIdPrefrase&".txt"
					filedestinazione=Replace(filedestinazione,"\","/")
					response.write("<br>fileorigine="&fileorigine)
					response.write("<br>filedestinazione"&filedestinazione)
					objFSO.CopyFile fileorigine,filedestinazione
					
				 end if


				 rsTabella2.movenext()
			  loop ' fine trasferimento prefrasi
		   	  Set objFSO = nothing


		   QuerySQL="Select * from preNodi where Id_Mod='"&ID_Mod_old & "' and Id_Paragrafo='"& Id_ParagrafoOld&"';"
			 		set rsTabella2=ConnessioneDB.Execute(QuerySQL) 
				
			  do while not rsTabella2.eof 
			      
				 scadenza=rsTabella2("Scadenza")
				'  response.Write("<br>"&scadenza)
				 if isnull(scadenza) then
				    scadenza=fine_anno
				 end if
				 img=rsTabella2("Img")
				  if isnull(img) then
				    img=0
				 end if 
			     QuerySQL="  INSERT INTO preNodi (Id_Mod,Id_Paragrafo,Quesito,Posizione,Scadenza,Img)  SELECT '" & Id_Mod_New & "','" & Id_ParagrafoNew & "', '" & rsTabella2("Quesito") & "', '" & rsTabella2("Posizione") & "', '" & scadenza & "', '" & img & "';"     		
				' response.Write("<br>"&QuerySQL)
				 ConnessioneDB.Execute QuerySQL  
				 rsTabella2.movenext()
			  loop ' fine trasferimento preNodi
			  
			    QuerySQL="Select * from preDomande where Id_Mod='"&ID_Mod_old & "' and Id_Paragrafo='"& Id_ParagrafoOld&"';"
			 		 'set rsTabella2=ConnessioneDB2.Execute(QuerySQL) ' attenzione era commentato e quello sotto non c'era
				'set rsTabella2=ConnessioneDB0.Execute(QuerySQL) 
			  do while not rsTabella2.eof 
			      
				 scadenza=rsTabella2("Scadenza")
		
				 if isnull(scadenza) then
				    scadenza=fine_anno
				 end if
				 img=rsTabella2("Img")
				  if isnull(img) then
				    img=0
				 end if 
			     QuerySQL="  INSERT INTO preDomande (Id_Mod,Id_Paragrafo,Quesito,Posizione,Scadenza,Img)  SELECT '" & Id_Mod_New & "','" & Id_ParagrafoNew & "', '" & rsTabella2("Quesito") & "', '" & rsTabella2("Posizione") & "', '" & scadenza & "', '" & img & "';"     		
				' response.Write("<br>"&QuerySQL)
				 ConnessioneDB.Execute QuerySQL  
				 rsTabella2.movenext()
			  loop ' fine trasferimento preDomande
		   
		   
		   
		   
		   i=i+1
		   rsTabella.movenext()
		   loop
		   
		   
		
		'objCreatedFile.Close %>
	   
<%   

 end if ' if condividi<>"" then
 
 
' Adesso devo copiare la cartella delle risorse
 Dim  folderorigine, folderdestinazione
   if Inport=0 then
		folderorigine=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & cartella &"/Risorse/Mod_"&posizione_old	
		  
		
	else
	    folderorigine=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/materia_"&Session("idxMat")&"/"&cartella&"/Risorse/Mod_"&posizione_old	
	end if
	folderorigine=Replace(folderorigine,"\","/")
	folderdestinazione=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/Risorse/Mod_"&right(Id_Mod_new,len(Id_Mod_new)-instr(Id_Mod_new,"_"))
	'folderdestinazione=Server.MapPath(homesito)& "/Db1/Materie/"&Session("ID_Materia")&"/"&Classe&"/Risorse/Mod_"&right(Id_Mod_new,len(Id_Mod_new)-instr(Id_Mod_new,"_"))
	folderdestinazione=Replace(folderdestinazione,"\","/")
	' creo la cartella Verifiche vuota
	folderdestinazione1=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/Verifiche/Mod_"&right(Id_Mod_new,len(Id_Mod_new)-instr(Id_Mod_new,"_"))
	folderdestinazione1=Replace(folderdestinazione1,"\","/")
		if fso.FolderExists (folderdestinazione1) then
		   
		else
			fso.CreateFolder (folderdestinazione1)
		end if
 
  Dim cartelle(6)
	   cartelle(0)="_Domande"
	   cartelle(1)="_Frasi"
	   cartelle(2)="_Nodi"
	   cartelle(3)="_Spiegazioni"
	   cartelle(4)="_Metafore"
	   cartelle(5)="_Esercizi"
 
 
	
	'on error resume next
if not(byUmanet<>"") then ' la cartelle delle risorse la copio solo per i moduli not umanet
	'response.write "<br>devo copiare la cartella delle risorse " & folderorigine & " <br> � stata copiata in " & folderdestinazione & "."
	if condividi="" then ' copio se non sto condividendo
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		response.write("<br>"&folderorigine&"<br>"&folderdestinazione)   
		set folder = fso.GetFolder (folderorigine)
		
		if fso.FolderExists (folderorigine) then
		   folder.Copy folderdestinazione,true
		   response.write "<br>La cartella delle risorse " & folderorigine & " <br> � stata copiata in " & folderdestinazione & "."
		else
			response.write "Nessuna cartella risorse da trasferire"
		end if
	end if
end if

 

	' creo sia in trasferimento che in condivisione le cartelle per il modulo domande,frasi,nodi, il parametro ID_Mod_new � settato in cima alla pagina, nel caso di condivisdione � uguale a quello vecchio
	 Set fso = CreateObject("Scripting.FileSystemObject") 
		for i=0 to 5 
			url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/"&ID_Mod_new&cartelle(i) 
			url=Replace(url,"\","/")
			if fso.FolderExists (url) then
				 response.Write( "La cartella " & url & " esiste gi�.<br>")
			else
				fso.CreateFolder (url) 
				fso.CreateFolder (replace(url&"/Img","\","/"))
				response.Write( "La cartella " & url&"/Img" & " � stata creata.<br>") 
			end if
		next 
		if condividi<>"" then
			response.Write( "<br><hr><br>Modulo condiviso correttamente! ") 
		else
		



 


	response.Write( "<br><hr><br>Modulo importato correttamente! ") 
	 end if
 

 
 %>
	</font>   

		  
		<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
 
 
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	