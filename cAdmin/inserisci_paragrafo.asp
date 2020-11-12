 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
 
</head>
<body>
  

<%@ Language=VBScript %>  
<% Response.Buffer=True 
   Dim ConnessioneDB, rsTabella, QuerySQL
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   Id_Mod=Request.QueryString("Id_Mod")  
   Id_Classe=Request.QueryString("Id_Classe")
   Classe=Request.QueryString("Classe")
   Paragrafo=Request.form("txtParagrafoNuovo")
   Url=Request.form("txtUrlParagrafoNuovo")
   SotParagrafo=Request.form("txtSotParagrafoNuovo")
   SotUrl=Request.form("txtSotUrlParagrafoNuovo")
    
    
   %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <%  
   
   if Paragrafo<>"" then
      QuerySQL=" select max(Posizione) from PARAGRAFI_POSIZIONE where Id_Modulo='" &Id_Mod &"' ;"
	  set rsTabella=connessioneDB.execute(QuerySQL)
	  Posizione=rsTabella(0)+1
	  Id_ParagrafoNew=ID_Mod&"_"&Posizione	
		
		 
		   ' inserisco i nuovi paragrafi
		   QuerySQL="  INSERT INTO Paragrafi (Id_Paragrafo,Titolo,Posizione,URL_L,URL_O)  SELECT '" & Id_ParagrafoNew & "','" & Paragrafo & "', '" & Posizione & "', '" & Url & "', '" & Url & "';"
          ConnessioneDB.Execute QuerySQL 
		 ' response.Write(QuerySQL)
		   ' inserisco l'associazione tra classe moduli paragrafi
		    QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod & "', '" & Id_ParagrafoNew & "';"
		    ConnessioneDB.Execute QuerySQL 
			'response.Write(QuerySQL)
   else

		'response.write("SotParagrafo="&SotParagrafo)
	  if SotParagrafo<>"" then  
	   
	'*  Id_SotParagrafoNew=Request("txtIdSotParNew")
	'*  SotPosPar=Request("txtSotPosPar")

		'strText="3Ct$5_U_4_2_12"
		strText=Id_SotParagrafoNew
		arrLines = Split(strText,"_")
		Id_Sot=""
		For i=0 to ubound(arrLines)-1 
		  Id_Sot=Id_Sot&arrLines(i)&"_"
		next 
	'*	Id_Paragrafo=left(Id_Sot, len(Id_Sot)-1)
		 
		' inserisco il nuovo sottoparagrafo
		   QuerySQL="  INSERT INTO Sottoparagrafi (ID_Sottoparagrafo,Titolo,Posizione,URL)  SELECT '" & Id_SotParagrafoNew & "','" & SotParagrafo & "', '" & SotPosPar & "', '" & SotUrl & "';"
      '*    ConnessioneDB.Execute QuerySQL 
		 ' response.Write(QuerySQL)
		   ' inserisco l'associazione tra classe paragrafo e sottoparagrafo 
		    QuerySQL="  INSERT INTO ParagrafiSottoparagrafi (Id_Paragrafo, Id_Sottoparagrafo)  SELECT '" & Id_Paragrafo  & "','" & Id_SotParagrafoNew & "';"
		'*    ConnessioneDB.Execute QuerySQL 
			'response.Write(QuerySQL)


	  else
	  'aggiorno la posizione dei paragrafi
		'  QuerySql="SELECT  Moduli.ID_Mod,Moduli.Titolo, Paragrafi.ID_Paragrafo, Paragrafi.Titolo as [Tit],URL_O,URL_OL, Moduli.Posizione as [posMod],Paragrafi.Posizione as [posPar] " &_
		'" FROM Paragrafi, Moduli, Classi_Moduli_Paragrafi " &_
		'" WHERE  Classi_Moduli_Paragrafi.Id_Modulo=Moduli.ID_Mod and Classi_Moduli_Paragrafi.Id_Paragrafo=Paragrafi.ID_Paragrafo " &_
		'" And Moduli.ID_Mod='" & Id_Mod&"' order by Moduli.Posizione, Paragrafi.Posizione ;"

			QuerySQL="SELECT [ID_Mod],[Titolo],[ID_Paragrafo],[Tit],[URL_O],[URL_OL],[posMod],[posPar] " &_
			" FROM MODULI_PARAGRAFI_CLASSE1 " &_
			" WHERE [ID_Mod]='" & Id_Mod&"' Order by posPar ;"

			
			'response.write(QuerySql & " " &Paragrafo)
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			i=1
			do while not rsTabella.eof
			QuerySQL="  UPDATE Paragrafi SET  Posizione = " & Request.form("txtPosPar"&i)  & ", URL_O='" & Request.form("txtURL"&i)  & "' , URLDOC='" & Request.form("txtURLDOC"&i)  & "' WHERE ID_Paragrafo='" & Request.form("txtIdPar"&i) & "';"
					response.Write(i&"-)" &QuerySQL&"<BR>")
					ConnessioneDB.Execute QuerySQL 
			
			QuerySQL="SELECT  * from ParagrafiSottoparagrafi2  where Id_Paragrafo='"&rsTabella("ID_Paragrafo") &"' order by Posizione;"
			
			Set rsTabellaSottopar = ConnessioneDB.Execute(QuerySQL) 	  
				if not rsTabellaSottopar.eof then %>
				<% 
					
				j=1

				do while not rsTabellaSottopar.eof
					
					QuerySQL="  UPDATE Sottoparagrafi SET  Posizione = " & Request.form("txtSotPosPar"&i&j)  & ", URL='" & Request.form("txtSotURL"&i&j)  & "', URLDOC='" & Request.form("txtSotURLDOC"&i&j)  & "' WHERE ID_Sottoparagrafo='" & rsTabellaSottopar("ID_Sottoparagrafo")& "';"
					response.Write(i&j&")" &QuerySQL&"<BR>")
					ConnessioneDB.Execute QuerySQL 
				
					j=j+1
					rsTabellaSottopar.movenext
					loop
				%>
				<%end if%>
				<%  
				if Request.form("txtSotParagrafoNuovo"&i&j)<>"" then
				      
					QuerySQL="  INSERT INTO Sottoparagrafi (ID_Sottoparagrafo,Titolo,Posizione,URL)  SELECT '" & Request.form("txtIdSotParNew"&i&j) & "','" & Request.form("txtSotParagrafoNuovo"&i&j) & "', '" & Request.form("txtSotPosPar"&i&j) & "', '" &  Request.form("txtSotUrlParagrafoNuovo"&i&j) & "';"
					response.Write(i&j&")" &QuerySQL&"<BR>")
					ConnessioneDB.Execute QuerySQL 
				 QuerySQL="  INSERT INTO ParagrafiSottoparagrafi (Id_Paragrafo, Id_Sottoparagrafo)  SELECT '" & rsTabella("ID_Paragrafo")  & "','" & Request.form("txtIdSotParNew"&i&j) & "';"
					response.Write(i&j&")" &QuerySQL&"<BR>")
					ConnessioneDB.Execute QuerySQL 
					
				
				end if
				i=i+1
				rsTabella.movenext
			loop 
		end if
   
   
   
   
   end if
    
	On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento dell'avviso avvenuto! "
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If
     
	  if Request.ServerVariables("HTTP_REFERER") <>"" then 
		'response.Redirect request.serverVariables("HTTP_REFERER") 
	  end if 

   %>
	 


 
 
    
 
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../cClasse/home_app.asp?id_classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Apprendimento... </a></h3> 
	
  
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	