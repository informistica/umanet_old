<%@ Language=VBScript %>


        <%


  


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>


        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->

    <%
	   
	   
	  domanda=Request("domanda")
      domanda=Replace(domanda,"'",chr(96))
      Modulo=Request("Modulo")
      Paragrafo=Request("Paragrafo")
      CodiceSottopar=Request("CodiceSottopar")
      Scadenza=Request("scadenza")
	  cartella=Request("cartella")
	  Img=Request("Img")
      
      
     
	  'testo=Request.QueryString("testo")
	  testo=Request("testo")
	  'testo=Request.form("editor1")
	 
      'testo=Replace(testo,"'",chr(96))
	  testo=Replace(testo,"editor1=","")
	  
	 	on error resume next


if CodiceSottopar<>"" then
	QuerySQL="Select count(*) from preFrasi where Id_Paragrafo='"&Paragrafo&"' and Id_Sottoparagrafo='"&CodiceSottopar&"';"
else
    QuerySQL="Select count(*) from preFrasi where Id_Paragrafo='"&Paragrafo&"';"
end if
'RESPONSE.WRITE("<br>"&querysql)
set rsTabella=ConnessioneDB.Execute (QuerySQL)
if rsTabella(0)>0 then
		if CodiceSottopar<>"" then
		  QuerySQL="Select max(Posizione) from preFrasi where Id_Paragrafo='"&Paragrafo&"'and Id_Sottoparagrafo='"&CodiceSottopar&"';"
		else
			QuerySQL="Select max(Posizione) from preFrasi where Id_Paragrafo='"&Paragrafo&"';"
		end if
    'RESPONSE.WRITE("<br>"&querysql)
	set rsTabella=ConnessioneDB.Execute (QuerySQL)
	contPos=rsTabella(0)
else
	 contPos=0
end if
contPos=contPos+1
'response.write(QuerySQL& "<br>" & "conPos="&contPos)
 


  Set fso = CreateObject("Scripting.FileSystemObject") 
  
	' creo la cartella per il modulo dentro la cartella risorse del corso  
 
        url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/"&Modulo&"_Esercizi" 
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			' response.Write( "<br>La cartella " & url & " esiste .<br>")
		else
		  '  response.Write( "<br>Creazione della cartella esercizi :" & url) 
	   	fso.CreateFolder (url)
		end if

'inserisco prefrase 
       
        cFile=0
		QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo,Estesa)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos & ",'" & Scadenza & "'," & Img & "," & cFile & ",'" & CodiceSottopar& "'," & 1 & ";"
	   ' RESPONSE.WRITE("<br>"&querysql)
        ConnessioneDB.Execute QuerySQL
        QuerySQL="Select max(Id_Prefrase) from preFrasi where Id_Paragrafo='"&Paragrafo&"';"
       ' RESPONSE.WRITE("<br>"&querysql)
	    set rsTabellaPos=ConnessioneDB.Execute (QuerySQL)		  
        ID=rsTabellaPos(0)

        url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" &cartella &"/" &Modulo&"_Esercizi/"&Paragrafo&"_"&ID&".txt"
		  
	  	url=Replace(url,"\","/")
	    if fso.FileExists (url) then
		'	 response.Write( "<br>Il file " & url & " esiste gi�.<br>")
             fso.DeleteFile (url)
		else
		 '   response.Write( "<br>Creazione file esercizio :" & url) 
			fso.CreateTextFile(url)
	
		end if
        Set objCreatedFile = fso.CreateTextFile(url, True)
	    objCreatedFile.WriteLine(testo)
        objCreatedFile.Close

		'	 QuerySQL="UPDATE FORUM_MESSAGES SET topic = '" & titolo &"', comments = '" & testo &"' WHERE ID="&id&";"
		'  ConnessioneDB.Execute(QuerySQL)
		'  RESPONSE.WRITE(querysql)
		  If Err.Number = 0 Then
			'Response.Write "Modifica avvenuta! "
				stato=1
				messaggio="Inserimento effettuato correttamente"
			Else
				stato=0
				messaggio=Err.Description
			Err.Number = 0
			End If

		''response.write(QuerySQL)
'
'messaggio=testo  ' **** da togliwere
		response.write(" { ")
		 response.write("""stato"": """&stato&""","  &_
		 """messaggio"": """&messaggio&"""")
		 response.write("}")
'


%>
