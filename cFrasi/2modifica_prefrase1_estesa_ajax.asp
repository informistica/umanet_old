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
	  cartella=Request("cartella")
      Id_Prefrase=Request("Id_Prefrase")
     ' Scadenza=Request("scadenza")
      
      
     
 
	  testo=Request("testo")
	  'testo=Request.form("editor1")
	 
      'testo=Replace(testo,"'",chr(96))
	  testo=Replace(testo,"editor1=","")
	  
	 	on error resume next


  

'inserisco prefrase 
        img=0
        cFile=0
		QuerySQL=" UPDATE preFrasi set Quesito= '" &  domanda & "' where Id_Prefrase='"&Id_Prefrase&"';"
	   ' RESPONSE.WRITE("<br>"&querysql)
        ConnessioneDB.Execute QuerySQL
       
'aggiorno file 

set fso=Server.CreateObject("Scripting.FileSystemObject")



        url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" &cartella&"/" &Modulo&"_Esercizi/"&Paragrafo&"_"&Id_Prefrase&".txt"
	  	url=Replace(url,"\","/")
	    if fso.FileExists (url) then
			' response.Write( "<br>Il file " & url & " esiste gi�.<br>")
             fso.DeleteFile (url)
		end if
		   '   response.write(url)
        Set objCreatedFile = fso.CreateTextFile(url, True)
	    objCreatedFile.WriteLine(testo)
        objCreatedFile.Close
 
		  If Err.Number = 0 Then
				stato=1
				messaggio="Aggiornamento effettuato"
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
