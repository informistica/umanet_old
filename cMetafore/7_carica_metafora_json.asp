<%@ Language=VBScript %>
 
     
        <% 
		
 
 function ReplaceCar(sInput)
 dim sAns
 sAns=sInput
   sAns = Replace(sAns, "  ", " ") 'sostituizione doppio spazio con uno singolo
   sAns = Replace(sAns, "	", " ") 'sostituzione spazi per evitare errori
   sAns = Replace(sAns, " ?", "?") ' rimozione spazio prima del punto di domanda
   sAns = Replace(sAns, "’", "'") ' sostituzione di un'apice con quello classico
   sAns = Replace(sAns, "…", "...") 'sostituzione tre puntini
   sAns = Replace(sAns, Chr(25), "'") 'sostituizione apice
   sAns = Replace(sAns, VBCrLf, "") 'sostituizione ritorno a capo  
  sAns = Replace(sAns,chr(96),chr(39)) ' sostituizione finale dell'apice storto con il classico apice

 ReplaceCar = sAns
 end function
 
   
  						
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
	 
        	  
          
         
 
    
    <%
	  CodiceMetafora = Request.QueryString("CodiceMetafora")
	  tipoMetafora=Request.QueryString("tipoMetafora") ' 0 Topolino, 1 Navigazione, 2 Sdesideri
	  daSimulazione=Request.QueryString("daSimulazione")
  'CodiceTest = Request.QueryString("CodiceTest")
  
  
 if daSimulazione="" then ' devo caricare le spiegazioni 
 
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")


end if
  
  
  
  
 ' 
	if strcomp(tipoMetafora,"0")= 0 then
'		'response.write("topolino")
		  QuerySQL="SELECT Tit, ID_Paragrafo, Cognome, CodiceMetafora, ID_Mod, Topolino, Formaggio, Fame, Labirinto, Strada, Strada_OK, Strada_KO, Testata, Distanza, In_Quiz,Posizione,Titolo, Posizione,Pi,Pf,Cartella " &_
" From Elenco_Metafore_topolino " &_
" Where CodiceMetafora =" & CodiceMetafora & "" 
'response.write(QuerySQL)
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 Pi=rsTabella("Pi") ' codice della metafora precedente
 Pf=rsTabella("Pf") ' ' codice della metafora seguente 
'

			if daSimulazione="" then
				url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & rsTabella("Cartella") &"/" &rsTabella("ID_Mod")&"_Metafore/"&rsTabella("ID_Mod")&"_"&rsTabella("Tit")&"_"&CodiceMetafora&".txt"
				url=Replace(url,"\","/")
				'response.write(url)
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				sReadAll = ReplaceCar(rtrim(objTextFile.ReadAll))
				'response.write(sReadAll)
				objTextFile.Close
			'
			end if
			
			response.write(" { ")  
			 response.write("""soggetto"&i&""": """&ReplaceCar(rsTabella("Topolino"))&""","  &_
			  """obiettivo"": """&ReplaceCar(rsTabella("Formaggio"))&""","  &_
			 """motivazione"": """&ReplaceCar(rsTabella("Fame"))&""","  &_
			 """ambiente"": """&ReplaceCar(rsTabella("Labirinto"))&""","  &_
			 """comportamento"": """&ReplaceCar(rsTabella("Strada"))&""","  &_
			 """ok"": """&ReplaceCar(rsTabella("Strada_OK"))&""","  &_
			 """ko"": """&ReplaceCar(rsTabella("Strada_KO"))&""","  &_
			 """testata"": """&ReplaceCar(rsTabella("Testata"))&""","  &_
			 """pi"": """&ReplaceCar(rsTabella("Pi"))&""","  &_
			  """pf"": """&ReplaceCar(rsTabella("Pf"))&""","  &_
			  """codicemetafora"": """&ReplaceCar(rsTabella("CodiceMetafora"))&""","  &_
			 """testata"": """&ReplaceCar(rsTabella("Testata"))&""","  &_
			  """distanza"&i&""": """&ReplaceCar(rsTabella("Distanza"))&""","  &_
			 """spiegazione"": """&sReadAll&"""")
			 response.write("}")
' 
	else 
		if strcomp(tipoMetafora,"1")= 0 then
		'response.write("navigazione")
			QuerySQL="SELECT Tit, ID_Mod,ID_Paragrafo,Cognome, CodiceMetafora, ID_Mod, Autista, Destinazione, Carburante, Luogo, Strada, Strada_OK, Strada_KO, Cespugli, Lupo,Cestino,Distanza, In_Quiz,Posizione,Cartella,Pi,Pf " &_
			" From Elenco_Metafore_Navigazione" &_
			" Where CodiceMetafora =" & CodiceMetafora & "" 
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			if daSimulazione="" then
				url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & rsTabella("Cartella") &"/" &rsTabella("ID_Mod")&"_Metafore/"&rsTabella("ID_Mod")&"_"&rsTabella("Tit")&"_"&CodiceMetafora&".txt"
				url=Replace(url,"\","/")
				'response.write(url)
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				sReadAll = ReplaceCar(rtrim(objTextFile.ReadAll))
				'response.write(sReadAll)
				objTextFile.Close
			'
			end if
		response.write(" { ") 
		 response.write("""soggetto"&i&""": """&ReplaceCar(rsTabella("Autista"))&""","  &_
		 """obiettivo"": """&ReplaceCar(rsTabella("Destinazione"))&""","  &_
		 """motivazione"": """&ReplaceCar(rsTabella("Carburante"))&""","  &_
		 """ambiente"": """&ReplaceCar(rsTabella("Luogo"))&""","  &_
		 """comportamento"": """&ReplaceCar(rsTabella("Strada"))&""","  &_
		 """ok"": """&ReplaceCar(rsTabella("Strada_OK"))&""","  &_
		 """ko"": """&ReplaceCar(rsTabella("Strada_KO"))&""","  &_
		 """feedback"": """&ReplaceCar(rsTabella("Cespugli"))&""","  &_
		 """eccessi"": """&ReplaceCar(rsTabella("Cestino"))&""","  &_
		 """conseguenze"": """&ReplaceCar(rsTabella("Lupo"))&""","  &_
		 """pi"": """&ReplaceCar(rsTabella("Pi"))&""","  &_
		 """pf"": """&ReplaceCar(rsTabella("Pf"))&""","  &_
		 """codicemetafora"": """&ReplaceCar(rsTabella("CodiceMetafora"))&""","  &_
		 """distanza"&i&""": """&ReplaceCar(rsTabella("Distanza"))&""","  &_
		 """spiegazione"": """&sReadAll&"""")
		 response.write("}")
'
		else
'			 QuerySQL="SELECT * " &_
'" From Elenco_Metafore_Desideri " &_
'" Where CodiceMetafora =" & CodiceMetafora & "" 
'			response.write("desideri")
		end if
  end if 
'	 
'	 
'                 
                 
               
  

 ' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logSimulazione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close  
'response.write(QuerySQL)
	'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	 'Pi=rsTabella("Pi") ' codice della metafora precedente
	 'Pf=rsTabella("Pf") ' ' codice della metafora seguente 
 
			 ' rsTabella.close : Set rsTabella = Nothing  %>
				   
                      
                      
                      
                      
                      
               
