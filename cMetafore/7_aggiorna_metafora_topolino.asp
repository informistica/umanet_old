<%@ Language=VBScript %>
 

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim Topolino,Formaggio,Fame,Labirinto,Strada,Strada_KO,Strada_OK,Testata,Distanza
   
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
  
   Cartella=Request.QueryString("Cartella")
  
    
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            
	CodiceMetafora=Request.QueryString("CodiceMetafora")
	Paragrafo=Request.QueryString("Paragrafo")
	Modulo=Request.QueryString("Modulo")
	'DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
 
	   Topolino = ucase(Request.querystring("txtTopolino"))
	   Topolino = Replace(Topolino, Chr(34), "'")
	   Topolino=  Replace(Topolino,"'",Chr(96))
  
   Formaggio = ucase(Request.querystring("txtFormaggio"))
   Formaggio = Replace(Formaggio, Chr(34), "'")
   Formaggio=  Replace(Formaggio,"'",Chr(96))


   Fame = ucase(Request.querystring("txtFame"))
   Fame = Replace(Fame, Chr(34), "'")
   Fame=  Replace(Fame,"'",Chr(96))

   Labirinto = ucase(Request.querystring("txtLabirinto"))
   Labirinto = Replace(Labirinto, Chr(34), "'")
   Labirinto=  Replace(Labirinto,"'",Chr(96))
   Strada = ucase(Request.querystring("txtStrada"))
   Strada = Replace(Strada, Chr(34), "'")
   Strada=  Replace(Strada,"'",Chr(96))

   Strada_KO = ucase(Request.querystring("txtStrada_ko"))
   Strada_KO = Replace(Strada_KO, Chr(34), "'")
   Strada_KO=  Replace(Strada_KO,"'",Chr(96))
   
   
   
   Strada_OK = ucase(Request.querystring("txtStrada_ok"))
   Strada_OK = Replace(Strada_OK, Chr(34), "'")
   Strada_OK=  Replace(Strada_OK,"'",Chr(96))
   
   Testata = ucase(Request.querystring("txtTestata"))
   Testata = Replace(Testata, Chr(34), "'")
   Testata=  Replace(Testata,"'",Chr(96))
   
   
   Distanza = ucase(Request.querystring("txtDistanza"))
   Distanza = Replace(Distanza, Chr(34), "'")
   Distanza=  Replace(Distanza,"'",Chr(96))
   
   Sintesi=ucase(Request.querystring("S1"))
   Sintesi= Replace(Sintesi, Chr(34), "'")
   Sintesi=  Replace(Sintesi,"'",Chr(96))
    Sintesi=  Replace(Sintesi,Chr(39),Chr(96))
 
   
   
   
   if ( (len(Topolino)=0) or (len(Formaggio)=0) or (len(Fame)=0) or (len(Labirinto)=0) or (len(Strada)=0) or (len(Strada_KO)=0) or(len(Strada_OK)=0) or(len(Testata)=0) or(len(Distanza)=0)) then
    errore=2
  
   end if
   
 if (errore=0) then
   
   ' devo vedere se il setting è tale da richiedere voto=1 come default oppure no  
    QuerySQL1="Select * from Setting"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
	Valutato=rsTabella.fields("Valutato") 
	rsTabella.close
	if Valutato=1 then
        Voto=1 ' valore di default 
	else
	    Voto=0
	end if
	
if daSimulazione<>"" then ' aggiorno 
 
  
 QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK &  "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata &"', Distanza= '" & Distanza & "'  WHERE CodiceMetafora =" &CodiceMetafora&";"
 response.write(QuerySQL)
	response.write("OK")		   
			   
  ConnessioneDB.Execute QuerySQL 
 'response.Redirect "6_simula_metafora_topolino.asp?Cartella="& Session("Cartella")&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&CodiceMetafora&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo&"&Modulo="&Modulo&"&nocache="&rand
				
 else	
  QuerySQL="INSERT INTO M_Topolino (Topolino, Formaggio, Fame,Labirinto,Strada,Strada_OK,Strada_KO,Testata,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ThreadParent) SELECT '" & Topolino & "','" & Formaggio & "', '" & Fame & "','" & Labirinto & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Testata & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4)& "',"& ThreadParent &";" 
 '  end if 
  ConnessioneDB.Execute QuerySQL 
  
    QuerySQL = "SELECT CodiceMetafora,Cartella FROM M_Topolino WHERE CodiceMetafora=(Select Max(CodiceMetafora) FROM M_Topolino);" 
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    ID=rsTabella(0)
  
  
     Session("CodiceMetafora")=ID
	 Session("Capitolo")=Capitolo
	 Session("Paragrafo")=Paragrafo
	 Session("CodiceTest")=CodiceTest

	 
    CARTA=rsTabella(1)
	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" 'per il server on line
     
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DELLA METAFORA
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
	 
	url=Replace(url,"\","/")
	  
		
					'url2="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logTopo.txt"
	'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
	'				objCreatedFile.WriteLine(url)
	'				objCreatedFile.Close
	
	'response.write(url)
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
	objCreatedFile.WriteLine(Sintesi)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
	'response.write(url)
  
  
  response.write("OK")
  
  
 end if 
   

 
 