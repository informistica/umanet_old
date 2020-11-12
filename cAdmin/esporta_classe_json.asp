
 <!-- #include file = "../cUtenti/adovbs.inc" -->
<%
Response.charset="utf-8"  
'Response.charset="iso-8859-1"
Call Response.AddHeader("Access-Control-Allow-Origin", "*")   
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
      <%
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
dim oRs,oRs1, oCmd,oCmd1, sSQL, sAns
dim oParam

function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sInput,"è","&egrave;")

  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")

ReplaceCar = sAns
 
end function


	umanet=request("txtUmanet")
 	id_mod=request("txtIDMOD") 'non si usa più
  id_classe=request("txtIDCLASSE")
	pk_corso=CInt(request("txtPKcorso"))
	pk_modulo=CInt(request("txtPKmodulo"))
	order_modulo=CInt(request("txtORDERmodulo"))
	pk_start_paragrafo=CInt(request("txtPKparagrafo"))
	pk_start_sottoparagrafo=CInt(request("txtPKsottoparagrafo"))
	pk_start_prefrase=CInt(request("txtPKprefrase"))


 	
Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
'nome_modulo="sistemi_operativi"
url=Server.MapPath(homesito & "/esportati")&"/classe"&id_classe&".json"
 url=Replace(url,"\","/")
'response.write(url)
Set objCreatedFile = objFSO.CreateTextFile(url, True)
QuerySQL="select * from MODULI_CLASSE where Id_Classe='"&id_classe&"' and Visibile=1"
    ' QuerySQL="select * from Moduli where ID_Mod='"&id_mod&"'"
 '    response.write(QuerySQL)
	Set rsTabellaMod = ConnessioneDB.Execute(QuerySQL)
  riga="["
   i=0
          k=0
          j=0
          m=0 'contatore moduli
  do while not rsTabellaMod.eof
            nome_modulo=replace(rsTabellaMod("Titolo")," ","_")
          'response.write("<br><br>Creato il file:"&url) 
            id_mod=rsTabellaMod("ID_Mod")
            folderorigine=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" &Session("cartella")&"/"&id_mod&"_Esercizi"
          file=""

          riga=riga&"{""model"":""courses.module"", ""pk"":"&pk_modulo+m&","
          response.write(riga)
          objCreatedFile.WriteLine(riga)
          'file=file&riga

          riga="""fields"": {""course"":"&pk_corso&",""title"":"""&rsTabellaMod("Titolo")&""",""description"":"""", ""order"":"&order_modulo&"}},"
          response.write(riga)
          objCreatedFile.WriteLine(riga)
          'file=file&riga



          QuerySQL="SELECT [ID_Mod],[Titolo],[ID_Paragrafo],[Tit],[URL_O],[posMod],[posPar],[Visibile] " &_
          " FROM MODULI_PARAGRAFI_CLASSE1 " &_
          " WHERE [ID_Mod]='" & id_mod&"' Order by posPar ;"

            
          Set rsTabellaPar = ConnessioneDB.Execute(QuerySQL)
          ' per ogni paragrafo, scrivo 

         
          do while not rsTabellaPar.eof

          pk_paragrafo=pk_start_paragrafo+i
          riga="{""model"":""courses.content"", ""pk"":"&pk_paragrafo&","
          'file=file&riga
          response.write(riga)
          objCreatedFile.WriteLine(riga)

          riga="""fields"": {""module"":"&pk_modulo+m&",""title"":"""&rsTabellaPar("Tit")&""",""url"":"""&rsTabellaPar("URL_O")&""", ""order"":"&i&"}},"
          response.write(riga)
          'file=file&riga
          objCreatedFile.WriteLine(riga)
                                
          qsl="SELECT * FROM  ParagrafiSottoparagrafi2 where Id_Paragrafo='"&rsTabellaPar("Id_Paragrafo")&"' order by Posizione"
          set rsTabSottoPar= ConnessioneDB.execute(qsl)

          hasSott=not rsTabSottoPar.eof
          deadline="2020-06-30 00:00:00"
           'deadline=""

              if hasSott then


                  do while not rsTabSottoPar.eof 
                  position=1
                  pk_sottoparagrafo=pk_start_sottoparagrafo+j
                  riga="{""model"":""courses.subcontent"", ""pk"":"&pk_sottoparagrafo&","
                  response.write(riga)
                  objCreatedFile.WriteLine(riga)
                  'file=file&riga
                  riga="""fields"": {""content"":"&pk_paragrafo&",""title"":"""&rsTabSottoPar("Titolo")&""",""url"":"""&rsTabSottoPar("URL")&""", ""order"":"&j&"}},"
                  response.write(riga)
                  objCreatedFile.WriteLine(riga)
                  'file=file&riga
                  QuerySQL="SELECT * " &_
                  "FROM preFrasi WHERE preFrasi.Id_Paragrafo='" & rsTabellaPar("Id_Paragrafo") & "' and Id_Sottoparagrafo='" &rsTabSottoPar("Id_Sottoparagrafo")  & "' order by Posizione;"
                  Set rsTabellaPrefrasi = ConnessioneDB.Execute(QuerySQL)	
                  
                      do while not rsTabellaPrefrasi.eof
                      
                      pk_prefrase=pk_start_prefrase+k
                      riga="{""model"":""courses.prephrase"", ""pk"":"&pk_prefrase&","
                      response.write(riga)
                      objCreatedFile.WriteLine(riga)
                    ' file=file&riga
                     ' response.write("140:" & rsTabellaPrefrasi("Estesa")=True)
                     ' response.write("<br>"&rsTabellaPrefrasi("Estesa"))
                      if rsTabellaPrefrasi("Estesa")=True then
                      extended="True"
                      'carico testo question dal file
                          
                    fileorigine=folderorigine&"/"&rsTabellaPar("Id_Paragrafo")&"_"&rsTabellaPrefrasi("ID_Prefrase")&".txt"
                    fileorigine=Replace(fileorigine,"\","/")
                              if objFSO.FileExists(fileorigine) then
                      Set objTextFile = objFSO.OpenTextFile(fileorigine, ForReading) 
                      question = Server.HTMLEncode(Replace(objTextFile.ReadAll, Chr(34), "'"))    
                      'sReadAll=url
                      objTextFile.Close
                    else
                      question=""
                    end if
                      else
                          extended="False"
                          question=""
                      end if
                      hasimg="False"
                        if rsTabellaPrefrasi("Img")=1 then
                      hasimg="True"
                      end if

                      riga="""fields"": {""title"":"""& replacecar(rsTabellaPrefrasi("Quesito"))&""",""module"":"&pk_modulo&",""content"":"&pk_paragrafo&",""subcontent"":"& pk_sottoparagrafo&", ""position"":"& position &",""deadline"":"""&deadline &""",""extended"":"""& extended &""",""img"":"""& hasimg &""",""question"":"""& question &"""}},"
                      response.write(riga)
                      objCreatedFile.WriteLine(riga)
                      'file=file&riga

                      k=k+1
                      position=position+1
                      rsTabellaPrefrasi.movenext
                      loop
                      j=j+1
                      rsTabSottoPar.movenext
                      loop
                else  

                  QuerySQL="SELECT * " &_
                  "FROM preFrasi WHERE preFrasi.Id_Paragrafo='" & rsTabellaPar("Id_Paragrafo")  & "' order by Posizione;"		 
                          pk_sottoparagrafo="null"
                          Set rsTabellaPrefrasi = ConnessioneDB.Execute(QuerySQL)	
                    
                      position=1
                      do while not rsTabellaPrefrasi.eof
                      pk_prefrase=pk_start_prefrase+k
                      riga="{""model"":""courses.prephrase"", ""pk"":"&pk_prefrase&","
                      response.write(riga)
                      objCreatedFile.WriteLine(riga)
                    ' file=file&riga
                      
                      if rsTabellaPrefrasi("Estesa")=True then
                              extended="True"
                      'carico testo question dal file
                    fileorigine=folderorigine&"/"&rsTabellaPar("Id_Paragrafo")&"_"&rsTabellaPrefrasi("ID_Prefrase")&".txt"
                    fileorigine=Replace(fileorigine,"\","/")
                              if objFSO.FileExists(fileorigine) then
                      Set objTextFile = objFSO.OpenTextFile(fileorigine, ForReading) 
                      question = Server.HTMLEncode(Replace(objTextFile.ReadAll, Chr(34), "'"))
                      'sReadAll=url
                      objTextFile.Close
                    else
                      question=""
                    end if
                      else
                          extended="False"
                          question=""
                      end if
                       hasimg="False"
                        if rsTabellaPrefrasi("Img")=1 then
                      hasimg="True"
                      end if
                      riga="""fields"": {""title"":"""& replacecar(rsTabellaPrefrasi("Quesito")) &""",""module"":"&pk_modulo&",""content"":"&pk_paragrafo&",""subcontent"":"& pk_sottoparagrafo&", ""position"":"& position &",""deadline"":"""&deadline &""",""extended"":"""& extended &""",""img"":"""& hasimg &""",""question"":"""& question &"""}},"
                      response.write(riga)
                      objCreatedFile.WriteLine(riga)
                    ' file=file&riga
                      k=k+1
                      position=position+1
                      rsTabellaPrefrasi.movenext
                      loop
                
                end if
              i=i+1 
            rsTabellaPar.movenext
              loop 
              riga=""
	 	m=m+1
    rsTabellaMod.movenext
    loop


     'riga=left(riga,len(riga)-1)
     'riga=riga&"]"
     '******************
     ' dal testo prodotto nel browser, togliere manualmente l'ultimo carattere , e mettere ]
            ' response.write(riga)
        ' objCreatedFile.WriteLine(riga)
        'file=file&riga 
        'objCreatedFile.WriteLine(file)

objCreatedFile.Close


%>

 








 