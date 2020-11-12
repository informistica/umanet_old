<%
'           Class to get data of a form with ENCTYPE="multipart/form-data"	
'Usage:
'			Set variable = New BinForm		
'Methods:							
'			Read()					
'				Reads the input of the multipart encoded form		
'				NOTE All other methods are only available after you	call this method
'			Save(Inputname)
'				Saves the file input in File field labeled Inputname
'				Inputname should be a string equal to the label of a formfield of type
'				"FILE".	NOTE: Only to be used after .Read
'				Example: .Save("UploadFile")
'			Form(Inputname)
'				Returns the value entered in the specified formfield. Use this method	'
'				the same as Request.Form. NOTE: Only to be used after .Read				'
'				Example: .Form("inputtext")												'
'Properties:																			'
'	Read / Write																		'
'			Extensions																	'
'				Array of strings with extensions allowed. Seperator included. If Only	'
'				"*" is specified all datatypes are allowed								'
'				Default: .txt, .gif, .jpg, .mp3, .wma									'
'				Example .Extensions = Array(".txt", ".gif", ".jpg", ".mp3", ".wma")
'			Directory	
'				String with the absolute path for download.NOTE: Inetuser should have	'
'				correct rights on folder.
'				Default: Server.Mappath(\upload)
'			Create
'				Boolean which states wether or not non exsisting folders should be		'
'				created.
'				Default: FALSE	
'			OverWrite
'				Boolean which states wether or not exsisting files should be
'				overwritten.
'				Default: FALSE	
'			MaxSize
'				Double with the maximum size in bytes of the files selected in the post '
'				form.
'				Default: 100,000 (97.7 k)
'			SaveAs
'				String which holds the name on how the file should be saved. Do not		'
'				include extension. SaveAs is reset to Empty after Save is called		'
'				Default: Empty (No rename)
'	Read-only
'			Success
'				Boolean stating wether or not the last Save, SaveAs or SaveToDb Method	'
'				was successfull. NOTE: Only contains valid data after one of the Methods'
'				is called
'			Version
'				String with software version of the class.
'			Log	
'				String with the result log. On success it holds the data of all
'				read-only Properties on failure it holds one or more error messages	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
Class BinForm
'Class Declarations (used througout this class)
Private aExtensions, sDirectory, bCreate, bOverWrite, bSuccess, sLog, dicUpload, bReadSuccess
Private bNoInput, nTotalBytes, sFilePathExt, sNewFile, nMaxSize, z, sSaveAs, sWhere, svnm
'Properties setting
	Public Property LET Extensions(byVal arrayInput)
		aExtensions = arrayInput
	End Property
	Public Property LET Directory(byVal strInput)
		sDirectory = strInput
	End Property
	Public Property LET Create(byVal boolInput)
		bCreate = boolInput
	End Property
	Public Property LET OverWrite(byVal boolInput)
		bOverWrite = boolInput
	End Property
	Public Property LET MaxSize(byVal lngInput)
		nMaxSize = lngInput
	End Property
	Public Property LET SaveAs(byVal strInput)
		sSaveAs = "" & strInput
	End Property
'Properties reading
	Public Property GET Extensions()
		
		'The following lines are commented by johnson so that all the files can be uploaded
		'If IsArray( aExtensions ) then
		'	Extensions = aExtensions
		'Else
			'If Len(arrExt) = 1 AND aExtensions = "*" Then
		'		Extensions = "*"
			'Else
				' default extensions
				'Extensions = Array(".txt", ".jpg", ".gif")
				Extensions = "*"
			'End If
		'End If
		
	End Property
	Public Property GET flName()
	  flName=svnm
	end Property
	Public Property GET Directory()
		Directory = sDirectory
	End Property
	Public Property GET Create()
		Create = CBool(bCreate)
	End Property
	Public Property GET OverWrite()
		OverWrite = CBool(bOverWrite)
	End Property
	Public Property GET MaxSize()
		If nMaxSize = 0 Then
'			nMaxSize = 100000
			nMaxSize = 10000000000000000000000000000000000000000000000
		End If	
		MaxSize = nMaxSize
	End Property
	Public Property GET SaveAs()
		SaveAs = sSaveAs
	End Property
	Public Property	GET Success()
		Success = bSuccess
	End Property
	Public Property	GET Version()
		Version = "Copyrights Bob Coppens 2001 Version 1.0."
	End Property
	Public Property	GET Log()
		Log = sLog
	End Property
'Events (automatically invoked)
	Private Sub Class_Initialize()
		Call ClearVariables()
		Call ClearDictionary()
	End Sub
	Private Sub Class_Terminate()
		Call ClearVariables()
		Call ClearDictionary()
	End Sub
'Methods to be called by programmer always call Read first
	Public Default Sub Read()
		' Checks for input (error 3), filesize (error 4) 
		' and extensions (error 5) and initializes the 
		' reading of all the binary data
		nTotalBytes = Request.TotalBytes
		If nTotalBytes <= 0 Then
			'Log the error of no data input
			Call BinForm_Log(3)
			bSuccess = False
		ElseIf nTotalBytes > 0 And nTotalBytes <= MaxSize Then
			'Goto Readbinary sub
			Call ReadBinary(nTotalBytes)
		Else
			'Log the error of too much data and stop
			Call BinForm_Log(4)
			bSuccess = False
		End If
	End Sub
	
Public Sub Save(byVal sFormInput)
 If bReadSuccess Then
			Dim nFileLength, sFilePathName, sSaveName, sContentType
			Dim spStr

			sContentType = dicUpload.Item(sFormInput).Item("ContentType")
			sFilePathName = dicUpload.Item(sFormInput).Item("FileName")

'			spStr=split(Replace(sFilePathName,"\","*",1),"*")
			spStr=split(sFilePathName,"\")

'			sSaveName = Right(sFilePathName,Len(sFilePathName)-InStrRev(sFilePathName,"\"))
			sSaveName = spStr(ubound(spStr))

			sFilePathExt = Right(sFilePathName, 4)
	
			nFileLength = LenB(dicUpload.Item(sFormInput).Item("Value")) 
			If IsArray(Extensions) Then
				'check the extension
				bSuccess = ExtensionCheck( sFilePathExt )
			Else
				If Extensions = "*" Then 
					bSuccess = True
				End If
			End If

			If not bSuccess Then
				'Illigal extension posted
				Call BinForm_Log(5)
			End If
			'Response.write "Directory="&Directory&"<br>"
			urldb=sSaveName ' solo il nome del file da inserire nel database per URL risorsa
			
			'If sSaveAs <> ""  Then
'				sSaveName = Directory & "\" & sSaveAs & Right(sSaveName, 4)
'				svnm=sSaveAs&right(sSaveName,4)  'added by miks
'			Else
'				sSaveName = Directory & "\" & "OK_"& sSaveName
'				svnm=sSaveName 'added by miks
'			End If

			'response.write("Directory="&Directory&"<br>")
			'response.write("UrlDb="&urldb&"<br>")
			'response.write("SAVE="&sSaveName&"<br>")
			' create the new file on the server
	
	
		dim dimensione
		dimensione=LenB(dicUpload.Item(sFormInput).Item("Value"))/1000
		dimensione=fix(dimensione)
	    if dimensione>200 then ' torno al chiamante
			   response.Redirect request.serverVariables("HTTP_REFERER") 
	    else ' eseguo inserimento
'response.write("Lunghezza file="& LenB(dicUpload.Item(sFormInput).Item("Value")	)	)			
'		  Set objFSO = CreateObject("Scripting.FileSystemObject")
'        				url1="C:\Inetpub\umanetroot\anno_2012-2013\logDim.txt"
'        				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'        				objCreatedFile.WriteLine("Dim="&dimensione& "Kb")
'        				objCreatedFile.Close		
'	
	
			if AggRisMod<>"" then
			
				QuerySQL="  UPDATE Moduli SET URL = '" & urldb & "',URL_OL='" &urldb&"'" &_
    " WHERE ID_Mod='" & Id_Mod& "';"
    		
				'response.write(QuerySQL&"<br>")
				ConnessioneDB.Execute QuerySQL 
				
				 'response.write("<br>" &"Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&Conta="&Conta)
				  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
  url=Replace(url,"\","/")
                '  response.write("<br>"& url&"/"&urldb)
			      sSaveName=url&"/"&urldb  
				  Call WriteFile( sSaveName, dicUpload.Item(sFormInput).Item("Value") )
				  sSaveAs = ""
				  response.Redirect "../inserisci_modulo1.asp?Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&inserito=1" 
				' response.write "../inserisci_modulo1.asp?Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&inserito=1" 
		     
			 end if
						
			 ' se sono stato chiamato da modificamodulo.asp per caricare le risorse dei paragrafi allora 
			 if AggRisPar<>"" then
			
			  	Id_Par = dicUpload.Item("txtId_Par").Item("Value")
'  
				 QuerySQL="  UPDATE Paragrafi SET URL_L = '" & urldb & "',URL_O='" &urldb&"'" &_
				" WHERE Id_Paragrafo='" & Id_Par& "';"
		
				'response.write(QuerySQL&"<br>")
				
				 ConnessioneDB.Execute QuerySQL 
				
				' response.write("<br>" &"Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&Conta="&Conta)
				  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
  url=Replace(url,"\","/")
                 ' response.write("<br>"& url&"/"&urldb)
				
			      sSaveName=url&"/"&urldb
				  
				  Call WriteFile( sSaveName, dicUpload.Item(sFormInput).Item("Value") )
				   sSaveAs = ""
				   
				   response.Redirect "../modificamodulo.asp?Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&Conta="&right(Id_Par,1)
		     end if
			 
			
			  
			 if AggRisFrase<>"" then
	 QuerySQL="Select count(*) from Frasi where Id_Prefrase=" &ID_Prefrase &" and Id_Stud='"&Session("CodiceAllievo")&"';" 
			  set rs=ConnessioneDB.Execute (QuerySQL )
			  num=rs(0)
			 ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logNUM.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(num)
'				objCreatedFile.Close
			  set rs=nothing
			 
			 
			 
			 
			 
			 
			    ' aggiungo la risorsa per la frase
				imgname = dicUpload.Item("imgname").Item("Value") ' nome logico dell'immagine
			  	
				CodiceAllievo=Session("CodiceAllievo")
				DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
                ' l'inserimento della frase lo faccio solo alla prima chiamata, o se sono chiamata da 2inserisci_valutazione_frase, cioè se by_UPLOAD="" 
				if num=0 then ' se non è già inserita la inserisco 
					if by_UPLOAD="" then
					Sintesi = dicUpload.Item("S1").Item("Value")
					 %>
					 <!--#include file="../../cFrasi/2inserisci_frase1_include.asp"--> 		 
					<%
					end if
				end if
				' response.write("<br>" &"Id_Mod="&Id_Mod&"&Classe="&Classe&"&Caricato=1&Id_Classe="&Id_Classe&"&divid="&divid&"&Conta="&Conta)
				  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Frasi/Img" 
  url=Replace(url,"\","/")
                 ' response.write("<br>File:"& url&"/"&urldb)
			      urldb=Paragrafo&"_"&ID&"_"&contDomande&"."&right(sSaveName,3)
				  sSaveName=url&"/"&urldb
					 'if num=0 then ' se non è già inserita
					   Call WriteFile( sSaveName, dicUpload.Item(sFormInput).Item("Value") )
					   
					   sSaveAs = ""
					 '  response.write(sSaveName)
					 ' inserisco il il link all'immagine
					   QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & urldb & "','" & imgname & "';"
					   
					   if ID<>"" and urldb<>"" and imgname<>"" then
					   ConnessioneDB.Execute QuerySQL 
					   else
					         
							response.Redirect "../../cFrasi/2compilaprefrase.asp?Cartella="&cartella&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&prefrase=1"
							 
					   end if
				  'end if
				  response.write(QuerySQL)
				    %><script language="javascript" type="text/javascript"> 
    window.alert("Inserimento avvenuto!");
     

</script>
				   
				   <%' torno al chimanate
				  		' if Request.ServerVariables("HTTP_REFERER") <>"" then 
'							response.Redirect request.serverVariables("HTTP_REFERER") 
'					 	end if 
				   
				   if AggImg<>"" then
					    if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
					 	end if 
				   else
				   ' devo fare il redirect ad 2inserisci_frase con un id per capire da dove provengo 
				   ' per inserire eventuali altre immagini
				    response.Redirect "../../cFrasi/2inserisci_frase.asp?daUpload=1&Quesito="&Quesito&"&ID_Prefrase="&ID_Prefrase&"&Capitolo="&Capitolo&"&Cartella="&Cartella&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&prefrase="&preFrase&"&by_UPLOAD=1&Img=1&Id_Frase="&ID&"&contDomande="&contDomande
			       end if
				   
				   
				   ' response.Redirect "../../2compilaprefrase.asp?Capitolo="&Capitolo&"&Cartella="&Cartella&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&Prefrase=1"
				   
				  
			  end if
			'********************************************
			' qua devo cambiare il nome del file 
			' userò variabile di sessione con codieallievo e codicefrase2
			
			
			 if AggRisDomanda<>"" then
			    ' aggiungo la risorsa per la frase
			
                ' l'inserimento della frase lo faccio solo alla prima chiamata, o se sono chiamata da 2inserisci_valutazione_frase, cioè se by_UPLOAD=""     
				 imgname = dicUpload.Item("imgname").Item("Value") ' nome logico dell'immagine
				if by_UPLOAD="" then
					
					 R1 =  dicUpload.Item("txtR1").Item("Value") 
					 R1 = Replace(R1, Chr(34), "'")
					 R1=  Replace(R1,"'",Chr(96))
					 R2 =  dicUpload.Item("txtR2").Item("Value") 
					 R2 = Replace(R2, Chr(34), "'")
					 R2=  Replace(R2,"'",Chr(96))
					 R3 =  dicUpload.Item("txtR3").Item("Value") 
					 R3 = Replace(R3, Chr(34), "'")
					 R3=  Replace(R3,"'",Chr(96))
					 R4 =  dicUpload.Item("txtR4").Item("Value") 
					 R4 = Replace(R4, Chr(34), "'")
					 R4=  Replace(R4,"'",Chr(96))
					 RE =  dicUpload.Item("txtRE").Item("Value") 
					 Spiegazione = dicUpload.Item("S1").Item("Value")
					 Domanda = dicUpload.Item("txtDomanda").Item("Value") 
					 Domanda =  Replace(Domanda, Chr(34), "'")
					 Domanda =  Replace(Domanda, "'",Chr(96))
					 Titolo=Domanda
					 
					CodiceAllievo=Session("CodiceAllievo")
					DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
					 Sintesi = dicUpload.Item("S1").Item("Value")
				 %>
                 <!--#include file="../../cDomande/inserisci_test1_include2.asp"--> 		 
				<%
				 
				end if
							   
				    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Domande/Img" 
  url=Replace(url,"\","/")
                 ' response.write("<br>File:"& url&"/"&urldb)
			      urldb=Paragrafo&"_"&ID&"_"&contDomande&"."&right(sSaveName,3)
				  sSaveName=url&"/"&urldb
		
				  
				   Call WriteFile( sSaveName, dicUpload.Item(sFormInput).Item("Value") )
				   sSaveAs = ""
				
				 ' inserisco il il link all'immagine
				   QuerySQL="INSERT INTO Domande_Img (Id_Domanda,Url,Nome) SELECT " & ID & ",'" & urldb & "','" & imgname & "';"
				  '  Set objFSO = CreateObject("Scripting.FileSystemObject")
'					url1="C:\Inetpub\umanetroot\anno_2012-2013\logDomande360.txt"
'					Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'					objCreatedFile.WriteLine(querySQL)
'					objCreatedFile.Close	
				   
				   
				   ConnessioneDB.Execute QuerySQL 
				  
 		   
				    ' torno al chimanate
				  		' if Request.ServerVariables("HTTP_REFERER") <>"" then 
'							response.Redirect request.serverVariables("HTTP_REFERER") 
'					 	end if 
				   
				   if AggImg<>"" then
					    if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
					 	end if 
				   else
				   ' devo fare il redirect ad 2inserisci_frase con un id per capire da dove provengo 
				   ' per inserire eventuali altre immagini
				    response.Redirect "../../cDomande/inserisci_test.asp?Quesito="&Domanda&"&Capitolo="&Capitolo&"&Cartella="&Cartella&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&by_UPLOAD=1&Tipo=1&Img=1&Id_Domanda="&ID&"&contDomande="&contDomande&"&Multiple="&Multiple
			       end if
				   
				  
			 end if
			
			
	    end if ' if dimensione>200k
 'end if
End Sub
	Public Function Form(sFormInput)
		If not bReadSuccess Then
			Form = ""
		Else
			Form = dicUpload.Item(sFormInput).Item("Value")
		End If
	End Function
'Routines (for internal class use)
	Private Sub ClearVariables()
		aExtensions = Array(".txt", ".jpg", ".gif", ".wma", ".mp3")
		sDirectory = Server.MapPath("/upload")
		bReadSuccess = False
		sSaveAs = ""
		bCreate = False
		bOverWrite = False
		'nMaxSize = 100000
		bSuccess = True
		sLog = ""
		bNoInput = True
	End Sub
	Private Sub ClearDictionary()
		 ' clear dictionary
		Set dicUpload = Server.CreateObject("Scripting.Dictionary")
		dicUpload.RemoveAll
		Set dicUpload = Nothing
	End Sub
	' **********************
	Private Sub ReadBinary(byVal MaxBytes)
		Dim nByteCount, sAllBinary
		nByteCount = MaxBytes
		sAllBinary = Request.BinaryRead (nByteCount)
		Set dicUpload = Server.CreateObject("Scripting.Dictionary")
		On Error Resume Next
			BuildUpload sAllBinary
			If not bSuccess Then
				Call BinForm_Log(9)
			End If
		On Error Goto 0
	End Sub
	'************************
	Private Function ExtensionCheck(byVal Ext)
		Dim Item
		For each Item in Extensions
			If LCase( Ext ) = LCase( Item ) Then 
				ExtensionCheck = True
				Exit Function
			End if
		Next
		ExtensionCheck = False
	End Function
	Private Sub BinForm_Log(nMessage)
		'First line in Logs with any error
		Select Case nMessage
			Case 0
				sLog = "The posted data was succesfully read." &vbCrLf
			Case 1
				sLog = sLog & "File " & sNewFile & " was saved succesfully." &vbCrLf
			Case 2
				sLog = sLog & "File " & sNewFile & " was succesfully uploaded to database." &vbCrLf
			Case 3
				sLog = "The posted data was not read." &vbCrLf
				sLog = sLog & "No data has been posted." &vbCrLf
			Case 4
				sLog = "The posted data was not read." &vbCrLf
				sLog = sLog & "The total bytes of the post (" & CStr(nTotalBytes) & " bytes) is larger than set maximum of " & CStr(nMaxSize) & "." &vbCrLf
			Case 5
				sLog = sLog & "Upload failed. Illegal Extension posted." &vbCrLf
			Case 6
				sLog = sLog & "Upload failed. " & Directory & " directory not found. Unable to create directory. Check the rights of Internet User." &vbCrLf
			Case 7
				sLog = sLog & "Upload failed. " & Directory & " directory not found. Set Create to True if directory should be created." &vbCrLf
			Case 8
				sLog = sLog & "Upload failed. File already exsists and OverWrite is set to False" &vbCrLf
			Case 9
				bSuccess = False
				sLog = sLog & "Upload failed. Error parsing file failed. Probably due to filesize" &vbCrLf
				Exit Sub
		End Select
	End Sub
	Private Sub BuildUpload(RequestBin)
		dim PosBeg,PosEnd,boundary,boundaryPos,Pos,Name,PosFile
		dim PosBound,FileName,ContentType,Value
		bSuccess = False
		PosBeg = 1
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
		boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos = InstrB(1,RequestBin,boundary)
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
			Dim UploadControl
			Set UploadControl = Server.CreateObject("Scripting.Dictionary")
			Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
			Pos = InstrB(Pos,RequestBin,getByteString("name="))
			PosBeg = Pos+6
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
			PosBound = InstrB(PosEnd,RequestBin,boundary)
			If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg = PosFile + 10
				PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg = Pos+14
				PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg = PosEnd+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
				Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg = Pos+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			End If
			UploadControl.Add "Value" , Value	
			dicUpload.Add name, UploadControl	
			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
		Loop
		bSuccess = True
		bReadSuccess = True
		
		' rinomino il file caricato
	
		
		Call BinForm_Log(0)
		Exit Sub
	End Sub	
	
	Private Sub WriteFile(byVal NAME, byVal CONTENTS)
		 ' create the file on the server
		dim soWriteFile, sTmpDir
		Set soWriteFile = Server.CreateObject("Scripting.FileSystemObject")
		If Not soWriteFile.FolderExists( Directory ) Then
			'If create new directories is allowed
			If bCreate Then
				' start from root and work your way down
				sTmpDir = Left( Directory, 3)
				Do While sTmpDir <> Directory
					' if arrived at final subdirectory
					If InStr( Mid( Directory, ( Len( sTmpDir ) + 2 ), (Len(Directory) - ( Len( sTmpDir ) + 2 ))), "\" ) = 0 Then
						sTmpDir = Directory
					Else
						'volume to subdirectory below
						sTmpDir = Left( Directory,( Len( sTmpDir ) + InStr(Mid( Directory, ( Len( sTmpDir ) + 2 ), (Len(Directory) - ( Len( sTmpDir ) + 2 ))), "\")) )
					End If
					If Not soWriteFile.FolderExists( sTmpDir ) Then
						'create subdirectory with internal error creation
						On Error Resume Next
						soWriteFile.CreateFolder( sTmpDir )
						If Not soWriteFile.FolderExists( sTmpDir ) Then
							' error creating directory (probably rights on server)
							Call BinForm_Log(6)
							bSuccess = False
						End If
						On Error Goto 0
					End If
				Loop
			Else
				' error creating directory since no rights in Class
				Call BinForm_Log(7)
				bSuccess = False
			End If
		End If
		If NOT OverWrite AND soWriteFile.FileExists( NAME ) Then
			 ' don't allow file overwrite
			Call BinForm_Log(8)
			bSuccess = False
		End If
		If bSuccess Then
		
	
			Set sNewFile = soWriteFile.CreateTextFile( NAME )
			For z = 1 to LenB( CONTENTS )
				 ' translate binary data into ASCII characters and write them into the file.
				sNewFile.Write chr( AscB( MidB( CONTENTS, z, 1) ) )
			Next
			 ' clean up and inform the user of successful upload
			sNewFile.Close
			Set sNewFile = Nothing
			sNewFile = NAME
			Call BinForm_Log(1)
		End If
		Set soWriteFile = Nothing
	End Sub
	Private Function getByteString(StringStr)
		dim char, i
		For i = 1 to Len(StringStr)
			char = Mid(StringStr,i,1)
			getByteString = getByteString & chrB(AscB(char))
		Next
	End Function
	Private Function getString(StringBin)
		dim intCount
		getString =""
		For intCount = 1 to LenB(StringBin)
			getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
		Next
	End Function
End Class
%>
