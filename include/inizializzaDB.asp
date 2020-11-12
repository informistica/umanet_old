	<%
						Domanda="?"
						R1="?"
						R2="?"
						R3="?"
						R4="?"
						Chi="?"
						Cosa="?"
						Dove="?"
						Quando="?"
						Come="?"
						Perche="?"
						Quindi="?"
						RE=1
						Spiegazione="?"
						CodiceAllievo=id
						CodiceTest="6Classe"
						Modulo="6C" 
						DataTest="12/12/2112"  ' NB DOPO QUESTA DATA SE ESISTEREMO ANCORA DOVRO' METTERE UNA DATA SEMPRE MAGGIORE DELL'A/S in CORSO IN MODO DA FAR FUNZIONARE IL LEFT JOIN  NELLE QUERY PER LE CLASSIFICHE CON DATA VARIABILE
						' dovrò inserire aNCHE I CREDITI INIZIZIALI !
						Cartella="?"
						Voto=0
						In_Quiz=2
						
						 
				   Autista = "?"
				   Destinazione = "?"
				   Carburante = "?"
				   Luogo = "?" 
				   Strada = "?"
				   Strada_KO = "?"
				   Strada_OK = "?"
				   Cespugli = "?"
				   Lupo = "?"
				   Cestino = "?" 
				   Distanza = 1
				   Sintesi="?"
				 
				   Topolino =  "?"
				   Formaggio =  "?"
				   Fame = "?" 
				   Labirinto = "?"
				   Strada =  "?"
				   Strada_KO = "?"    
				   Strada_OK = "?"    
				   Testata =  "?"   
				   Distanza =  1 
				   Sintesi= "?"
				   
				   
				   SoggettoC =  "?"
				   DomandaC =  "?"
				   MotivazioneC = "?" 
				   DesiderioC = "?"
				   BisognoC="?"
				   SoggettoS =  "?"
				   RispostaS = "?"    
				   MotivazioneS = "?"    
				   DesiderioS =  "?" 
				   BisognoS="?"
				   TipoEvento = 1 
				   TolleranzaC = 3 
				   URL_teoria="?"
				   Cartella="?"
				 
				 
				 
				   
				
						QuerySQL="INSERT INTO Domande (Quesito, Risposta1, Risposta2,Risposta3,Risposta4,RispostaEsatta,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Domanda & "','" & R1 & "', '" & R2 & "','" & R3 & "','" & R4 & "','" & RE & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella &  "','"& In_Quiz &"';"
						
						
								'dim objFSO,objCreatedFile
				'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'				Dim sRead, sReadLine, sReadAll, objTextFile
				'				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_registrazione2.txt"
				'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'				'QuerySQL=cint(left(id_classe,1))  
				'				'QuerySQL="cicco"
				'				objCreatedFile.WriteLine(QuerySQL)
				'				objCreatedFile.Close
						'response.write(QuerySQL&"<br>")
						
						ConnessioneDB.Execute QuerySQL 
						
						
							QuerySQL="INSERT INTO Nodi (Chi, Cosa, Dove,Quando,Come,Perche,Quindi,Id_Stud,Id_Arg,Id_Mod,Data,Cartella) SELECT '" & Chi & "','" & Cosa & "', '" & Dove & "','" & Quando & "','" & Come & "','" & Perche & "','" & Quindi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Cartella & "';"
				 
				'response.write(QuerySQL&"<br>")
				   ConnessioneDB.Execute QuerySQL 
						   
				QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & Cartella & "','" & In_Quiz & "';"
				 
				'response.write(QuerySQL&"<br>")
				   ConnessioneDB.Execute QuerySQL 
				  
				  
							  'dim objFSO,objCreatedFile
				'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'				Dim sRead, sReadLine, sReadAll, objTextFile
				'				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_registrazione.txt"
				'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'				QuerySQL=cint(left(id_classe,1))  
				'				'QuerySQL="cicco"
				'				objCreatedFile.WriteLine(QuerySQL)
				'				objCreatedFile.Close
				 
				 ' per aggiungere il credito per l'iscrizione
				 querySQL="Select ID_Esercitazione  " &_
				 " from [dbo].[2ESERCITAZIONI_SINGOLI] " &_
				 " where Descrizione='Iscrizione' and Id_Classe='" &id_classe&"';"  
				  Set rsTabella = ConnessioneDB.Execute (QuerySQL)   
				  
				 ' response.write(QuerySQL&"<br>")
				  id_eser=rsTabella(0) 
				  rsTabella.close()
				   
				QuerySQL="INSERT INTO [dbo].[2CREDITI] (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & id_eser  & "','" & CodiceAllievo & "','" & 1 & "';"
				 
				   ConnessioneDB.Execute QuerySQL    
				'response.write(QuerySQL&"<br>")
				
				
				 QuerySQL="INSERT INTO M_Navigazione (Autista, Destinazione, Carburante,Luogo,Strada,Strada_OK,Strada_KO,Cespugli,Lupo,Cestino,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella) SELECT '" & Autista & "','" & Destinazione & "', '" & Carburante & "','" & Luogo & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Cespugli & "','" & Lupo & "','" & Cestino & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "';" 
				 
				 
				' response.write(QuerySQL&"<br>")
				ConnessioneDB.Execute QuerySQL 
				
				 QuerySQL="INSERT INTO M_Topolino (Topolino, Formaggio, Fame,Labirinto,Strada,Strada_OK,Strada_KO,Testata,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora) SELECT '" & Topolino & "','" & Formaggio & "', '" & Fame & "','" & Labirinto & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Testata & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & DataTest &"';" 
		 
				  'response.write(QuerySQL&"<br>")
						ConnessioneDB.Execute QuerySQL  
								 
 
 QuerySQL="INSERT INTO M_Desideri (SoggettoC, DomandaC, MotivazioneC,DesiderioC,BisognoC,SoggettoS,RispostaS,MotivazioneS,DesiderioS,BisognoS,TipoEvento,TolleranzaC,URL_Teoria,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora) SELECT '" & SoggettoC & "','" & DomandaC & "', '" & MotivazioneC & "','" & DesiderioC & "','" & BisognoC & "','" & SoggettoS & "','" & RispostaS & "','" & MotivazioneS & "','" & DesiderioS & "','" & BisognoS  & "'," & TipoEvento & "," & TolleranzaC &",'"  & URL_teoria &"','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) &"';" 
	'response.write("<br>"&QuerySQL)
   ConnessioneDB.Execute(QuerySQL)
						
					 
						
						
						
						
						
						 
				 
				 Messaggio="InizializzaDB"
	QuerySQL="INSERT INTO FORUM_MESSAGES (comments,CodiceAllievo,Id_Classe,Punti,DatePosted) SELECT '" &Messaggio & "','" & CodiceAllievo & "','" & Id_Classe & "',0,'"&DataTest &"';"
 '
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_ago.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
  ConnessioneDB.Execute QuerySQL 
   'response.write(QuerySQL&"<br>") 
  
   
  
   
   'Set ConnessioneDB = Nothing
   %>
   