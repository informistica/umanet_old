 <%@ Language=VBScript %>
 <%
  Response.Buffer = true
 ' On Error Resume Next  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../service/controllo_sessione.asp" -->
<%    
  		 
		  cartella=Request.QueryString("cartella")
		  dasviluppa=Request.QueryString("dasviluppa")
		  CodiceAllievo=Request.QueryString("CodiceAllievo")
		  
		  CodiceTest=Request.QueryString("Codice_Test")
		  Modulo=Request.QueryString("Modulo")
		  Paragrafo=Request.QueryString("Paragrafo")
		  ID_Premetafora=Request.QueryString("ID_Premetafora") 
	 	  errore=0
		  voto=1

   
 
 Select case CodiceTest
   Case Cartella&"_U_2_3" 		
  		   ID=Request.QueryString("CodiceMetafora")
		   Topolino = Request("txtTopolino")
		   
		   Topolino = Replace(Topolino, Chr(34), "'")
		 
		   Topolino=Replace(Topolino,"'","''")
		   Formaggio = Request("txtFormaggio")
		   Formaggio = Replace(Formaggio, Chr(34), "'")
		   Formaggio=Replace(Formaggio,"'","''")
		    
		   Fame = Request("txtFame")
		   Fame = Replace(Fame, Chr(34), "'")
		   Fame=Replace(Fame,"'","''")
		   Labirinto = Request("txtLabirinto")
		   Labirinto = Replace(Labirinto, Chr(34), "'")
		   Labirinto=Replace(Labirinto,"'","''")
		   Strada = Request("txtStrada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request("txtStrada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request("txtStrada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Testata = Request("txtTestata")
		   Testata = Replace(Testata, Chr(34), "'")
		   Testata=Replace(Testata,"'","''")
		  
		   Distanza=Request("txtDistanza")
		   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		   
		if (instr(Topolino,"E")<>0) then ' virgola per il prulare di più soggetti Soggetto1 E Soggetto2
			plurale=1
		else
		 	plurale=0
		end if
		if (instr(Topolino,",")<>0) then ' virgola per il prulare di più soggetti  Soggetto1,Soggetto2
			plurale1=1
		else
		 	plurale1=0
		end if

		 narrazione = ""
		if ((plurale = 0) and  (plurale1 = 0)) then
			volere = "vuoi"
			raggiungere = "raggiungerai"
			avere = "hai"
			scegliere = "scegli"
			avvicinarsi = "ti avvicina"
			allontanarsi = "ti allontana"
			allontanarsi1 = "ti sei allontanato troppo hai"
			scontrarsi = "e ti sei scontrato"
			continuare = "continua"
			fare = "ci sei quasi fai"
	
		else  
			volere = "volete"
			raggiungere = "raggiungerete"
			avere = "avete"
			scegliere = "scegliete"
			avvicinarsi = "vi avvicina"
			allontanarsi = "vi allontana"
			allontanarsi1 = "vi siete allontanati troppo avete"
			scontrarsi = "e vi siete scontrati"
			continuare = "continuate"
			fare = "ci siete quasi fate"
		end if
		contSi = 0
		contNo = 0
		Motivato = 0

		narrazione = Topolino & " " & volere & " raggiungere " & Formaggio & " ?"
		narrazione = narrazione & " NO!  Mancando " & Fame & " per raggiungere " & Formaggio & " , " & Topolino & " nel contesto " & Labirinto & " non " & raggiungere & " l'obiettivo ! "
		narrazione = narrazione & " " & Topolino & " " & volere & " raggiungere " & Formaggio & " ?" & "Si!"
		narrazione = narrazione & "   " & Topolino & "  quale  " & Strada & " " & scegliere & " ?  "

		narrazione = narrazione & "ATTENZIONE  " & Topolino & "  la scelta  " & Strada_KO & " " & allontanarsi & " da  " & Formaggio&"."
		narrazione = narrazione & " :-(  " & Topolino & " " & allontanarsi1 & " scelto la strada chiusa  " & Strada_KO & " " & scontrarsi & " con " & Testata & "  "
		narrazione = narrazione & "  " & Topolino & "  quale  " & Strada & " " & scegliere & " ?  "
		narrazione = narrazione & "  " & Topolino & "  la scelta  " & Strada_OK & " " & avvicinarsi & " a  " & Formaggio & "  " & continuare & " così !  "
		narrazione = narrazione & " Coraggio " & fare & " l'ultimo passo ! '"
		narrazione = narrazione & " :-) COMPLIMENTI  " & Topolino & " " & avere & " raggiunto " & Formaggio & "!!!"
		    
	 
	
  '  if ((len(Topolino)=0) or (len(Formaggio)=0) or (len(Fame)=0) or (len(Labirinto)=0) or (len(Strada)=0) or (len(Strada_OK)=0) or(len(Strada_KO)=0) or(len(Distanza)=0) or(len(Testata)=0)) then
'  
'   errore=2
'   response.write(errore&"")
'   response.write(len(Topolino)&"-" &len(Formaggio)&"-" &len(Fame)&"-" &len(Labirinto)&"-" &len(Strada)&"-" &len(Strada_OK)&"-" &len(Strada_KO)&"-" &len(Distanza)&"-" &len(Testata))
'   
'  end if 
  
   
   
	 Case Cartella&"_U_2_5"  	 
	
           ID=Request.QueryString("CodiceMetafora")
		   Autista = Request("txtAutista")
		   Autista = Replace(Autista, Chr(34), "'")
		   Autista=Replace(Autista,"'","''")
		   Destinazione = Request("txtDestinazione")
		   Destinazione = Replace(Destinazione, Chr(34), "'")
		   Destinazione=Replace(Destinazione,"'","''")
		   Carburante = Request("txtCarburante")
		   Carburante = Replace(Carburante, Chr(34), "'")
		   Carburante=Replace(Carburante,"'","''")
		   Luogo = Request("txtLuogo")
		   Luogo = Replace(Luogo, Chr(34), "'")
		   Luogo=Replace(Luogo,"'","''")
		   Strada = Request("txtStrada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request("txtStrada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request("txtStrada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Cespugli = Request("txtCespugli")
		   Cespugli = Replace(Cespugli, Chr(34), "'")
		   Cespugli=Replace(Cespugli,"'","''")
		   Lupo = Request("txtLupo")
		   Lupo = Replace(Lupo, Chr(34), "'")
		   Lupo=Replace(Lupo,"'","''")
		   Cestino = Request("txtCestino")
		   Cestino = Replace(Cestino, Chr(34), "'")
		   Cestino=Replace(Cestino,"'","''")
		   Distanza=Request("txtDistanza")
		   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		   

		if (instr(Autista,"E")<>0) then ' virgola per il prulare di più soggetti Soggetto1 E Soggetto2
			plurale=1
		else
		 	plurale=0
		end if
		if (instr(Topolino,",")<>0) then ' virgola per il prulare di più soggetti  Soggetto1,Soggetto2
			plurale1=1
		else
		 	plurale1=0
		end if

		 narrazione = ""
		if ((plurale = 0) and  (plurale1 = 0)) then
			 volere = "vuoi"
            raggiungere = "raggiungerai"
            avere = "hai"
            scegliere = "scegli"
            avvicinarsi = "ti avvicina"
            avvicinarsi1 = "avvicinarti"
            avvicinarsi2 = "avvicinarsi"
            allontanarsi = "ti allontana"
            allontanarsi1 = "ti sei allontanato troppo hai"
            scontrarsi = "e ti sei scontrato"
            continuare = "continua"
            fare = "ci sei quasi fai"
            dovere = "devi"
            ti_vi = "ti"
	
		else  
		 volere = "volete"
            raggiungere = "raggiungerete"
            avere = "avete"
            scegliere = "scegliete"
            avvicinarsi = "vi avvicina"
            avvicinarsi2 = "avvicinarsi"
            avvicinarsi1 = "avvicinarvi"
            allontanarsi = "vi allontana"
            allontanarsi1 = "vi siete allontanati troppo avete"
            scontrarsi = "e vi siete scontrati"
            continuare = "continuate"
            fare = "ci siete quasi fate"
            dovere = "dovete"
            ti_vi = "vi"
		end if
		contSi = 0
		contNo = 0
		Motivato = 0
		narrazione=""
		narrazione =  Autista & " " & volere & "  raggiungere " & Destinazione & " ?"
        narrazione = narrazione & "NO!   Mancando " & Carburante & " per raggiungere " & Destinazione & " , " & Autista & " nel contesto " & Luogo & " non " & raggiungere & " l'obiettivo ! "
        narrazione = narrazione & " " & Autista & " " & volere & " raggiungere " & Destinazione & " ?"
        narrazione = narrazione &"   " & Autista & "  quale  " & Strada & " " & scegliere & " ?  "
        narrazione = narrazione &"ATTENZIONE  " & Autista & "  la scelta  " & Strada_KO & " " & allontanarsi & " da  " & Destinazione&"."
        narrazione = narrazione &" " & Cespugli & " " & ti_vi & " segnalano il pericolo ! "
        narrazione = narrazione &" :-(  " & Autista & "  " & allontanarsi1 & " scelto la strada chiusa  " & Strada_KO & " " & scontrarsi & " con " & Lupo & ".  "
        narrazione = narrazione &"  " & Autista & "  per risolvere la situazione " & dovere & "  abbandonare  " & Cestino & " cosi' da " & avvicinarsi1 & " a " & Destinazione & ".  "
        narrazione = narrazione &"  " & Autista & "  quale  " & Strada & " " & scegliere & " ?  "
        narrazione = narrazione &"  " & Autista & "  la scelta  " & Strada_OK & " " & avvicinarsi & " a  " & Destinazione & "  " & continuare & " così !  "
        narrazione = narrazione &" Coraggio " & fare & " l'ultimo passo ! '"
        narrazione = narrazione &" :-) COMPLIMENTI  " & Autista & " " & avere & " raggiunto " & Destinazione & "!!!"
		  ' errore=0
		   
'		  ' response.write(len(Autista) &" " & len(Destinazione)&" "  & len(Carburante)&" " & len(Luogo)&" " & len(Strada)&" " & len(Strada_OK)&" " & len(Strada_KO)&" " & len(Distanza)&" " & len(Cespugli)&" "& len(Lupo)&" "& len(Luogo)&" "& len(Cestino)&" ")
  '  if ((len(Autista)=0) or (len(Destinazione)=0) or (len(Carburante)=0) or (len(Luogo)=0) or (len(Strada)=0) or (len(Strada_OK)=0) or(len(Strada_KO)=0) or(len(Distanza)=0) or(len(Cespugli)=0) or(len(Lupo)=0) or(len(Luogo)=0) or(len(Cestino)=0)) then
'   errore=2
'   end if 
' 
'   
   Case Cartella&"_U_2_8"  
  	 ID=Request.QueryString("CodiceMetafora")
       SoggettoC = ucase(Request("txtSoggettoC"))
	   SoggettoC = Replace(SoggettoC, Chr(34), "'")
	   SoggettoC=  Replace(SoggettoC,"'",Chr(96))
  
	   DomandaC = ucase(Request("txtDomandaC"))
	   DomandaC = Replace(DomandaC, Chr(34), "'")
	   DomandaC =  Replace(DomandaC,"'",Chr(96))
	
	
	   MotivazioneC = ucase(Request("txtMotivazioneC"))
	   MotivazioneC = Replace(MotivazioneC, Chr(34), "'")
	   MotivazioneC =  Replace(MotivazioneC,"'",Chr(96))
	
	   DesiderioC = ucase(Request("txtDesiderioC"))
	   DesiderioC = Replace(DesiderioC, Chr(34), "'")
	   DesiderioC=  Replace(DesiderioC,"'",Chr(96))
	   BisognoC = ucase(Request("txtBisognoC"))
	   BisognoC = Replace(BisognoC, Chr(34), "'")
	   BisognoC =  Replace(BisognoC,"'",Chr(96))
	
	   SoggettoS = ucase(Request("txtSoggettoS"))
	   SoggettoS = Replace(SoggettoS, Chr(34), "'")
	   SoggettoS =  Replace(SoggettoS,"'",Chr(96))
	   
	   RispostaS = ucase(Request("txtRispostaS"))
	   RispostaS = Replace(RispostaS, Chr(34), "'")
	   RispostaS=  Replace(RispostaS,"'",Chr(96))
	   
	   MotivazioneS = ucase(Request("txtMotivazioneS"))
	   MotivazioneS = Replace(MotivazioneS, Chr(34), "'")
	   MotivazioneS =  Replace(MotivazioneS,"'",Chr(96))
	   
	   
	   DesiderioS = ucase(Request("txtDesiderioS"))
	   DesiderioS = Replace(DesiderioS, Chr(34), "'")
	   DesiderioS=  Replace(DesiderioS,"'",Chr(96))
	   
		BisognoS = ucase(Request("txtBisognoS"))
	   BisognoS = Replace(BisognoS, Chr(34), "'")
	   BisognoS=  Replace(BisognoS,"'",Chr(96))
		
	   TipoEvento = Request("txtTipoEvento")
	    
	       
		   TolleranzaC=Request("txtTolleranzaC")
	   	   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		  
end select
'
'  
'  
'   
' 
'response.write("CodiceTest="&CodiceTest)
 
 ThreadParent=0
 
if (errore=0) then 
     

	'response.write(CodiceTest)
         Select case CodiceTest
		 Case Cartella&"_U_2_3"  
			  QuerySQL="INSERT INTO M_Topolino (Topolino, Formaggio, Fame,Labirinto,Strada,Strada_OK,Strada_KO,Testata,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ThreadParent,Id_Premetafora) SELECT '" & Topolino & "','" & Formaggio & "', '" & Fame & "','" & Labirinto & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Testata & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  &  FormatDateTime(now, 2) & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4)& "',"& ThreadParent &","& ID_Premetafora &";" 
			  QuerySQL1="select max(CodiceMetafora) from M_Topolino"
			  ConnessioneDB.Execute(QuerySQL)
			  set rsMaxId=ConnessioneDB.execute(QuerySQL1)
			  ID=rsMaxId(0)
			  QuerySQL ="UPDATE M_Topolino SET ThreadParent = '" & ID & "' WHERE CodiceMetafora =" &ID&";"
			  ConnessioneDB.Execute QuerySQL
			  

		Case Cartella&"_U_2_5" 
		
			
		    QuerySQL="INSERT INTO M_Navigazione (Autista, Destinazione, Carburante,Luogo,Strada,Strada_OK,Strada_KO,Cespugli,Lupo,Cestino,Distanza,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ThreadParent,Id_Premetafora) SELECT '" & Autista & "','" & Destinazione & "', '" & Carburante & "','" & Luogo & "','" & Strada & "','" & Strada_OK & "','" & Strada_KO & "','" & Cespugli & "','" & Lupo & "','" & Cestino & "','" & Distanza & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  &  FormatDateTime(now, 2) & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4)& "',"& ThreadParent &","& ID_Premetafora &";" 
			QuerySQL1="select max(CodiceMetafora) from M_Navigazione"
			ConnessioneDB.Execute(QuerySQL)
			set rsMaxId=ConnessioneDB.execute(QuerySQL1)
			ID=rsMaxId(0)
			QuerySQL ="UPDATE M_Navigazione SET ThreadParent = '" & ID & "' WHERE CodiceMetafora =" &ID&";"
			ConnessioneDB.Execute QuerySQL
			
			
			
			
			
		Case Cartella&"_U_2_8"  
		  
		    QuerySQL="INSERT INTO M_Desideri (SoggettoC, DomandaC, MotivazioneC,DesiderioC,BisognoC,SoggettoS,RispostaS,MotivazioneS,DesiderioS,BisognoS,TipoEvento,TolleranzaC,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ThreadParent,Id_Premetafora) SELECT '" & SoggettoC & "','" & DomandaC & "', '" & MotivazioneC & "','" & DesiderioC & "','" & BisognoC & "','" & SoggettoS & "','" & RispostaS & "','" & MotivazioneS & "','" & DesiderioS & "','" & BisognoS & "','" & TipoEvento & "'," & TolleranzaC & ",'" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  &  FormatDateTime(now, 2) & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) & "'," & ThreadParent &","& ID_Premetafora &";"  
			QuerySQL1="select max(CodiceMetafora) from M_Desideri"
		
		end select
	
	' response.write(QuerySQL)
	
	
	   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"  
	   url=Replace(url,"\","/")
	   url2=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_simula_"&ID&".txt"  
	   url2=Replace(url2,"\","/")
'	
'	
'	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE e LA SIMULAZIONE DELLA METAFORA
 
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'response.write("<br>url="&url)
	'response.write("<br>url2="&url2)
	'On Error Resume Next
	if instr(Spiegazione,"<script>")<>0 then
	   Spiegazione=Replace(Spiegazione,"<script>","")
	   Spiegazione=Replace(Spiegazione,"</script>","")
	end if
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
    objCreatedFile.WriteLine(Spiegazione)
	objCreatedFile.Close

	if instr(narrazione,"<script>")<>0 then
	   narrazione=Replace(narrazione,"<script>","")
	   narrazione=Replace(narrazione,"</script>","")
	end if
	Set objFSO2 = CreateObject("Scripting.FileSystemObject")
	Set objCreatedFile2 = objFSO2.CreateTextFile(url2, True)
    objCreatedFile2.WriteLine(narrazione)
	objCreatedFile2.Close
 
	If Err.Number = 0 Then
	  stato="Inserimento effettuato correttamente"
	  CodiceMetafora=ID
	  response.write(" { "  &_
 """stato"": """ & stato& """," &_
 """id"": """ & CodiceMetafora & """}")

	Else
	   stato=Err.Description 
	   Err.Number = 0
	   response.write(" { "  &_
 """stato"": """ & stato& """," &_
 """id"": """ & CodiceMetafora & """}")

	End If
	
		
 
   else
		stato=errore &" Controlla che non ci siano campi lasciati vuoti"
		CodiceMetafora=0
		response.write(" { "  &_
 """stato"": """ & stato& """," &_
 """id"": """ & CodiceMetafora & """}")

		
   

end if 
'




   %>
	 
 