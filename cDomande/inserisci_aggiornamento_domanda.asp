<%@ Language=VBScript %>
 
  
  <%   
   
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   
        	  
          
     <%            
 
	
   'voto=clng(Request.QueryString("Voto"))
   
   function ReplaceCar(sInput)
dim sAns
  
  sAns = sInput
  'sAns1 = sInput
  
 sAns = Replace(sInput,chr(236),"i"&Chr(96))
 sAns = Replace(sAns,chr(237),"i"&Chr(96))
 sAns = Replace(sAns,chr(242),"o"&Chr(96))
 sAns = Replace(sAns,chr(243),"o"&Chr(96))
 sAns = Replace(sAns,chr(249),"u"&Chr(96))
 sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
 sAns = Replace(sAns,chr(133),"a"&Chr(96))
 sAns = Replace(sAns,chr(138),"e'")
 sAns = Replace(sAns,"é","e'")
  sAns = Replace(sAns,chr(130),"e'")
 sAns = Replace(sAns, Chr(34), "'") 'sostituisco gli apici " con l'apice singolo
 sAns=  Replace(sAns,"'",Chr(96))  'sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
 sAns=  Replace(sAns,chr(58),Chr(44))  'sostituisco : con , per non disturbare la creazione del file
 sAns=  Replace(sAns,"&","e") 
 sAns=  Replace(sAns,"/","-") 
 sAns=  Replace(sAns,"\","-") 
 sAns=  Replace(sAns,"?",".") 
 sAns=  Replace(sAns,"*","x") 
 sAns=  Replace(sAns,"<","_")
 sAns=  Replace(sAns,">","_") 
   sAns = Replace(sAns,"è","e'" )
   
   sAns=  Replace(sAns,"«",Chr(96))
   sAns=  Replace(sAns,"»",Chr(96))
   sAns=  Replace(sAns,"à","a'")
   sAns=  Replace(sAns,"ò","o'")
   sAns=  Replace(sAns,"ù","u'")
   sAns = Replace(sAns,"’","'")
   sAns = Replace(sAns,"“","'")
   sAns = Replace(sAns,"”","'")
   sAns=  Replace(sAns,"'",Chr(96))
   sAns = Replace(sAns, "È", "E'")
   sAns = Replace(sAns, "ì", "i'")
   sAns = Replace(sAns, "–", "-")
   sAns=  Replace(sAns,"'",Chr(96))
   sAns=  Replace(sAns, vbcrlf,"")
   sAns=  Replace(sAns, chr(13),"")
   sAns=  Replace(sAns, chr(10),"")
   
   
   
   'sAns = Replace(sAns,VBCrlf,"<br>")
    
ReplaceCar = sAns

end function
   
   
   CodiceDomanda=Request("CodiceDomanda")
   quesito=Request("quesito")
   ID=Request("id")
   R1 = ReplaceCar(Request("r1"))
   R2 = ReplaceCar(Request("r2"))
   R3 = ReplaceCar(Request("r3"))
   R4 = ReplaceCar(Request("r4"))
   RE = ReplaceCar(Request("re"))
   url=Request("url")
   Spiegazione=ReplaceCar(Request("spiegazione"))
   Segnalata=Request("segnalata")
  ' response.write("<br>"&Spiegazione)
   

 
    QuerySQL ="UPDATE Domande SET Segnalata = " & Segnalata & ", Quesito = '" & quesito & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "'   WHERE CodiceDomanda =" &ID&";"
   ' response.write(QuerySQL)
 
	ConnessioneDB.Execute(QuerySQL)
 	 
  


'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Err.Number=0
objFSO.DeleteFile url
Set objCreatedFile = objFSO.CreateTextFile(url, True)
objCreatedFile.WriteLine(Spiegazione)
objCreatedFile.Close

' Set objFSO1 = CreateObject("Scripting.FileSystemObject")
' url1="C:\inetpub\umanetroot\expo2015Server\logajax.txt"
' Set objCreatedFile1 = objFSO1.CreateTextFile(url1, True)
' objCreatedFile1.WriteLine(QuerySQL) 
' objCreatedFile1.Close
 

'On Error Resume Next
If Err.Number = 0 Then
Response.Write "ok"
Else
Response.Write "ko"
End If%>

 