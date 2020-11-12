&#65279;&#65279;<%@ Language=VBScript %>
<html>
<head>	
<link rel="stylesheet" type="text/css" href="../../stile.css">
<style>
<!--
li.MsoNormal
{mso-style-parent:"";
margin-bottom:.0001pt;
font-size:12.0pt;
font-family:"Times New Roman";
margin-left:0cm; margin-right:0cm; margin-top:0cm}
-->
</style>
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
	location.href="../home.asp"
	//location.href=window.history.back();
	}
	</script>
	
	</head>
	
	<%
	Response.Buffer = true
' On Error Resume Next  
' per il controllo della validità della sessione, se è scaduta -> nuovo login
	if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	<BODY onLoad="showText2();"> </BODY>
	<% else %>
	<body bgcolor="#FFFFFF">
	<% end if %>
	<div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">
	
	
	
	<% Response.Buffer=True 
	Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet,k
	Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
'Apertura della connessione al database  
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	%>   
	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<%  
'Scadenza=Request.Form("txtScadenza")   
	Scadenza=Request.QueryString("Scadenza")
	Num=Request.QueryString("txtNum")
	if Scadenza<>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
		Scadenza=cdate(Request.QueryString("Scadenza"))
	end if 
	Capitolo=Request.QueryString("Capitolo")
	Paragrafo=Request.QueryString("Paragrafo")
	Modulo=Request.QueryString("Modulo")
	by_UECDL=Request.QueryString("by_UECDL")
	
	BoxApro=Request.QueryString("BoxApro")
	Segnalibro=Request.QueryString("Segnalibro")
	Sottoparagrafo=Request.QueryString("Sottoparagrafo")
	CodiceSottopar = Request.QueryString("CodiceSottopar") 
	
	function ReplaceCar(sInput)
		dim sAns
		sAns = Replace(sInput,chr(224),"a"&Chr(96))
		sAns = Replace(sAns,chr(225),"a"&Chr(96))
		sAns = Replace(sAns,chr(232),"e"&Chr(96))
		sAns = Replace(sAns,chr(233),"e"&Chr(96))
		sAns = Replace(sAns,chr(236),"i"&Chr(96))
		sAns = Replace(sAns,chr(237),"i"&Chr(96))
		sAns = Replace(sAns,chr(242),"o"&Chr(96))
		sAns = Replace(sAns,chr(243),"o"&Chr(96))
		sAns = Replace(sAns,chr(249),"u"&Chr(96))
		sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
		sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
		sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
		sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
		sAns=  Replace(sAns,"&","e") 
		sAns=  Replace(sAns,"/","-") 
		sAns=  Replace(sAns,"\","-") 
		sAns=  Replace(sAns,"?",".") 
		sAns=  Replace(sAns,"*","x") 
		sAns=  Replace(sAns,"<","_")
		sAns=  Replace(sAns,">","_") 
		
		ReplaceCar = sAns
	end function
	if CodiceSottopar<>"" then
		QuerySQL="Select count(*) from preDomande where Id_Paragrafo='"&Paragrafo&"' and Id_Sottoparagrafo='"&Codicesottopar&"' ;"
	else
		
		QuerySQL="Select count(*) from preDomande where Id_Paragrafo='"&Paragrafo&"';"
	end if
	set rsTabella=ConnessioneDB.Execute (QuerySQL) 
	if rsTabella(0)>0 then
		QuerySQL="Select max(Posizione) from preDomande where Id_Paragrafo='"&Paragrafo&"';"  
		set rsTabella=ConnessioneDB.Execute (QuerySQL) 
		contPos=rsTabella(0)
	else
		contPos=0
	end if 
	cont=0
	
	if  txtNum<>"" then
		
		
		for k=1 to Num
			
			Domanda = Request.Form("txtDomanda"&k)   
			Domanda = ReplaceCar(Domanda) 
			if Domanda<>"" then ' controllo per le righe vuote 
				
'Esecuzione della query per  
'QuerySQL="INSERT INTO preDomanda (Id_Mod, Id_Paragrafo,Quesito,Eseguita) SELECT '" & Modulo & "','" & Paragrafo "','" & Domanda & "'," & 0 & "';"
				if Scadenza<>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
					QuerySQL="  INSERT INTO preDomande (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & (contPos+k) & ",'" & Scadenza & ",'" & CodiceSottopar & "';"
				else
					QuerySQL="  INSERT INTO preDomande (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & (contPos+k) & ",'" &CodiceSottopar & "';"
					
				end if 
			end if
			response.write(QuerySQL)
'end if 
			ConnessioneDB.Execute QuerySQL 
			
			
		next 
		
		
	else  
		Scadenza=Request.Form("date3")
		strText = Request.Form("MyTextArea")
'strText = txtDomande
		arrLines = Split(strText, vbCrLf)
		k=1
		For Each strLine in arrLines
			img=0
			cFile=0
			if instr(strLine,"$")=0 and instr(strLine,"#")=0 then ' senza immagine nè file
				img=0
				cFile=0
				Domanda=strLine
			else ' immagine
				if instr(strLine,"$")<>0 then ' immagine
					img=1
					Domanda=left(strLine,instr(strLine,"$")-1)
				end if
				if instr(strLine,"#")<>0 then ' file
					cFile=1
					if img=0 then 
						Domanda=left(strLine,instr(strLine,"#")-1)
					end if
				end if
				
				
			end if
			
			Domanda =  ReplaceCar(Domanda) 
			
			if Scadenza<>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
				QuerySQL="  INSERT INTO preDomande (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & (contPos+k) & ",'" & Scadenza & "','" & CodiceSottopar & "';"
			else
				QuerySQL="  INSERT INTO preDomande (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & (contPos+k) & ",'" &CodiceSottopar & "';"
				
			end if 
' response.write(QuerySQL)
			ConnessioneDB.Execute QuerySQL 
			
			response.write Domanda & "<br>"
			k=k+1
		Next
		
		
		end if%>
		
		
		
		
		
		<%response.write(QuerySQL)
		
		
		
		
		On Error Resume Next
		If Err.Number = 0 Then
			Response.Write "Inserimento avvenuto! Stai per essere reindirizzato al libro "
		Else
			Response.Write Err.Description 
			Err.Number = 0
		End If
		
		
		%>
		</font>   
		
		<h4><a href="inserisci_predomande.asp?Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceSottopar=<%=CodiceSottopar%>">Continua ...</a></h4>
		<p>&nbsp;</p>
		<div id=piede_pagina>
		<p><p>
		<%if by_UECDL<>"" then %>
		<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
		<h3><a href="../cClasse/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Apprendimento... </a></h3> 
		<%else%>
			<!-- REDIRECT INTELLIGENTE  -->
		<h3><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Apprendimento... </a></h3> 
		
		<%end if%>	
		
		
		<br>  <a  id="msg" href="#modal-1" role="button" class="btn notify" data-notify-title="Inserimento avvenuto" data-notify-message="Stai tornando al libro...">Stai per essere reindirizzato ....</a>       
		
					   
		
		
		
		<% ' con timer
		Response.AddHeader "REFRESH","2;URL=../cClasse/home_app.asp?dividApro="&BoxApro&"&id_classe="&session("Id_Classe")&"#"&BoxApro-3
		%>
		
		
		</div>
		<!-- se il login è corretto richima la pagina per inserire le domande del test -->
		
		
		</body>
		</html>
		
