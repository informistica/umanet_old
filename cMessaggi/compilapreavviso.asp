<%@ Language=VBScript %>
 <!--#include file = "../include/header.asp"-->
<% 


   
  Id_Par=Request.QueryString("Id_Par")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Id_Mod=Request.QueryString("Id_Mod")
  ID_Classe=Session("Id_Classe")
  idxSel=Request.QueryString("idxSel")
  idxSelPar=Request.QueryString("idxSelPar")
  classeCont = session("id_classe")
  
  response.write("<br>")
  'response.write("Classe: "&classeCont)
  'response.write("Variabili: <br>")
  'response.write("Paragrafo: "&idxSel)
  'response.write("<br>Modulo: "&idxSelPar&"<br><br>")
  
 ' Id_Mod=idxSel
'if  
byUmanet=Request.QueryString("byUmanet")
'Session("byUmanet")=byUmanet
 Dim ConnessioneDB , rsTabella,QuerySQL
  
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
%> 
 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
if byUmanet="" then
 querySQL="SELECT DISTINCT (Titolo) AS TitoloMod, Id_Classe,ID_Mod,Cartella,Classe,Posizione FROM MODULI_NOT_UMANET WHERE Id_Classe='"& ID_Classe &"'" &_
" ORDER BY MODULI_NOT_UMANET.Posizione;"
else
 querySQL="SELECT DISTINCT (Titolo) AS TitoloMod, Id_Classe,ID_Mod,Cartella,Classe,Posizione FROM MODULI_UMANET1 WHERE Id_Classe='"& ID_Classe &"'" &_
" ORDER BY MODULI_UMANET1.Posizione;"
end if
'response.write(querySQL)
'response.write("byUmanet="&byUmanet)
Set rsTabellaMod = ConnessioneDB.Execute(QuerySQL) 
if idxSel="" then
  if Session("idxSel")<>"" then
   idxSel=Session("idxSel")
   'idxSelPar=Session("idxSelPar")
   else
		if classeCont <> "6COM" then
			idxSel=1
			idxSelPar=1
		else
			idxSel=0
			idxSelPar=0
		end if
   end if
end if
		
		%> 
        <html>
<head>
</head>
<body> 
   <div class="box-content">
		<form action="#" method="POST" class='form-horizontal'>  
        
       							 <div class="control-group">
										<label for="select" class="control-label">Modulo</label>
										<div class="controls">
											<select  id="select" class='input-large' name="txtModulo" onChange="window.document.location='compilapreavviso.asp?byUmanet=<%=byUmanet%>&idxSel='+this.options[this.selectedIndex].value;+'&Id_Mod='+this.options[this.selectedIndex].value;">
											
         
	  <%
		  
		 if classeCont <> "6COM" then
			cont = 1
		 else
			cont = 0
		end if
			
		 do while not rsTabellaMod.eof%>
         <% if cont=cint(idxSel) then %> 
         <%Session("PosMod")=rsTabellaMod("Posizione")
		   Session("IdMod")=rsTabellaMod("ID_Mod")
		   Session("idxSel")=rsTabellaMod("Posizione")
		   %>
			<option selected="selected" value="<%=rsTabellaMod("Posizione")%>"><%=rsTabellaMod("TitoloMod")%></option>
			<%else%>
            <option  value="<%=rsTabellaMod("Posizione")%>"><%=rsTabellaMod("TitoloMod")%></option>   
            <%end if%>
		    <%
			 cont=cont+1
			 rsTabellaMod.movenext
		 loop
	  %>
	</select>
    </div>
	 
    
    <br />
    
    <% 
	if byUmanet="" then
	querySQL="SELECT  Titolo,PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar  FROM MODULI_NOT_UMANET " &_
	" WHERE Id_Classe='"&Id_Classe &"' and PosMod="& idxSel &"" &_
" ORDER BY PosPar;"
   else
  querySQL="SELECT  Titolo,PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar  FROM MODULI_UMANET1" &_
	" WHERE Id_Classe='"&Id_Classe &"' and PosMod="& idxSel &"" &_
" ORDER BY PosPar;"
   end if

 'response.write(querySQL)
Set rsTabellaPar = ConnessioneDB.Execute(QuerySQL) 
 %>		
 								
										<label for="select" class="control-label">Paragrafo</label>
										<div class="controls">
<select name="txtPar" onChange="window.document.location='compilapreavviso.asp?byUmanet=<%=byUmanet%>&idxSelPar='+this.options[this.selectedIndex].value;"> 
<%		  
		 cont=1
		 do while not rsTabellaPar.eof%>
         <% if cont=cint(idxSelPar) then %> 
         <% 'Session("ID_ModSel")=rsTabellaPar("ID_Mod")
		    Session("ID_ParSel")=rsTabellaPar("ID_Paragrafo")
			Session("PosPar")=rsTabellaPar("PosPar") 
			 
		 %>
			<option selected="selected" value="<%=rsTabellaPar("PosPar")%>"><%=rsTabellaPar("TitPar")%></option>
			<%else%>
            <option  value="<%=rsTabellaPar("PosPar")%>"><%=rsTabellaPar("TitPar")%></option>   
            <%end if%>
		    <%
			 cont=cont+1
			 rsTabellaPar.movenext
		 loop
	  %>
	</select>
    </div>
    </div>
<%
' devo calcolare il segnalibro , cioè il numero del paragrafo in totale per il box da aprire
if byUmanet="" then
querySQL="SELECT  Titolo, PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar FROM MODULI_NOT_UMANET " &_
	" WHERE Id_Classe='"&Id_Classe &"'" &_
" ORDER BY PosMod,PosPar;"
 else
 querySQL="SELECT  Titolo, PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar FROM MODULI_UMANET1" &_
	" WHERE Id_Classe='"&Id_Classe &"'" &_
" ORDER BY PosMod,PosPar;"
 end if
cont=1
Set rsTabellaPos = ConnessioneDB.Execute(QuerySQL) 
	do while not rsTabellaPos.eof
	   if strcomp(Session("ID_ParSel"),rsTabellaPos("ID_Paragrafo"))=0 then
		segnalibro=cont 
		'response.write("<br>" & Session("ID_ParSel") & "="& rsTabellaPos("ID_Paragrafo") &"Cont="&cont)
	end if
    cont=cont+1 'veniva fatoo cont = cont+0 per cui non veniva incrementato il numero del totale del box -> rimaneva sempre 1
	rsTabellaPos.movenext
loop
Session("idBox")=segnalibro
'response.write("Numero box: "&segnalibro)

'divid_nuovo = Right(session("ID_ParSel"), 1)
'divid_nuovo = divid_nuovo
'session("divid_nuovo") = divid_nuovo
%>
</form>
</div>
</body>
</html>