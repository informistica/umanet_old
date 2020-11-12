<%@ Language=VBScript %>
<%  
 ' Id_Par=Request.QueryString("Id_Par")
 ' Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Id_Mod=Request.QueryString("Id_Mod")
  ID_Classe=Session("Id_Classe")
  idxSel=Request.QueryString("idxSel")
  idxSelPar=Request.QueryString("idxSelPar")
 ' Id_Mod=idxSel
  byUmanet=request.QueryString("byUmanet") ' vale 1 se sono chiamato per trasferire un modulo umanet
 
 ' è settato se devo condivieder un modulo tra più classi
  condividi=request.QueryString("condividi")
 Dim ConnessioneDB , rsTabella,QuerySQL
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
    Set ConnessioneDBStory = Server.CreateObject("ADODB.Connection") 
'Session("DB2")=1	
%>
<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<!-- #include file = "../stringhe_connessione/stringa_connessione_story.inc" -->
<%
   'Apertura della connessione al database
 
  
     

   response.write(Server.MapPath("/"))  
%> 
 
 
<% 	


QuerySQL="SELECT Cartella from Classi where ID_Classe='"&ID_Classe&"';" 
Set rsTabellaCartella = ConnessioneDB.Execute(QuerySQL) 
Cartella=rsTabellaCartella(0)

QuerySQL="SELECT Home,Anno, Url,Posizione,Materia from DB_Story;" 
Set rsTabellaMod = ConnessioneDB.Execute(QuerySQL) 
'response.write(QuerySQL&"   "& rsTabellaMod(0) )
if idxSel="" then
  if Session("idxSel")<>"" then
   idxSel=Session("idxSel")
   else
   idxSel=0
   Session("idxSel")=0
   end if
end if

'if idxSel="" then
'  if Session("idxSel")<>"" then
'   idxSel=Session("idxSel")
'  else
'   idxSel=1
'   Session("idxSel")=1
'   Session("PosMod")=rsTabellaMod("Anno")
'   Session("UrlDB")=rsTabellaMod("Url")
'   Session("homesito_origine")=rsTabellaMod("Home")
'  end if
'end if
		
		   
		    	'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logSelezionaorigine.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine("idxSel="&idxSel&" session="&Session("idxSel"))
'			 
'		
		
		%>  
        
        Gestire anche la posizione in cui deve essere inserito il modulo<br /><br />
          <b>Database&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b>
         <select name="txtModulo" onChange="window.document.location='seleziona_origine.asp?condividi=<%=condividi%>&byUmanet=<%=byUmanet%>&idxSel='+this.options[this.selectedIndex].value;+'&Id_Mod='+this.options[this.selectedIndex].value;">
	  <%		
		 cont=0
		 ' lo metto per evitare errore se viene premuto il bottone senza aver selezionato nulla, lasciando il default
		 ' preseleziono il primo db
		   Session("PosMod")=rsTabellaMod("Anno")
		   Session("UrlDB")=rsTabellaMod("Url")
		   Session("homesito_origine")=rsTabellaMod("Home")
		   Session("idxSel")=cont
		   Session("idxMat")=rsTabellaMod("Materia") ' tolto, settato in home.asp
		  '  objCreatedFile.WriteLine("S(PosMod)="&Session("PosMod")&" S(URLDB)="&Session("UrlDB")&" S(IdxSel)="&Session("idxSel") &"cont="&cont&" idxSel="&cint(idxSel)&"RsTabella(0)"&rsTabellaMod("Url"))
			 ' response.write("ci sono")
		 do while not rsTabellaMod.eof%>
         <% 
		 
		 if cont=cint(idxSel) then %> 
         <%	  	   
		   Session("PosMod")=rsTabellaMod("Anno")
		   Session("UrlDB")=rsTabellaMod("Url")
		   Session("idxSel")=cont
		   Session("idxMat")=rsTabellaMod("Materia")' tolto, settato in home.asp
		   Session("homesito_origine")=rsTabellaMod("Home")
		   'apro la connessione al db
		   '  Set ConnessioneDBStory = Server.CreateObject("ADODB.Connection")  
'			 ConnessioneDBStory.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
'              "DBQ=" & Server.MapPath(rsTabellaMod("Url"))  
'			  response.write("<br>"&rsTabellaMod("Url"))
'			  
'			   objCreatedFile.WriteLine("ci sono ")
'			 objCreatedFile.close
		   
		   %>
			<option selected="selected" value="<%=rsTabellaMod("Posizione")%>"><%=rsTabellaMod("Anno") &"  "%></option>
            
			<%else%>
            <option  value="<%=rsTabellaMod("Posizione")%>"><%=rsTabellaMod("Anno") &"  "%></option>   
            <%end if%>
		    <%
			 cont=cont+1
			 rsTabellaMod.movenext
		 loop
	  %>
	</select>
    
    <% 
	 if byUmanet<>"" then ' seleziono moduli umanet
	    querySQL="SELECT  distinct(ID_Mod),Titolo, Cartella,Posizione,Id_Classe FROM MODULI_UMANET order by ID_Mod,Cartella,Posizione; "  
	 else
	  querySQL="SELECT  distinct (ID_Mod),Titolo, Cartella,Posizione,Id_Classe FROM MODULI_NOT_UMANET order by  ID_Mod,Cartella,Posizione; "  
     end if
 'response.write(querySQL)
Set rsTabellaMod = ConnessioneDBStory.Execute(QuerySQL) 
 %>	<br><br>	
	
<b>Modulo</b>
<select name="txtMod" onChange="window.document.location='seleziona_origine.asp?condividi=<%=condividi%>&byUmanet=<%=byUmanet%>&idxSelPar='+this.options[this.selectedIndex].value;"> 
<%		  
		 cont=1 'preseleziono il primo item
		 Session("ID_ModSel")=rsTabellaMod("ID_Mod")
		 Session("PosMod")=rsTabellaMod("Posizione")
		 Session("TIT_ModSel")=rsTabellaMod("Titolo")
		  Session("ID_ClaSel")=rsTabellaMod("Id_Classe")
		 do while not rsTabellaMod.eof%>
         <% if cont=cint(idxSelPar) then %> 
         <% 'Session("ID_ModSel")=rsTabellaMod("ID_Mod")
		    Session("ID_ModSel")=rsTabellaMod("ID_Mod")
			Session("PosMod")=rsTabellaMod("Posizione") 
			Session("TIT_ModSel")=rsTabellaMod("Titolo")
			Session("ID_ClaSel")=rsTabellaMod("Id_Classe")
			
		 %>
			<option selected="selected" value="<%=cont%>"><%=rsTabellaMod("Id_Classe")&" - "&rsTabellaMod("Cartella") &" - "&rsTabellaMod("ID_Mod") & " -  " & rsTabellaMod("Titolo")%></option>
			<%else%>
            <option  value="<%=cont%>"><%=rsTabellaMod("Id_Classe")&" - "&rsTabellaMod("Cartella")&" - "&rsTabellaMod("ID_Mod") & " - " & rsTabellaMod("Titolo")%></option>   
            <%end if%>
		    <%
			 cont=cont+1
			 rsTabellaMod.movenext
		 loop
		 
	  %>
	</select>
    
	<%' byUmanet="" ' provvisorio da togliere il 29/11/2018 serve per trasferire modulo umanet da libro U a libro %>
    <%if condividi<>"" then%>
     <form method="POST"  action="trasferisci_modulo.asp?condividi=<%=condividi%>&byUmanet=<%=byUmanet%>&Url=<%=Session("UrlDB")%>&ID_ModSel=<%=Session("ID_ModSel")%>&ID_ClaSel=<%=Session("ID_ClaSel")%>&homesito_origine=<%=Session("homesito_origine")%>&cartella=<%=Cartella%>">
    <%else%> 
   <form method="POST"  action="trasferisci_modulo.asp?byUmanet=<%=byUmanet%>&Url=<%=Session("UrlDB")%>&ID_ModSel=<%=Session("ID_ModSel")%>&ID_ClaSel=<%=Session("ID_ModSel")%>&homesito_origine=<%=Session("homesito_origine")%>&cartella=<%=Cartella%>">
   <%end if%>
   <br />
  <b> <%response.write("ID del modulo ?")
  
  if byUmanet<>"" then ' seleziono moduli umanet
	    QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&Session("Cartella")&"';"
	else
	  QuerySQL="SELECT max(posizione) FROM MODULI_NOT_UMANET where Cartella='"&Session("Cartella")&"';"
	
     end if
 
	 
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	   if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1
	'InStrRev([inizio,]stringa1,stringa2[,compara])
	%> 
    </b> 
    <% if byUmanet<>"" then %>
     <input  type="text" name="txtID_Mod" size="7" value="<%=Session("Cartella")%>_U_<%=posizione%>"  > <p>
    <%else%>
    <input  type="text" name="txtID_Mod" size="7" value="<%=Session("Cartella")%>_<%=posizione%>"  > <p>
    <% end if%>
	<b><%response.write("Cartella risorse ?")%> </b>
     <input type="text" name="txtCartella" size="50" value="<%=Cartella%>"> <P>
	<b><%response.write("Titolo del nuovo modulo ?")%> </b>
     <input type="text" name="txtTitolo" size="50" value="<%=Session("TIT_ModSel")%>"> <P>
        <%if condividi<>"" then%>
        <input type="submit" value="Condividi" />
        <%else%>
      <input type="submit" value="Trasferisci" />
      
      <%end if%>
   </form>
  
<%
' devo calcolare il segnalibro , cioè il numero del paragrafo in totale per il box da aprire

'querySQL="SELECT  Paragrafi.Titolo, PosMod ,PosPar,ID_Mod,ID_Paragrafo FROM MODULI_NOT_UMANET " &_
'	" WHERE Id_Classe='"&Id_Classe &"'" &_
'" ORDER BY PosMod,PosPar;"
' 
'cont=1
'Set rsTabellaPos = ConnessioneDB.Execute(QuerySQL) 
'	do while not rsTabellaPos.eof
'	   if strcomp(Session("ID_ParSel"),rsTabellaPos("ID_Paragrafo"))=0 then
'		segnalibro=cont 
'		'response.write("<br>" & Session("ID_ParSel") & "="& rsTabellaPos("ID_Paragrafo") &"Cont="&cont)
'	end if
'    cont=cont+1
'	rsTabellaPos.movenext
'loop
'Session("idBox")=segnalibro
''response.write(segnalibro)
ConnessioneDBStory.close
ConnessioneDB.close

%>