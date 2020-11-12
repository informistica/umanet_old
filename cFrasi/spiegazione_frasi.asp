<!-- esegui_test_MODBC3.asp -->

<%@ Language=VBScript %>
<% Response.Buffer=True %>

<HTML>
<HEAD>
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
<TITLE>FRASI DELLA RETE</TITLE>
</HEAD>
<%Response.Buffer = true
  On Error Resume Next  
  %>
  

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

 Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato

    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
 
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")

 
 'per il copia incolla
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	rsTabellaCI.close
 
' codice per permettere la visualizzazione solo delle proprie domande 
'QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close
'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
 
   ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
      <% if (CIAbilitato=0) then ' disabilito copia incolla%>
        <body  oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="Effect.toggle('dAttività','BLIND');Effect.toggle('dAvvisi','BLIND'); return false;">  
        <%else%>
        <body> 
        <%end if%>
  <% end if %>

<%Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  		'response.write("Stato"&stato)				
%>
 <%'response.write("Stati :  " & stato & " " & stato0)%>
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color="#FF0000">SPIEGAZIONE FRASI:</font></b>  
<!-- stampa il titolo del test -->

 <table border="0" align=center width="60%" id="background-image">
     <thead>
		<tr>
			<th colspan=3 align=center>
			  <font color="#000000" size="+1"> <%=Capitolo%>  </font>
			</th>
		</tr>
		<tr>
			<th colspan=3 align=center>
			  <font color="#000000" size="+1"> <%=Paragrafo%></h4></b></font>
			</th>
		</tr>
     </thead>
		
	</table>
 
 
  
	<br>
<%   
  

 
if (clng(Stato)=0) or (clng(Stato0)=0) then 
' 'Definzione codice SQl della query per ricercare le frasi del paragrafo 
   if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le frasi del PARAGRAFO altrimenti solo quelle dello       studente loggato  
  
	QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Frasi ON Allievi.CodiceAllievo = Frasi.Id_Stud) ON Moduli.ID_Mod = Frasi.Id_Mod) ON Paragrafi.ID_Paragrafo = Frasi.Id_Arg" &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
	" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Frasi.In_Quiz<>0 " &_   
	" ORDER BY Paragrafi.ID_Paragrafo;"

   else
   QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Frasi ON Allievi.CodiceAllievo = Frasi.Id_Stud) ON Moduli.ID_Mod = Frasi.Id_Mod) ON Paragrafi.ID_Paragrafo = Frasi.Id_Arg" &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
	" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Frasi.In_Quiz<>0  and Frasi.Id_Stud='"& Session("CodiceAllievo")& "'" &_   
	" ORDER BY Paragrafi.ID_Paragrafo;"
   end if

else 
 
	if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le frasi del MODULO altrimenti solo quelle dello       studente loggato  
   							'0						1				2					3			

		QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
		" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Frasi ON Allievi.CodiceAllievo=Frasi.Id_Stud) ON Moduli.ID_Mod=Frasi.Id_Mod) ON Paragrafi.ID_Paragrafo=Frasi.Id_Arg" &_
		" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
		" HAVING Moduli.ID_Mod='" & Modulo & "' AND Frasi.In_Quiz<>0 " &_ 
		" ORDER BY Paragrafi.ID_Paragrafo;"
    else
	    QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
		" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Frasi ON Allievi.CodiceAllievo=Frasi.Id_Stud) ON Moduli.ID_Mod=Frasi.Id_Mod) ON Paragrafi.ID_Paragrafo=Frasi.Id_Arg" &_
		" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Frasi.Chi, Frasi.CodiceFrase, Moduli.ID_Mod,Frasi.In_Quiz,Frasi.Cartella,Frasi.Id_Stud" &_
		" HAVING Moduli.ID_Mod='" & Modulo & "' AND Frasi.In_Quiz<>0 and Frasi.Id_Stud='"& Session("CodiceAllievo") & "'" &_
		" ORDER BY Paragrafi.ID_Paragrafo;"
	end if 


end if    
    
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'response.Write(QuerySQL)
      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%><center>
  <H4>Frasi non ancora disponibile!</h4></center>
  
<% Else
  
  i=1 'inizializza la variabile i (contatore delle domande)
  Do until rsTabella.EOF
  		 
 
    ID=rsTabella(4)
   url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
   url=Replace(url,"\","/")
 
               ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url2="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFrasi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.Close
'response.write(url)
' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll
'sReadAll = url
objTextFile.Close   ' la soluzione seguente la rimuovo e dirò di copiare ed incollare la domanda plus nella spiegazione
' così da avere il livello di apprendimento comprensibile , diversamente dovrei prevedere il modo di far apparire il testo della domanda plus 
' anche nell'approfondimento di fine quiz.
 
                
%>

  
  
    
  
  <table border="1"  align=center width="60%" id="zebra_stud">
		<tr>
			<td align="center" colspan="3"><font size="+0"><%=rsTabella(0)%></font></td>
			 
			
		</tr>
		<tr><td width="10%"><font size="-2"><%=rsTabella(2)%></font></td>
			<td colspan=2>
			<p align="center"><b><%=rsTabella(3)%></b></td>
			 
		</tr>
		<tr>
			<td colspan=3>
			
			<p align="center">
			 <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" cols="100"><%
			 ' if clng(rsTabella(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			 
			 Response.write(sReadAll)%> </textarea><br>
	      </td>
		 
		</tr>
	</table>
	<br>
<%    

       i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
   <div class="citazioni">
  <a href="../cClasse/scegli_azione_app.asp?Cartella=<%=Cartella%>&Stato=1&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>">	Indietro </a>
  </div>


  </div>
  </div>
</BODY>
<% 'else 
   'Response.Redirect "../home.asp"
   '   end if %>
</HTML>