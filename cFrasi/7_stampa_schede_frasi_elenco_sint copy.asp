<!-- modifica_domande.asp -->
<%@ Language=VBScript %>

<html>
<meta charset="utf-8">
<%
  Response.Buffer = true
  'On Error Resume Next
    ' per il controllo della validit� della sessione, se � scaduta -> nuovo login

  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO,i
  Dim ConnessioneDB,rsTabella, QuerySQL,CodiceTest,StringaConnessione
  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")

function ReplaceCar(sInput)
dim sAns

 'sostituzioni inutilizzate: non cancellare e non utilizzare fino a prova contraria
  sAns=  Replace(sInput,"'",Chr(96)) 'sostituisco l'apice ' con quello storto per non dist. la sintassi
 sAns=  Replace(sAns,Chr(39),Chr(96))
  sAns=  Replace(sAns,Chr(44),Chr(96))
  sAns = Replace(sAns,chr(146),Chr(96))
  sAns = Replace(sAns,chr(147),Chr(96))
  sAns = Replace(sAns,chr(148),Chr(96))
  sAns = Replace(sAns,chr(239),chr(96))

  sAns = Replace(sAns,"�",chr(96))
  sAns = Replace(sAns,"gradi",chr(248))
 'sAns = Replace(sAns, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto

  sAns=  Replace(sAns,Chr(34),"") ' sostituisco " con niente per non disturbare la sintassi
  sAns=  Replace(sAns,Chr(36),"")  ' rimuovo il simbolo $

  sAns = Replace(sAns, "  ", " ") 'sostituizione doppio spazio con uno singolo
  sAns = Replace(sAns, "	", " ") 'sostituzione spazi per evitare errori
  sAns = Replace(sAns, " ?", "?") ' rimozione spazio prima del punto di domanda
  sAns = Replace(sAns, "�", "'") ' sostituzione di un'apice con quello classico
  sAns = Replace(sAns, "�", "...") 'sostituzione tre puntini
  sAns = Replace(sAns, Chr(25), "'") 'sostituizione apice

  'sostituzione caratteri vari
  sAns = Replace(sAns,Chr(224),"a'") '�
  sAns = Replace(sAns,Chr(232),"e'") '�
  sAns = Replace(sAns,Chr(233),"e'") '�
  sAns = Replace(sAns,chr(236),"i'") '�
  sAns = Replace(sAns,chr(237),"i'") '�
  sAns = Replace(sAns,chr(242),"o'") '�
  sAns = Replace(sAns,chr(243),"o'") '�
  sAns = Replace(sAns,chr(249),"u'") '�
  sAns = Replace(sAns,chr(250),"u'") '�
  sAns = Replace(sAns, "&#8230;", "...")
  sAns = Replace(sAns, "&#224;","a'") '�
  sAns = Replace(sAns, "&#225;", "�") '�
  sAns = Replace(sAns, "&#249;","u'") '�
  sAns = Replace(sAns, "&#8217;", "'")
  sAns = Replace(sAns, "&#8211;", "-")
  sAns = Replace(sAns, "&#232;","e'") '�
  sAns = Replace(sAns, "&#233;","e'") '�
  sAns = Replace(sAns, "&#242;","o'") '�
  sAns = Replace(sAns, "&#171;","'")
  sAns = Replace(sAns, "&#187;","'")
  sAns = Replace(sAns, "&#8220;","'")
  sAns = Replace(sAns, "&#8221;","'")
  sAns = Replace(sAns, "&#236;","i'") '�
  sAns = Replace(sAns, "&#250;","u'") '�
  sAns = Replace(sAns, "&#176;",chr(248)) 'gradi
  sAns = Replace(sAns, "'", "'")
  sAns = Replace(sAns, "&quot;", "") 'sostituzione delle virgolette alte con niente per evitare errori json

  sAns = Replace(sAns,chr(96),chr(39)) ' sostituizione finale dell'apice storto con il classico apice
  'sAns1 = ucase(left(sAns,1)) ' maiuscola frasi
  'sAns2 = right(sAns, len(sAns)-1)

  'sAns = sAns1&sAns2



ReplaceCar = sAns
end function
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
           <!-- #include file = "../var_globali.inc" -->

    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<%
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceAllievo=Request.QueryString("CodiceAllievo")

 if (strcomp(session("CodiceAllievo"),CodiceAllievo) = 0) or (session("admin")=true) then


  'cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceFrase=Request.QueryString("CodiceFrase")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("CodiceTest")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' � la variabile i contatore per scorrere il form e fare update
  sint=Request.QueryString("sint")
  supersint=Request.QueryString("supersint")
   urlRisp=Request.QueryString("url") ' vale 1 se devo mostrare l'url delle pagine per il pdf


  
  '-----
 ' Data=Request.Form("txtDATA")
'  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
'  CodiceAllievo=Request.QueryString("CodiceAllievo")
'  ID_MOD=Request.QueryString("ID_MOD")
'  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
'  if left(Cartella,1)<>"" then ' DA SISTEMARE NELLE QUERY PER I GRUPPI !!!!!!!!!!!!!
'     Classe=clng(left(Request.QueryString("Cartella"),1))
'  end if
'
'

  '----
 FraseScelta=Request.QueryString("FraseScelta")
 'QuerySQL=Request.QueryString("QuerySQL")
 'QuerySQL=Request.Form("txtSQL")
  tutto=Request.QueryString("tutto")
if tutto = 1 then
   QuerySQL="SELECT M.TitPar,M.ID_Paragrafo,M.Cognome,M.Chi,M.CodiceFrase AS CF,M.Voto,M.Data,M.CodiceAllievo,M.Nome,M.Titolo,M.ID_Mod,M.Cartella,M.In_Quiz,M.Posizione,M.Expr1,M.Ora,M.Segnalata,M.Img,M.In_Umanet,M.Id_Prefrase,M.SotPar,P.ID_Prefrase,P.Id_Mod,P.Id_Paragrafo,P.CodiceFrase,P.Quesito,P.Eseguita,P.Posizione,P.Scadenza,P.Img,P.Files,P.Id_Sottoparagrafo,C.Titolo,C.Id_Classe,C.ID_Mod,C.Posizione,C.Paragrafo,C.ID_Paragrafo,C.Expr1 FROM MODULO_PARAGRAFO_FRASI1 as M,preFrasi as P,MODULI_PARAGRAFI_CLASSE as C Where M.Id_Prefrase = P.ID_Prefrase and P.Id_Paragrafo = C.ID_Paragrafo and M.ID_MOD='"&  Modulo &"' and CodiceAllievo='"&CodiceAllievo&"' and Data >= '"&Request.QueryString("DataCla")&"' and Data <= '"&Request.QueryString("DataCla2")&"' order by  C.Expr1,P.Id_Paragrafo,P.Posizione, M.CodiceFrase;"
   else
   QuerySQL="SELECT M.TitPar,M.ID_Paragrafo,M.Cognome,M.Chi,M.CodiceFrase AS CF,M.Voto,M.Data,M.CodiceAllievo,M.Nome,M.Titolo,M.ID_Mod,M.Cartella,M.In_Quiz,M.Posizione,M.Expr1,M.Ora,M.Segnalata,M.Img,M.In_Umanet,M.Id_Prefrase,M.SotPar,P.ID_Prefrase,P.Id_Mod,P.Id_Paragrafo,P.CodiceFrase,P.Quesito,P.Eseguita,P.Posizione,P.Scadenza,P.Img,P.Files,P.Id_Sottoparagrafo FROM MODULO_PARAGRAFO_FRASI1 as M,preFrasi as P Where M.Id_Prefrase = P.ID_Prefrase and M.ID_MOD='"&  Modulo &"' and M.ID_Paragrafo = '"&Paragrafo&"' and CodiceAllievo='"&CodiceAllievo&"' order by P.Id_Paragrafo, P.Posizione, M.CodiceFrase;"
   end if
 'QuerySQL=Request.QueryString("Query")
 'response.write("<br>1"&Request.Form("txtSQL") )
 'response.write("<br>2"&Request.QueryString("QuerySQL"))
'


'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\anno_2012-2013\logStampaFrasi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL &"---" & Modulo &"-----"& Cartella&"---")
'				objCreatedFile.Close
'
 'response.write("<br>3"&QuerySQL)
 %>

 <%
if (InStr(QuerySQL,"drop")=0) and (InStr(QuerySQL,"delete")=0) then
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
end if

%>


   <head>
<title><%=server.htmlencode(rsTabella(0))%></title>
<!--
<link rel="stylesheet" type="text/css" href="../stile.css">
-->
<link rel="stylesheet" type="text/css" href="custom2.css">
<style>
 
 /*body {
  margin: 0;
  padding: 0;
}*/
li {
margin-top: 0px;
}
 ol {
 /*margin: 0;
  padding: 0;*/
}

.sopra
{
margin-left:0cm; margin-right:0cm; margin-top:-10px
}
 li.MsoNormal
	{mso-style-parent:"";
	margin-bottom:.0001pt;
	font-size:14.0pt;
	font-family:"Times New Roman";
	margin-left:0cm; margin-right:0cm; margin-top:0cm}
<!--Il layout della pagina in fase di stampa sar� quello di un normale foglio A4 con margini di 3cm su tutti i lati.-->
@page {size: 210mm 297mm; margin: 30mm;}
<!--e vitare che l'interruzione avvenga nel corpo della tabella ->
table {page-break-inside: avoid;}

<!-- larghezza del div che contiene la tabella (facoltativo): il valore potrebbe essere omesso se vale 100% oppure se viene definito altrove -->
.table-responsive {width: 95%;}

<!-- stile del bordo per la tabella (facoltativo) -->
.table-responsive table {border: #ccc dotted 1px;}

<!--  istruzioni per le celle (alcune sono obbligatorie)-->
.table-responsive table td, .table-responsive table th
{min-width: 50px; width: 24%; border: #ccc dotted 1px; word-break: break-all; text-align: center; padding: 1%;}

<!--  larghezza delle immagini (facoltativo) -->
.table-responsive table td img {max-width: 50%;}


#content {
	float:center;
	width:95%;
	padding:2% 2% 2% 2%;
}

.flex {max-width: 100%}


</style>
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta charset="utf-8">

 <script type="text/javascript">
window.onload=function() {
//window.print();
}

function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")
 
 location.href="../../../../index.html";

 }
</script>


</head>

<%  if (strcomp(session("CodiceAllievo"),CodiceAllievo) <>0) and (session("admin")=false) then %>
	 <BODY onLoad="showText();"> </BODY>
  <% else %>
  <body bgcolor="#FFFFFF" style="font-family:Calibri, Candara, Segoe, 'Segoe UI', Optima, Arial, sans-serif">
  <%end if%>

<div id="container">
<div id="bloc_destra_cont" class="contenuti_login" style="width:95%">
<%


Set objFSO = CreateObject("Scripting.FileSystemObject")
%>
<b> <font size="+2"><center> Compiti di <%=rsTabella(2) & " " & rsTabella("Nome")%> </center></font></b>
<br><center><b>  <%=ucase(rsTabella("Titolo"))%></b></center>
<ul><li style="margin-top:-0px;"><b> <font size="+1"><%=rsTabella(0)%></font></b></li>


	<% capitolo=rsTabella("Titolo")
	   titoloParagrafo=rsTabella(0)

	   ' response.write("cbzero=" & Request.Form("cbzero") )
	i=1

	'response.write("eof="& rsTabella.EOF)
	Sottoparagrafo=""



	Do until rsTabella.EOF
	'response.write("eof="& rsTabella.EOF)
  	'if request.form("cb"&i).checked=true then
'	   response.write("Selezionato="&i)
'	end if
  if i>1 then
  	if strcomp(titoloParagrafo,rsTabella(0))<>0 then
	    titoloParagrafo=rsTabella(0)%>
<li><font size="+1"> <b><%=server.htmlencode(rsTabella(0))%></b></font></li>

	<%
	end if

	 if StrComp(Sottoparagrafo, rsTabella("SotPar")) <> 0 then
			  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
			   Sottoparagrafo=rsTabella("SotPar")
                %>
<b> <li><%=server.htmlencode(rsTabella("SotPar"))%> </li></b>
			 <%end if%>


 <% end if
 

 

   if clng(Request.Form("cb"&i)=0) and strcomp(tutto,"1")<>0 then ' se non lo devo stampare lo salto

        rsTabella.movenext()
   else

   ID=rsTabella(4)
  
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"


   url=Replace(url,"\","/")
   'response.write(url&"<br>")

      ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url2="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFrasi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.Close
'       response.write(url)
' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll







 '***2 KO
 '23/12/2019 DEvo poter distinguire il tipo di contenuto nel file in modo da scegliere se usare Server.HTMLEncode oppure no
 ' lo devo usare per le domande che contengono tag html 
 ' non lo devo usare per le risposte create con il nuovo sista di ckeditro che salvo html nei file di testo
 ' se non applico Server.HTMLEncode il testo si vede bene per le nuove risposte ckEditor ma i tag delle risposte su html vengono renderizzati e la stampa sballa
'*******DEvo applicare un controllo che se le risposte riguardano il capitolo HTML uso l'istruzione seguente
 'sReadAll = Server.HTMLEncode(sReadAll)


'response.write(sReadAll)
'sReadAll = ReplaceCar(sReadAll)
'sReadAll = url
sReadAll0=sReadAll
'sReadAll = replace(sReadAll,vbCr, "<br>") ' ***************** se commento questa riga il testo è più compatto(anche troppo) senza  spazi
'sReadAll = replace(sReadAll,vbLf, "<br>")
objTextFile.Close   ' la soluzione seguente la rimuovo e dir� di copiare ed incollare la domanda plus nella spiegazione
' cos� da avere il livello di apprendimento comprensibile , diversamente dovrei prevedere il modo di far apparire il testo della domanda plus
' anche nell'approfondimento di fine quiz.


%>


 <% if supersint="" then%>
<ul>
<li> <b><%=server.htmlencode(rsTabella(3))%> </b></li>
		<%if sint="" and supersint="" then%>
<span>


        <%'inserisco le eventuali immagini
		if rsTabella("Img")<>1 then%>
<%=sReadAll%>
    <%else%>
<%=sReadAll0%>

		 <%     QuerySQL1="Select * from Frasi_Img where Id_Frase="& rsTabella("CF")&";"
		 'response.write QuerySQL1
			   'url= "../Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Frasi/Img" ' vuole il percorso relativo della cartella

			    url="../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/Img"

			  ' url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/Img"

			   url=Replace(url,"\","/")


			   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
			   do while not rsTabella1.eof
			   'response.write(url&"/"& rsTabella("Url")&"<br>")
			  ' response.write("<br>x="& instr(urlimg,"tp://"))
			  if (instr(rsTabella1("Url"),"tp://")<>0) or (instr(rsTabella1("Url"),"tps://")<>0)then
			   urlimg= rsTabella1("Url") ' aggiungo al percorso il nome del file
			  else
			    urlimg=url&"/"& rsTabella1("Url") ' aggiungo al percorso il nome del file
			  end if

			   urldelete=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Frasi/Img/"&rsTabella1("Url")  ' per cancellare l'immagine.jpg
			   urldelete=Replace(urldelete,"\","/")
			   %>
			   

              <%  if ((instr(rsTabella1("Url"),"docs.google.com")<>0) or (instr(rsTabella1("Url"),"drive.google.com")<>0) or (instr(rsTabella1("Url"),"colab.research.google.com")<>0))  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente 
 gdoc="true"
		 response.write("<a href='"& rsTabella1("Url") &"' target='_blank'>apri url google drive</a>")
		  %>
	 <% else%>

		   <% if ((instr(rsTabella1("Url"),"https://www.mrwebmaster.it/img/guide/app-inventor/")<>0))  then
   ' è inutile tanto la width e height non li sente
		   %><br><center>
                   <img src="<%=urlimg%>" border="1" width="auto" height="auto"></center> <br>
			 <% else%>

	           <% if ((instr(rsTabella1("Url"),"altervista")<>0))  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
  if urlRisp<>"" then
		  response.write(rsTabella1("Url")&"<br>")
		 end if
		 response.write("<span style='margin-top:-100px'><a href='"& rsTabella1("Url") &"' target='_blank'>apri pagina web</a></span><br><br>")
		 
		  %>
	          <% else%>
<br><center>
		           <img class="flex" src="<%=urlimg%>" border="1" ></center> <br>
                <%end if%>

		 <%end if%>
        <%end if%>

			  <%rsTabella1.movenext
			   loop
		end if
	end if '	if sint="" and supersint=""
		%>
</span>
</ul>

   <%end if 'if supersint="" then%>



<%        rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
      end if ' di  if clng(Request.Form("cb"&i)=0) then  all'inizio
       i = i+ 1

    Loop

else
response.redirect "https://www.umanetexpo.net/"

end if
%>
</ul>



  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
   <script>

 $(document).ready(function(){
  //  window.print();
});
 </script>


</body>



</html>
