<%@ Language=VBScript %>

  <% Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome


   'Apertura della connessione al database
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	ID=request.querystring("ID")
	revocaPS=request.querystring("revocaPS")
     revoca=request.querystring("revoca")
	  revocatutti=request.querystring("revocatutti")
	CodiceAllievo=request.querystring("CodiceAllievo")
	CodiceAllievoPost=request.querystring("CodiceAllievoPost")
  Zip=request.querystring("Zip")
	iThreadParent=request.querystring("IDPARENT")
	'MaxStelline=3
	'response.write("aa="& request.querystring("MaxStelline"))
	if request.querystring("MaxStelline")="" then
	    MaxStelline=3
	else
	  MaxStelline=request.querystring("MaxStelline")
	end if
	MaxStelline=cint(MaxStelline)
	scegli=request.querystring("scegli")
	Topic=request.querystring("Topic")
	categoria=request.querystring("categoria")
	id_categoria=request.querystring("id_categoria")

	function ReplaceComments(sInput)
dim sAns
'sAns = replace(sInput, "  ", "&nbsp; ")
'if inStr(sValue,"www.youtube.com")=0 then
   sAns = replace(sInput, chr(34), "")
'end if
sAns = replace(sAns, "<!--", "&lt;!--")
sAns = replace(sAns, "-->", "--&gt;")

ReplaceComments = sAns
end function

       %>
<html>
<head>
	<link rel="stylesheet" type="text/css" href="../../stile.css">

    <script language="javascript" type="text/javascript">
function showText() {window.alert("Non puoi votare per te stesso!")

location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%=Zip%>"
//location.href=window.history.back();
 }
 </script>
</head>

<% Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   	<!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->
	<!--#include file = "../service/controllo_sessione.asp"-->

<%

scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="lavagna"
  case "2"
    session("social")="diario"
    case "3"
      session("social")="interrogazioni"

 end select
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
if revocatutti<>"" then

      QuerySQL="Delete from Voti WHERE  ThreadParent="&ID&";"

	' response.write(QuerySQL)
	  conn.Execute(QuerySQL)%>

	  <script language="javascript" type="text/javascript">
 window.alert("Tutte le votazioni sono state resettate")
 location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%=Zip%>"
//location.href=window.history.back();

 </script>

<%
'Response.Redirect("javascript:history.go(-1)")
'Response.Redirect(Request.UrlReferrer.AbsolutePath )

else
    if revocaPS<>"" then
	   QuerySQL="Update FORUM_MESSAGES Set Punti=0 WHERE  ThreadParent="&ID&";"
	   conn.Execute(QuerySQL)%>

	  <script language="javascript" type="text/javascript">
 window.alert("Tutti i PS sono stati annullati")
 location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%=Zip%>"
	</script>
	<%end if

if (strcomp(ucase(CodiceAllievo),ucase(CodiceAllievoPost))<>0)  then  %>
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">

<%

'messaggio=Replace(messaggio, Chr(39), "''")

 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	VotoPalese=rsTabella("VotoPalese")
if revoca="" then

 QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=-1;"
 set rs=conn.Execute(QuerySQL)
QuerySQL="select count(*) from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&";"
'AppQuery=QuerySQL
		' dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logVota.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
''
			' response.write(QuerySQL &"<br>")
		 set rs=conn.Execute(QuerySQL)
		 numVoti=rs(0)


		 if numVoti<MaxStelline then
					voto=1
					 QuerySQL="INSERT INTO Voti (CodiceAllievo,ThreadParent,ThreadQuote,Data,Voto,Cognome,Nome) SELECT '" & CodiceAllievo & "','" & iThreadParent & "','" & ID & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "';"

				 'response.write(QuerySQL)

						conn.Execute(QuerySQL)
				' inserisco solo se non � un voto palese
			 'Testo="Ha risposto ad un tuo post !"
		if VotoPalese=1 then

			 Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
			 Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

			 QuerySQL="select * from AVVISI where IdPost=" & ID & " and CodiceAllievo='"&CodiceAllievoPost &"' and CodiceAllievo2='"& Session("CodiceAllievo")&"' and Social="&scegli&";"
			 set rs=ConnessioneDB.Execute(QuerySQL)
			 ' metto solo una notifica anche se ci sono pi� voti, evito di intasare le notifiche quando la classe vota il singolo
			 if rs.eof and rs.bof then
			 QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore,Social,IdPost) SELECT '" & CodiceAllievoPost & "','" & ReplaceComments(Topic) &"','"& Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "','" & scegli & "'," & ID &";"
	         end if



			   ConnessioneDB.Execute(QuerySQL)
			 '  response.write(QuerySQL)

			    sSQL="select max(ID_Avviso) from AVVISI;"
	  set rs=ConnessioneDB.Execute(sSQL)
		 maxIDAvviso=rs(0)
	   	rs.close

		  Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&maxIDAvviso="&maxIDAvviso&"&Zip="&Zip&">Ho quotato un tuo post !</a>"

		 sSQL="Update AVVISI set Azione ='"& Azione&"' where ID_Avviso="&maxIDAvviso&";"
		 ConnessioneDB.Execute(sSQL)
		   end if ' votopalese
		  %>


			 <script language="javascript" type="text/javascript">
		 window.alert("Mi piace applicato! Voto assegnato <%=6 + numVoti%>")

		location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%Zip%>"
		//location.href=window.history.back();

		 </script>
			  <%
	  else
	       %>

                  <script language="javascript" type="text/javascript">
		 window.alert("Hai già utilizzato tutte le stelline!")

		location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%Zip%>"
		//location.href=window.history.back();

		 </script>
	  <%


	  end if

else  ' revoco eventuale mi piace
       QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=1;"
	  conn.Execute(QuerySQL)
	 QuerySQL="select count(*) from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&";"

		 set rs=conn.Execute(QuerySQL)
		 numVoti=rs(0)

		  'if numVoti<3 then

	  QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=1;"
	  conn.Execute(QuerySQL)


	  ' se non ho utilizzato i MAx voti disponibili aggingo voto=-1
	  if numVoti<MaxStelline then
	  voto=-1

			  QuerySQL="INSERT INTO Voti (CodiceAllievo,ThreadParent,ThreadQuote,Data,Voto,Cognome,Nome) SELECT '" & CodiceAllievo & "','" & iThreadParent & "','" & ID & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "';"
				conn.Execute(QuerySQL)%>
	          <script language="javascript" type="text/javascript">
 window.alert("Non mi piace applicato! Voto assegnato <%=5-numVoti%>")
 			location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>&Zip=<%Zip%>"
//location.href=window.history.back();

 </script>
       <%else%>
	       <script language="javascript" type="text/javascript">
 window.alert("Hai gi� utilizzato tutte le stelline!")
location.href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>&Zip=<%Zip%>"
//location.href=window.history.back();

 </script>
	   <%end if%>


	  <%


end if





'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA


'response.write(url)

		On Error Resume Next
		If Err.Number = 0 Then
				Response.Write "Voto avvenuto! "
				'Response.Redirect "ShowMessage.asp?ID="&ID
		Else
				Response.Write Err.Description
				Err.Number = 0
		End If
%>
	<center><br><br><font size="3">
<!--#include file = "footer.inc"-->
</center>
<!--#include file = "database_cleanup.inc"-->
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>

   <BODY onLoad="showText();">

	<%end if%>

<%end if ' inziale revocatutti	%>
	</body>
	</html>
