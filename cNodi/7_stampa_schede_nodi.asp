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

 }
 </script>
  <script type="text/javascript">
window.onload=function() {
window.print();
}
</script>
<TITLE>NODI DELLA RETE</TITLE>
</HEAD>
<%Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>

 <%   

    Dim ConnessioneDB, rsTabella,rsLink,QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato
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
  Codice_Test=Request.QueryString("CodiceTest") 
  'response.write("Codice_Test:"& Codice_Test)

 QuerySQL=Request.QueryString("QuerySQL")  
 
  
' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\anno_2012-2013\logSpiegazioneNodi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine("-"& ucase(CodiceAllievo)& "-" & ucase(session("CodiceAllievo")))
'				objCreatedFile.Close 

'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
'  
 
  Dim objFSO, objTextFile
  Dim liv(8) ' serve per indicizzare il chi,cosa,....
  liv(1)="Chi"
  liv(2)="Cosa"
  liv(3)="Dove"
  liv(4)="Quando"
  liv(5)="Come"
  liv(6)="Perchè"
  liv(7)="Quindi"
  
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  						
%>
<body bgcolor="#FFFFFF">
<div id="container">  
 <!--<div id="bloc_destra_cont" class="contenuti_login" style="font-style:normal;"> -->
<div id="bloc_destra_cont"  style="font-style:normal;">
 

  <table border="0" align=center width="60%"  id="zebra_stud">
		<tr>
			<td colspan=3 align=center>
			  <font color="#000000"><b><h3><%=Capitolo%></h3></b></font>
			</td>
		</tr>
	
	</table>
 
 
 
	<br>
<%   
  
 

 
 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 
      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%><center>
  <H4>Nodi della rete non ancora disponibili!</h4></center>
  
<% Else
  
	  i=1 'inizializza la variabile i (contatore delle domande)
	  Do until rsTabella.EOF
	  'response.Write(rsTabella(12))
		if (strcomp(rsTabella(12),"12/12/2112")<>0) then  'apro l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
					 
				 
					ID=rsTabella(3)
					url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&rsTabella("CodiceNodo")&".txt"
					 ' NB c'è una / nell'url locale
				
					' url=Server.MapPath("/ECDL") & "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
					   url1= "../" & Cartella & "/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&rsTabella("CodiceNodo")&".txt"
				
				url=Replace(url,"\","/")
				 
				'response.write(url)
				' Open file for reading.
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
				'sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
				%>
				<%' devo controllare se ID nodo esiste nella tabella dei link in tal caso leggo la L1 ed in quella posizione invece dell'ancora metto href
										  '0		   1		 2			3		4			5          6
				QuerySql="Select Link.ID_Link, Link.Id_n1, Link.L1, Link.Id_n2, Link.L2, Link.Id_Stud,Link.Testo2 FROM Link WHERE Id_n1="&ID&";"
				 
			
				Set rsLink = ConnessioneDB.Execute(QuerySQL)
				If rsLink.BOF=True And rsLink.EOF=True Then  ' se il nodo non compare nella tabella link allora metto tutte ancore
				%>
			
					  <table border="1"  align=center width="60%"  id="zebra_stud">
							<tr>
							  <td width="10%"><b>Nodo n</b>.<%=rsTabella("CodiceNodo")%></td>
							  <td width="18%"><%=rsTabella.fields("Titolo")%></td>
							  <td width="69%"><%=rsTabella.fields("Cognome")%></td>
							</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
							<tr><td><b><a name="<%=ID%>_1">Chi</a></b></td><th colspan=3><p align="center"><b><%=rsTabella.fields("Chi")%></b></th></tr>
							<tr><td><b><a name="<%=ID%>_2">Cosa</a></b></td><td colspan=2><p align="center"><%=rsTabella.fields("Cosa")%></td></tr>
							<tr><td><b><a name="<%=ID%>_3">Dove</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Dove")%></td></tr>
							<tr><td><b><a name="<%=ID%>_4">Quando</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Quando")%></td></tr>
							<tr><td><b><a name="<%=ID%>_5">Come</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Come")%></td></tr>
							<tr><td><b><a name="<%=ID%>_6">Perchè</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Perche")%></td></tr>
							<tr><td><b><a name="<%=ID%>_7">Quindi</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Quindi")%></td></tr>
							<tr>
							<td colspan=3>
							<p align="center">
							 <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" cols="80"><% 
							 Response.write(sReadAll)%> </textarea><br>
							</td>
							</tr>
				</table>
				<br>
				<%else ' devo mettere href nel livello indicato %> 
					
					
					<table border="1"  align=center width="62%"  id="zebra_stud">
							<tr>
								<td width="10%"><b>Nodo n</b>.<%=rsTabella.fields("CodiceNodo")%></td>
							  <td width="16%"><%=rsTabella.fields("Titolo")%></td>
							  <td width="60%"><%=rsTabella.fields("Cognome")%></td>
								<td width="14%">Link to   </td>
							</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
							
							<%' per ogni livello di ogni nodo vedo i link che ha ad altri nodi, e metto una stellina per ognuno
							  ' per ogni livello controllo il rsLink, se trovo che il livello è coinvolto in un link metto href, la prima volta metto il <td> le altre aggiungo allo stesso <td>
							   for i=1 to 7
								primo=0
								primo1=0 %>
							   <tr>
							   <td><b><a name="<%=ID%>_<%=i%>" title="<%=ID%>_<%=i%>"><%=liv(i)%></a></b></td><td colspan=2><p align="center"><%=rsTabella(4+i)%> </td>
											
								<%	 rsLink.Movefirst()
									 Do until rsLink.EOF
											L1=rsLink(2)
											Id_n1=rsLink(1)
											Id_n2=rsLink(3)
											L2=rsLink(4)
											T2=rsLink(6)
										   if i=L1 then
												 if primo=0 then 
													primo=1 %>
													<td><a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>">></a>
												<%else%>
													 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"><</a>
												<%end if%>  
										   <%end if  
										  rsLink.Movenext()
										Loop%>
								</td></tr>
							  <% next
								
							 %>
							 
							<tr>
							<td colspan=4>
							<p align="center">
							 <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" cols="80"><% 
							 Response.write(sReadAll)%> </textarea><br>
							</td>
							</tr>
				</table>
				<br>	
				
				<%end if %>
			<%
			
       i = i+ 1 
	   	end if  'chiudo l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
	
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
 
  </div>
  </div>
  </div>
</BODY>
<% 'else 
  'Response.Redirect "../home.asp"
    '  end if %>
</HTML>