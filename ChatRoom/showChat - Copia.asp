<%@ Language=VBScript %>
<% 
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 %>
<!--#include file = "database.inc"-->
<!--#include file = "controllo_sessione.asp"-->
<!--#include file = "stringa_connessione.inc"--> 

<%
  divid=request.querystring("divid")
  cartella=request.querystring("cartella")
  id_classe=request.querystring("id_classe")
 
 
Function isBlank(Value)

if isNull(Value) then
	bAns = true
else
	bAns = trim(Value) = ""
end if
isBlank = bAns

end function

Function FixNull(Value)
if isNull(Value) then
	sAns = ""
else
	sAns = trim(Value)
end if

FixNull = sAns
end function


%>
<HTML>

<HEAD>

<TITLE>Chat  Umanet </TITLE>
<META https-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link rel="stylesheet" type="text/css" href="../../stile.css">
</HEAD>


<body onLoad="cambiaSessione();">

<div class="contenuti_forum">
<center>
<HR><center>
  <b><font size="+1"><img src="../forum/img/icon_star_blue.gif"> CHATROOM  <font color="#FF0000"><%=cartella%></font> <img src="../forum/img/icon_star_blue.gif"></font></b></center><p><bR>
<FORM ACTION = 'forum_search.asp'><b> Cerca nelle Chat : <img src="../forum/img/icon_aim.gif"></b>
    <input type="text" name="search" size="25">
 <input type="submit" value="Cerca" name="searchbutton" disabled="true">
</Form>

<P>


<%

Dim querySQL,rs  

iPageSize = 20
iPage = cint(Request.QueryString("Page"))
if iPage = 0 then iPage = 1


'sSQL = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"';"
'cmd.CommandText = sSQL
 'cmd.CommandText = "MESSAGETHREADS"
'cmd.CommandType = 4
'rs.open cmd, , 1, 3





sSQL = "select count(*) from CHAT_SESSION where Id_Classe='"&Id_Classe&"' ;"
cmd.CommandText = sSQL
set rs = cmd.Execute	
conn.execute sSQL
lTotalRecords=rs(0)

sSQL = "select * from CHAT_SESSION where Id_Classe='"&Id_Classe&"' order by Inizio desc  ;"
cmd.CommandText = sSQL
set rs = cmd.Execute	
conn.execute sSQL

'set rs = cmd.Execute


if not rs.Eof and not rs.bof then
'rs.MoveLast non supportto per le mie query
'lTotalRecords = rs.RecordCount
' calcola il numero di pagine necessarie in base al numero di post da mostrare
iTotalPages = int(lTotalRecords / iPageSize)
	if lTotalRecords MOD iPageSize <> 0 then iTotalPages = iTotalPages + 1
	' se basta una pagina
		if lTotalRecords <=  iPageSize then
			rs.MoveFirst
			bOnePage = true
			lPageEnd = lTotalRecords
			lPageStart = 1
			iTotalPages = 1
		else
			lPageStart = ((iPage - 1) * iPageSize) + 1
			lPageEnd = lPageStart + (iPageSize - 1)
		
		
			if lPageEnd >= lTotalRecords Then 
				lPageEnd = lTotalRecords
				bLastPage = true
			end if
			' posiziona il recordset in base alla pagina da visualizzare
			if iPage > 1 then
				rs.AbsolutePosition = ((iPage - 1) * iPageSize) + 1
			else
			' se ce una sola pagina va all'inizio
				rs.MoveFirst
			end if
		end if
	
	else
		bNoRecords = true
	
	end if
	



%>

</SELECT></TD></FORM>
<FORM onClick="PopUpWindow(409,481)" ACTION = "chatroom.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&cartella=<%=cartella%>" target="ChatWindow2" METHOD = "GET" >

 
<TD>
<%
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") &"'"
 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
  
 if rsTabella1("ChatAbilitata")=0 then %>
<INPUT disabled="disabled" TYPE = "SUBMIT" VALUE = "Inizia nuova Chat"></TD></FORM>
<%else%>
<INPUT  TYPE = "SUBMIT" VALUE = "Inizia nuova Chat"></TD></FORM>
<%end if
rsTabella1.close
set rsTabella1=nothing
connessioneDB.close
set connessioneDB=nothing
%>

     </tr>



</TABLE><P>
<%
if not bNoRecords then
    response.write "<P><B>Pagina " & iPage & " di " & iTotalPages & "</B><P>"
end if
%>
<TABLE WIDTH = 100%  id="zebra" border=1 align="center" bordercolor=pink>
<thead>
<TR>
<Th><B><FONT COLOR = "RED">Titolo</FONT></B></Th>
<Th><B><FONT COLOR = "RED">Inizio</FONT></B></Th>
<Th ALIGN = CENTER><B><FONT COLOR = "RED">Fine</FONT></B></Th>
 
<%
if Session("Admin")=true then%>
<Th><B><FONT COLOR = "RED">Elimina</FONT></B></Th></TR></thead>
<%
else%>
</TR></thead>
<%end if 
if bNoRecords then
 response.write "<TD COLSPAN = 4><B>Non ci sono chat nello storico</B></TD>"

else
 for lCtr = lPageStart to lPageEnd
 if (lCtr mod 2) = 0  then 
	    classe_riga="zebra-dispari"
	else
	    classe_riga=""
end if	
 response.write "<tr class="&classe_riga&"> <TD><A HREF='ShowChat2.asp?ID_Chat=" & rs("ID_Chat") &"'>"  & rs("Titolo") & "</A></FONT></TD>"
 response.write " <TD>" 
 if session("Admin")=true then ' se sono admin visualizzo il codice autore post
   response.write "<A title='" & rs("ID_Chat") &"' HREF = '#'>" & rs("Inizio") & "</A>" 
 else
 response.write "<A HREF = '#'>" & rs("Inizio") & "</A>" 
 end if

response.write "</FONT></TD>"

response.write "</TD><TD ALIGN = CENTER>" & rs("Fine") & "</FONT></TD>"
'response.write "</TD><TD>" & rs("Fine") & "</FONT></TD>"
if Session("Admin")=true then
ID_Chat=rs("ID_Chat")
%>

 
<TD align=center><A onClick="return window.confirm('Vuoi veramente cancellare la Chat?');" HREF="cancella_chat.asp?ID_Chat=<%=ID_Chat%>&nome=<%=rs("Nome")%>"> X</a></TD></TR> 
<%
else%>
</TR> 
<%end if 

 rs.movenext
 Next
end if
response.write "</TABLE>"

if bOnePage = false and bNoRecords = false then

response.write "<TABLE WIDTH = '100%'><TR><TD>&nbsp;</TD></TR><TR><TD WIDTH = '10%'>&nbsp;</TD><TD WIDTH = '60%'>"
 
if iPage > 1 then
sPrevQuery = "Page=" & iPage - 1
response.write "<A HREF = 'default.asp?" & sPrevQuery & "'><B><< Previous Page</B></A>"
  else
response.write "&nbsp;"
end if
		
response.write "</TD><TD VALIGN = TOP NOWRAP>"

if bLastPage = false then
		
sNextQuery = "Page=" & iPage + 1 
response.write "<A HREF = 'default.asp?" & sNextQuery & "'><B>Next Page >></B></A>"
else
response.write "&nbsp;"
end if
response.write "<TD WIDTH = '10%'>&nbsp;</TD>"
response.write "</TD></TR></TABLE>"
response.write "<P><CENTER><FONT SIZE =-1>"

for iCtr = 1 to iTotalPages
sPageQuery = "Page=" & iCtr & sQuery
if iCtr <> iPage then
 response.write "<A HREF = 'ShowChat.asp?" & sPageQuery & "'>"
else

 response.write "<B>"
end if
response.write iCtr

if iCtr <> iPage then
response.write "</A>"
else
response.write "</B>" 
end if
if iCtr < iTotalPages then response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"
	

Next
response.write "</FONT></CENTER>"
end if
%><center><br><br>
 <!--#include file = "footer.inc"--></center>
</div>
 
</body>

<!--#include file = "database_cleanup.inc"-->


</HTML>


