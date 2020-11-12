<%@ Language=VBScript %>
<% Response.Buffer=True
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
ID=Request.QueryString("CodiceFrase")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
Cartella=Request.QueryString("Cartella")
CodiceAllievo=Request.QueryString("CodiceAllievo")
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Replace(url,"\","/")
on error resume next
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then
				QuerySQL ="DELETE  FROM FRASI WHERE CodiceFrase =" &ID&";"
				ConnessioneDB.Execute(QuerySQL)
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				objFSO.DeleteFile url
				If Err.Number = 0 Then
					Response.Write "Cancellazione avvenuta!"
				Else
					Response.Write Err.Description
					Err.Number = 0
				End If
else
				response.write("Non puoi cancellare i dati degli altri studenti!")
end if
%>
