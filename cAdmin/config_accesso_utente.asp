<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
Dim ConnessioneDB, rsTabella, QuerySQL,Privato,Valutato,Classe   
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
  Id_Classe = Request.QueryString("Id_Classe")
  Id_Classe_byRunner = Request.form("txtId_Classe_byRunner")
 ' Privato = Request.Form("TxtPrivato")  
  
 ' Valutato = Request.Form("TxtValutato") 
'  TestAbilitato= Request.Form("TxtTestAbilitato")
'  ChatAbilitata= Request.Form("TxtChatAbilitata")
'  CIAbilitato= Request.Form("TxtCIAbilitato")
'  DVAbilitato= Request.Form("TxtDVAbilitato")
'  ScalaValutaz= Request.Form("TxtVVMax")
'  Runner0=Request.Form("TxtLepre")
'  DataVal=Request.Form("TxtDataVal")
'  cancella=Request.QueryString("cancella")
'  idperiodo=Request.QueryString("idperiodo")
'  divid=Request.QueryString("divid")
'  Runner=Request.Form("txtRunner")
'  idperiododb=Request.Form("txtId_Periodo_byRunner")
'  VotoPalese=Request.Form("TxtVotoPalese")
'  MaxStelline=Request.Form("txtMaxStelline")
  
   Privato = Request.Form("CheckPrivato") 
   Valutato = Request.Form("CheckValutato") 
  TestAbilitato= Request.Form("CheckAbilitato")
  ChatAbilitata= Request.Form("CheckChatAbilitata")
  CIAbilitato= Request.Form("CheckCIAbilitato")
  DVAbilitato= Request.Form("CheckDVAbilitato")
  JSAbilitato= Request.Form("CheckJSAbilitato")
  ScalaValutaz= Request.Form("TxtVVMax")
  Runner0=Request.Form("CheckLepre")
  DataVal=Request.Form("TxtDataVal")
  cancella=Request.QueryString("cancella")
  idperiodo=Request.QueryString("idperiodo")
  Runner=Request.Form("CheckRunner")
  idperiododb=Request.Form("CheckId_Periodo_byRunner")
  VotoPalese=Request.Form("CheckVotoPalese")
  MaxStelline=Request.Form("txtMaxStelline")
  VotoAttivo=Request.Form("CheckVotoAttivo")
  Registrazione=Request.Form("CheckReg")
  ValidaQuiz=Request.Form("CheckValidaQuiz")
  Nodi=Request.Form("CheckNodi")
  Recupero=Request.Form("CheckRecupero")
 
  
 ' on error resume next
  if Runner=0 then
 ' data=Request.Form("data")
'QuerySQL="UPDATE Setting SET Privato = "& cint(Privato) &",Valutato="&cint(Valutato)& ",TestAbilitato="&cint(TestAbilitato)& ",ChatAbilitata="&cint(ChatAbilitata)& ",CIAbilitato="&cint(CIAbilitato)& ",DVAbilitato="&cint(DVAbilitato)&  ",ScalaValutaz="&cint(ScalaValutaz)&  ",Runner="&cint(Runner0)&  ",VotoPalese="&cint(VotoPalese)&  ",MaxStelline="&cint(MaxStelline)& ",VotoAttivo="&cint(VotoAttivo)& ",Registra="&cint(Registrazione)&  " WHERE Id_Classe='" & Id_Classe & "';"

QuerySQL="UPDATE Setting SET Privato = "& cint(Privato) &",Valutato="&cint(Valutato)& ",TestAbilitato="&cint(TestAbilitato)& ",ChatAbilitata="&cint(ChatAbilitata)& ",CIAbilitato="&cint(CIAbilitato)& ",DVAbilitato="&cint(DVAbilitato)&  ",ScalaValutaz="&cint(ScalaValutaz)&  ",Runner="&cint(Runner0)&  ",VotoPalese="&cint(VotoPalese)&  ",MaxStelline="&cint(MaxStelline)&  ",ValidaQuiz="&cint(ValidaQuiz)&",Registra="&cint(Registrazione)&",Nodi="&cint(Nodi)& ",JSAbilitato="&cint(JSAbilitato)& ",RecuperoAttivo="&cint(Recupero)&" WHERE Id_Classe='" & Id_Classe & "';"
 response.write(QuerySQL&"<br>")
	ConnessioneDB.execute(QuerySQL)  
	
	
	' non si capisce perchè da errore se aggiungo questi due campi,stesso bug per le notifiche di home_uecdl_app
	'previsti parametri ??? non riconosce il nome del campo
	'QuerySQL="UPDATE Setting SET VotoAttivo = "& cint(VotoAttivo)& ",Registra="&cint(Registrazione)&   " WHERE Id_Classe='" & Id_Classe & "';"

    'response.write(QuerySQL&"<br>")
	'ConnessioneDB.execute(QuerySQL)  

	
	
	
	
 
	 response.Redirect "admin.asp?Id_Classe="&Id_Classe  
 
 else 
 	response.write(QuerySQL&"<br>")
 QuerySQL="UPDATE Setting SET Runner = "& cint(Runner) &" WHERE Id_Classe='" & Id_Classe_byRunner & "';"
	ConnessioneDB.execute(QuerySQL) 

	QuerySQL="UPDATE [dbo].[3PERIODI] SET Runner = "& cint(Runner) &" WHERE Id_Classe='" & Id_Classe_byRunner & "' and ID_Periodo=" & cint(idperiododb) &";"
	'ConnessioneDB.execute(QuerySQL) 
	response.write(QuerySQL&"<br>")
	
	
 
        if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
		' response.write(QuerySQL)
 
 end if
 
 
%>
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Accesso utente</title>
</head>

<body>
</body>
</html>
