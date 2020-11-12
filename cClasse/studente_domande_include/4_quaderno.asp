<%


'Response.AddHeader "Refresh", "600"

 ' Cartella=Request.QueryString("Cartella")
  Cartella=Request.QueryString("classe")
 ' response.Cookies("Dati")("Cartella")=Cartella
 ' TitoloCapitolo=Request.QueryString("Capitolo") 
 ' Paragrafo=Request.QueryString("Paragrafo")
  'Modulo=Request.QueryString("Modulo")
 ' CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Session("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")  
  dividA=request.QueryString("dividApro")
    On Error Resume Next
xEstrazione=request.querystring("xEstrazione")
id_classe=request.querystring("id_classe")
  Response.Cookies("Dati")("Id_Classe")=id_classe
classe=request.querystring("classe")
divid=request.querystring("divid")
 
PS=request.querystring("PS") ' vale 1 se devo mostrare anche i Punti Social chiamato da javasscript
if PS="" then ' per la prima chiamata mostrio i PS
   PS=1
end if
 
daStud=Request.QueryString("daStud") ' chiamato da href della classifica
daForm=Request.QueryString("daForm") ' chiamato dal bottone invia di scelta periodi dal al
daMenu=Request.QueryString("daMenu")
daMenuMappe=Request.QueryString("daMenuMappe")

DataCla=request.form("txtData") 
DataCla2=request.form("txtData2")
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")


if daForm<>"" then
 ' Session("DataClaq")=DataClaq
 ' Session("DataClaq2")=DataClaq2
  
end if
if daStud<>"" then
  
end if

if DataCla="" then
   if DataClaq2<>"" then
      DataCla=DataClaq
	  DataCla2=DataClaq2
   else
     DataCla=Session("DataCla")
	  DataClaq=Session("DataClaq")
	 DataClaq2=Session("DataClaq2")
	end if 
end if

'if daMenu<>"" then
'    DataCla=request.QueryString("DataClaq") 
'    DataCla2=request.QueryString("DataClaq2")
'end if
'if daStud<>"" then
'   'DataClaq= DataCla
'   'DataClaq2=DataCla2
'    DataClaq=request.QueryString("DataClaq") 
'	DataClaq2=request.QueryString("DataClaq2")
'   
'end if

'response.write(DataClaq & "<br>" & DataClaq2)
'if session("DataClaq")="" then
'Session("DataClaq")=DataClaq
'Session("DataClaq2")=DataClaq2
'else
' DataClaq=Session("DataClaq")
' DataClaq2=Session("DataClaq2")
' DataCla=Session("DataClaq")
' DataClaq=Session("DataClaq2")
' end if
'' response.write("dopo session OK "& DataClaq & "<br>" & DataClaq2) 
'' se è la prima chiamata il valore del form sopra la classifica è nullo
'if (DataCla<>"") and (DataCla2<>"") then
'	Session("DataCla")=DataCla
'	Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
'else
'   Session("DataCla")= Session("DataClaq")
'   Session("DataCla2")= Session("DataClaq2")
'end if
'  
  
  
  
  cod=Request.QueryString("cod")
  if strcomp(cod&"","")=0 then
     cod=Session("CodiceAllievo")
	
	 
  end if
  
 box_apri="toggleCapitolo"&request.querystring("tCap")
 box_apri1="toggleSottoPar"&request.querystring("tSot")
 box_apri2="toggleDomande"&request.querystring("tDom")
 box_apri3="toggleFrasi"&request.querystring("tFra")
 box_apri4="toggleNodi"&request.querystring("tNod")
 
  
  
  
  
function ReplaceCar(sInput)
dim sAns
 
  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
  sAns=  Replace(sAns,"'",Chr(96))
  
ReplaceCar = sAns
end function

   
  
  
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario
 
		%> 