

<%@Language="vbscript"%>
 
<%if Session("CodiceAllievo")="" or Session("Id_Classe")=""then %> 
 <script language="javascript" type="text/javascript"> 
    window.alert("La sessione è scaduta, effettua nuovamente il Login!");
     location.href="../../../home.asp";
//location.href=window.history.back();

</script>
 
 <% end if%>
 
 
  <!--#include file="upload.asp"-->
  <!-- #include file = "../../cDomande/tabella_corrispondenze.inc" -->
<%  
	                        'Lettura dei dati memorizzati nei cookie. 
  ' CodiceTest = Request.Cookies("Dati")("CodiceTest")
    Function controlla(RisposteEsatte)
	 controlla=0
	 i=0
	 while (i<=13) and not(esiste)
	 'response.Write(v2(i) & "=" & RisposteEsatte & "<br>")
		if strcomp(v2(i),RisposteEsatte)= 0 then 
		    esiste=true
		    controlla=1
			'response.Write(" Trovato <br>")
		end if
		i=i+1
	 wend
 end function
    

Dim nome ,Conta,AggRisPar
  


  ' serve per sapere se sono stato chiamato da modificamodulo.asp per caricare le risorse dei paragrafi
 AggRisPar=Request.QueryString("AggRisPar")
 ' serve per sapere se sono stato chiamato da modificamodulo.asp per caricare la risorsa del modulo
 AggRisMod=Request.QueryString("AggRisMod")
 Classe=Request.QueryString("Classe")
 Id_Mod=Request.QueryString("Id_Mod")
 divid=Request.QueryString("divid")
 Id_Classe=Request.QueryString("Id_Classe")
 
 AggImgForum=Request.QueryString("AggImgForum") ' se sono chiamato per aggiungere immagine ad un post
 AggRisFrase=Request.QueryString("AggRisFrase")
 AggRisDomanda=Request.QueryString("AggRisDomanda") ' se sono chiamato da insersci_test
 by_UPLOAD=Request.QueryString("by_UPLOAD")
 ID=Request.QueryString("ID")
 Img=Request.QueryString("ID") 
 contDomande=Request.QueryString("contDomande") ' incremento per il nome delle immagini multiple per la stessa frase
  
 'var="ContDOmande="&contDomande&" by_UPLOAD="&by_UPLOAD&" ID="&ID 
'  dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013_ma\logUpdate.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(var)
'				objCreatedFile.Close
 
 
 if contDomande="" then
    contDomande=1
 else 
    contDomande=cint(contDomande)
	contDomande=contDomande+1
 end if	
 
' response.write("AggRisFrase="&AggRisFrase)
 'se sono stato chiamato da inserisci frase  per caricare immagine

 

if AggRisFrase<>"" then


   CodiceAllievo=Request.QueryString("CodiceAllievo")
   Quesito=Request.QueryString("Quesito")
   Chi=Quesito ' da non togliere serve per il file di include
   CodiceTest = Request.QueryString("CodiceTest")
   Cartella=Request.QueryString("Cartella")
   preFrase=Request.QueryString("prefrase") ' serve per capire il chiamante e quindi sapere se alla fine 
   ID_Prefrase=Request.QueryString("ID_Prefrase") ' serve per controllare se è già stata inserita
   by_UECDL=Request.QueryString("by_UECDL")
   Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   Modulo=Request.QueryString("Modulo")
   AggImg=Request.QueryString("AggImg") ' vale 1 se sono chiamata da inserisci_valutazione, cioè aggiunta di immagine
   ' serve per distinguire dove fare il redirect dopo l'inserimento
 
end if
  
  if AggRisDomanda<>"" then
   'Quesito=Request.QueryString("Quesito")
   In_Quiz_Stud=Request.QueryString("In_Quiz_Stud")
   Tipo=Request.QueryString("Tipo")
   CodiceAllievo=Request.QueryString("CodiceAllievo")
   Quesito=Request.QueryString("Quesito")
   Chi=Quesito ' da non togliere serve per il file di include
   CodiceTest = Request.QueryString("CodiceTest")
   Cartella=Request.QueryString("Cartella")
   Multiple=Request.QueryString("Multiple")
   by_UECDL=Request.QueryString("by_UECDL")
   Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   Modulo=Request.QueryString("Modulo")
   Img=Request.QueryString("Img")
    ID=Request.QueryString("Id_Domanda")
  ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'					url1="C:\Inetpub\umanetroot\anno_2012-2013\logMultiple0.txt"
'					Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'					objCreatedFile.WriteLine(Multiple & " --" & Img)
'					objCreatedFile.Close	

 '  AggImg=Request.QueryString("AggImg") ' vale 1 se sono chiamata da inserisci_valutazione, cioè aggiunta di immagine
   ' serve per distinguire dove fare il redirect dopo l'inserimento
'response.write("<br>ID_Prefrase="& ID_Prefrase)
'response.write("<br>Quesito="&Quesito )
'response.write("<br>CodiceTest="&CodiceTest )
'response.write("<br>Cartella="&Cartella )
'response.write("<br>prefrase="&prefrase )
'response.write("<br>by_UECDL="& by_UECDL)
'response.write("<br>Capitolo="& Capitolo)
'response.write("<br>Paragrafo="&Paragrafo )
'response.write("<br>Modulo="&Modulo )
end if
 

				
   
 Dim ConnessioneDB, rsTabella, QuerySQL  ,QuerySQL1
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   
  
   
 %>

 <!-- #include file = "../../stringhe_connessione/stringa_connessione.inc" -->
  
<html>
<!-- #BeginTemplate "/Templates/homepage.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Home - Website Products</title>
<!-- #EndEditable -->
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<link rel="stylesheet" href="/style.css" type="text/css">
</head>

<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top"> 
      <!--include file="includes/product_header.htm" -->
    </td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="5" align="left" valign="top">&nbsp;</td>
    <td width="385" align="left" valign="middle"></td>
    <td width="385" align="right" valign="middle"><span class=date><font color="#999999"><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
      </font></font></font></span></td>
    <td width="5" align="left" valign="top">&nbsp;</td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="5" align="left" valign="top">&nbsp;</td>
    <td width="150" align="left" valign="top"> 
      <!--include virtual="/includes/home_left_new.htm" -->
    </td>
    <td width="470" align="left" valign="top"><!-- #BeginEditable "content" --> 

<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td valign="top">
		<!--#include file="update.asp"-->
      	<%
		'Id_Par=dicUpload.Item("txtId_Par").Item("Value")
 ' response.write("IDPART="&Id_Par)
		%>
        <br>Please check the site to see whether the changes are reflected online.
<%
Else
	response.write "<br><font face='verdana' color='red' size='2'>Il file non può essere caricato, o non hai selezionato alcun file.</font>"
	%>
	<br><a href="javascript:history.back()">Indietro</a>
	<%
end if%>
	
     </td>
  </tr>
</table>
      <!-- #EndEditable --></td>
    <td width="100" align="right" valign="top">&nbsp; </td>
    <td width="5" align="right" valign="top">&nbsp; </td>

  </tr>
  <tr> 
    <td colspan="5" align="center" valign="middle"> 
      <!--include virtual="/includes/footer_new.htm" -->
    </td>
  </tr>
</table>
</body>
<!-- #EndTemplate --></html>
