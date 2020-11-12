<%
'Sample file Field-SaveAs.asp
'Store extra upload info to a database
' and file contents to the disk
Server.ScriptTimeout = 5000
	 session("CaricatoFile")=false
     session("NomeFileForum")=""
	 daForum=request.QueryString("daForum")
	 daDiario=request.QueryString("daDiario")
	 daFrase=request.QueryString("daFrase")

dim ConnessioneDB: Set ConnessioneDB = CreateObject("ADODB.Connection")
'Create upload form
'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm %>

<!--#INCLUDE FILE="_upload.asp"-->
<!--#INCLUDE FILE="../var_globali.inc"-->
<!--#INCLUDE FILE="../stringhe_connessione/stringa_connessione.inc"-->
<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>


<%


Server.ScriptTimeout = 1000
Form.SizeLimit = &HA00000'10MB

'was the Form successfully received?
Const fsCompletted  = 0




 if daForum<>"" then

		 	session("id_social")=0
	else  if daDiario<>"" then
		session("id_social")=1

		  else
		   	session("id_social")=2

		  end if
	end if

'if daForum<>"" then
'Conn1.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBForum")
'else
'Conn1.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBLavagna")
'end if
If Form.State = fsCompletted Then 'Completted





  'Create destination path+filename for the source file.
  Dim DestinationPath, DestinationFileName

  'DestinationPath = Server.mapPath("UploadFolder")
  'DestinationFileName = DestinationPath & "\" & Form("SourceFile").FileName
   if daForum<>"" then
	      DestinationPath = Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("Cartella") &"/file_forum"
	else  if daDiario<>"" then
		     DestinationPath = Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("Cartella") &"/file_diario"
		  else
	         DestinationPath = Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("Cartella") &"/file_lavagna"
		  end if
	end if

	DestinationPath=Replace(DestinationPath,"\","/")
	on error resume next

	If Err.Number = 0 Then
		 'RESPONSE.WRITE(querysql)
	'Response.Write "Modifica avvenuta! "
		stato=1
		messaggio="Modifica avvenuta"
	Else
		stato=0
		messaggio=Err.Description&"<br>"&Err.Source&"<br>"&Err.Number
		response.write(messaggio)
	Err.Number = 0
	End If

	'response.write("riga81:"&DestinationPath)
	' se ci sono gli archivi compressi li carico nella sottocartella codsiceallievo
	if Session("Zip")=1 then
	DestinationPath=DestinationPath&"/"&Session("IDTHREAD")&"/"&Session("CodiceAllievo")
	end if




   DestinationFileName = DestinationPath & "\" & Form("SourceFile").FileName

   if (instr(Form("SourceFile").FileName,".java")<>0) then
    SuffissoFileName=right(DestinationFileName,4) ' per il .java
   else
   SuffissoFileName=right(DestinationFileName,3) ' per il .zip,.jpg,...,
   end if
   DestinationPath=Replace(DestinationPath,"\","/")
   DestinationFileName=Replace(DestinationFileName,"\","/")
  ' Session("DestinationFileName")=DestinationFileName
  'Open recordset to store uploaded data
  Dim RS: Set RS = OpenUploadRS

  'Store extra info about upload to database
 ' RS.AddNew
'  ' RS("UploadDT") = Now()
'   RS("Description") = Form.Texts.Item("Description")
'   RS("SourceFileName") = Form("SourceFile").FilePath
'   RS("DestFileName") = DestinationFileName
'   RS("DataSize") = Form("SourceFile").Length
'   '...
'  RS.Update


    if daFrase="" then

   QuerySQL="  INSERT INTO FILE_FORUM (CodiceAllievo)  SELECT '" & Session("CodiceAllievo")  & "';"
   ConnessioneDB.Execute (QuerySQL)
   'QuerySQL="select max (ID_Smile) , Pos from FILE_FORUM group by Pos;"
   QuerySQL="select max (ID_Smile) from FILE_FORUM;"
   set rsTabella=ConnessioneDB.execute(QuerySQL)
   MAXID=rsTabella(0)
   MAXPOS=0
   'MAXPOS=rsTabella(1)
   ' nome del file di destinazione

   url=MAXID&"."&SuffissoFileName ' il nome sar� numero .zip,.jpg,ecc...

  ' codice=":;"&MAXID
   QuerySQL ="UPDATE FILE_FORUM SET Url = '" & url & "', Pos = " & MAXPOS+1 & ",  Nome = '" & Form("SourceFile").FileName & "', Href_O = '" & linkto &"' WHERE ID_Smile =" &MAXID&";"
   ConnessioneDB.execute(QuerySQL)
   'response.write(QuerySQL &"<br>")


if session("Zip")=1 then
' da implementare 23/10/13 per cambiare nome.zip da ID a CodiceALlievo
'destinazione=DestinationPath&"/"&Form("SourceFile").FileName
destinazione=DestinationPath&"/"&MAXID&"."&SuffissoFileName
else
destinazione=DestinationPath&"/"&MAXID&"."&SuffissoFileName
end if
' serve per farlo vedere a php che deve decomprimere
Session("Nomefilezip")=destinazione
Session("NomefilezipSint")=MAXID&"."&SuffissoFileName
Session("NomefilezipOrig")=Form("SourceFile").FileName
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url=server.MapPath("../../3PC/img_social/img/LogFileName.txt")
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(imgPath & "<br>" & destinazione)
'				objCreatedFile.Close
end if ' id daFrase="" then

 ' Response.write "<br>Nome del file :"
  Dim Field: For Each Field in Form.Files.Items
   ' Response.write "&nbsp;" & Field.FileName
  Next
  '{b}Save file to the destination
  Form("SourceFile").SaveAs DestinationFileName
  '{/b}

  'response.write "<Font color=green><br>Il file � stato salvato in  " & DestinationFileName
 ' response.write "<br>See ListFiles table in " & Server.MapPath(homesito)& "/database/upload.mdb"   & " database.</Font>"


' rinomino i file del file  copiando e cancellando
Set fso = CreateObject("Scripting.FileSystemObject")
set OggFile = fso.GetFile (DestinationFileName)
if daFrase="" then
   OggFile.Copy destinazione,true
   OggFile.Delete
else

       ' Set objTextFile = fso.OpenTextFile(DestinationFileName, ForReading)
'		sReadAll = objTextFile.ReadAll
'		'sReadAll=url
'		response.write(sReadAll)
'	    objTextFile.Close

    Session("NomeFileForum3")=DestinationFileName
	'if Request.ServerVariables("HTTP_REFERER") <>"" then
'							response.Redirect request.serverVariables("HTTP_REFERER")
'		 end if

end if

' salvo il nome del file che poi inserisco nella forum_messages per collegare il post al file
Session("CaricatoFile")=true
Session("NomeFileForum")=MAXID&"."&SuffissoFileName
Session("NomeFileForum2")=Form("SourceFile").FileName

if  not (daFrase="") then
%>
	<script language="javascript">
		opener.location.reload();
		window.close();
	</script>
<%
else%>
<script language="javascript">
    alert("File caricato, clicca su Invia")
		window.close();
	</script>

<%end if





ElseIf Form.State > 10 then
  Const fsSizeLimit = &HD
  Select case Form.State
		case fsSizeLimit: response.write  "<br><Font Color=red>La dimensione di (" & Form.TotalBytes & "B) supera i (" & Form.SizeLimit & "B) del limite</Font><br>"
		case else response.write "<br><Font Color=red>Form error.</Font><br>"
  end Select
End If'Form.State = 0 then

%>

<%


Function OpenUploadRS()
  Dim RS  : Set RS = CreateObject("ADODB.Recordset")
  'Open dynamic recordset, table Upload
  RS.Open "FILE_FORUM", GetConnection, 2, 2
  Set OpenUploadRS = RS
end Function

Function GetConnection()
  dim Conn: Set Conn = CreateObject("ADODB.Connection")


if session("DB")=1 then
 Conn.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=informistica; Password=123Maurosho;"
else
  Conn.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=informistica; Password=123Maurosho;"
end if

	set GetConnection = Conn

end function



%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
 <TITLE>Carica file</TITLE>
 <STYLE TYPE="text/css">
 TD	{font-family:Arial,Helvetica,sans-serif }
 TH	{font-family:Arial,Helvetica,sans-serif }
 TABLE	{font-size:10pt;font-family:Arial,Helvetica,sans-serif }
</STYLE>
 <meta name="robots" content="noindex,nofollow">
 <link rel="stylesheet" type="text/css" href="../../stile.css">
</HEAD>
<BODY BGColor=white>


<Div style=width:500>





<TABLE  id="zebra_forum1"  cellSpacing=1 cellPadding=3 bordercolor=silver bgcolor=GAINSBORO width="60%" border=1>
<form method=post ENCTYPE="multipart/form-data">

<TR>
 <TD>File da caricare</TD>
 <TD><input type="file" name="SourceFile" size="60" id="source" ></TD>
</TR>

 <TD>Dimensione massima (<%=(Form.SizeLimit \ 1024)\1024 %>Mb) </TD>
 <TD Align=Right><input type="submit" Name="Action" value="Upload file" class="btn-primary"></TD>
</TR>

</form></Table>


<HR COLOR=silver Size=1>

</Div>
</BODY>
 <script type="text/javascript">


$(window).load(function () {

	  $('#source').click();


	    event.stopPropagation();

	});

</script>
</HTML>
