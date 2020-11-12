<%
'Sample file Field-SaveAs.asp 
'Store extra upload info to a database
' and file contents to the disk
Server.ScriptTimeout = 5000
	 session("CaricatoFile")=false 
     session("NomeFileForum")=""
	 daForum=request.QueryString("daForum")
	 daDiario=request.QueryString("daDiario")

'Create upload form
'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm %>

<!--#INCLUDE FILE="_upload.asp"-->
<!--#INCLUDE FILE="../var_globali.inc"-->

<% 


Server.ScriptTimeout = 1000
Form.SizeLimit = &HA00000'10MB

'was the Form successfully received?
Const fsCompletted  = 0
dim Conn1: Set Conn1 = CreateObject("ADODB.Connection")
Conn1.Provider = "Microsoft.Jet.OLEDB.4.0"

 if daForum<>"" then
	     Conn1.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBForum")    
	else  if daDiario<>"" then
		     Conn1.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBDiario")  
		  else
	        Conn1.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBLavagna")  
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
	      DestinationPath = Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/file_forum"   
	else  if daDiario<>"" then
		     DestinationPath = Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/file_diario"  
		  else
	         DestinationPath = Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/file_lavagna" 
		  end if
	end if
	 
	 
	
   
   DestinationFileName = DestinationPath & "\" & Form("SourceFile").FileName
   SuffissoFileName=right(DestinationFileName,3) ' per il .zip,.jpg,...,
   DestinationPath=Replace(DestinationPath,"\","/")
   DestinationFileName=Replace(DestinationFileName,"\","/")

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


   QuerySQL="  INSERT INTO FILE_FORUM (CodiceAllievo)  SELECT '" & Session("CodiceAllievo")  & "';"   
   Conn1.Execute (QuerySQL) 
   QuerySQL="select max (ID_Smile) , Pos from FILE_FORUM group by Pos;"
   set rsTabella=Conn1.execute(QuerySQL)
   MAXID=rsTabella(0)
   MAXPOS=rsTabella(1)
   ' nome del file di destinazione 
  
   url=MAXID&"."&SuffissoFileName ' il nome sarà numero .zip,.jpg,ecc...
   
  ' codice=":;"&MAXID
   QuerySQL ="UPDATE FILE_FORUM SET Url = '" & url & "', Pos = " & MAXPOS+1 & ",  Nome = '" & Form("SourceFile").FileName & "', Href_O = '" & linkto &"' WHERE ID_Smile =" &MAXID&";"
   Conn1.execute(QuerySQL)
   'response.write(QuerySQL &"<br>")
 

destinazione=DestinationPath&"/"&MAXID&"."&SuffissoFileName
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url=server.MapPath("../../3PC/img_social/img/LogFileName.txt")
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(imgPath & "<br>" & destinazione)
'				objCreatedFile.Close


 ' Response.write "<br>Nome del file :"
  Dim Field: For Each Field in Form.Files.Items
   ' Response.write "&nbsp;" & Field.FileName
  Next
  '{b}Save file to the destination
  Form("SourceFile").SaveAs DestinationFileName
  '{/b}

  'response.write "<Font color=green><br>Il file è stato salvato in  " & DestinationFileName
 ' response.write "<br>See ListFiles table in " & Server.MapPath(homesito)& "/database/upload.mdb"   & " database.</Font>"
 
 
' rinomino i file del file  copiando e cancellando	
Set fso = CreateObject("Scripting.FileSystemObject")
set OggFile = fso.GetFile (DestinationFileName)
OggFile.Copy destinazione,true
OggFile.Delete

' salvo il nome del file che poi inserisco nella forum_messages per collegare il post al file
Session("CaricatoFile")=true
Session("NomeFileForum")=MAXID&"."&SuffissoFileName
Session("NomeFileForum2")=Form("SourceFile").FileName

%>
	<script language="javascript">
		
		window.close();
	</script>
<%






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
  Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
 ' Conn.open "Data Source=" & Server.MapPath("upload.mdb") 
 'Conn.open "Data Source=" & Server.MapPath("upload.mdb") 
  
    if daForum<>"" then
 	Conn.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBForum")  
  else
	Conn.open "Data Source=" & Server.MapPath(homesito)& "/database/" & Session("DBLavagna")  
  end if
	set GetConnection = Conn
end function



%>  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
 <TITLE>Carica file</TITLE>
 <STYLE TYPE="text/css"><!--TD	{font-family:Arial,Helvetica,sans-serif }TH	{font-family:Arial,Helvetica,sans-serif }TABLE	{font-size:10pt;font-family:Arial,Helvetica,sans-serif }--></STYLE>
 <meta name="robots" content="noindex,nofollow">
 <link rel="stylesheet" type="text/css" href="../../stile.css">
</HEAD>
<BODY BGColor=white>


<Div style=width:500>
 
 



<TABLE  id="zebra_forum1"  cellSpacing=1 cellPadding=3 bordercolor=silver bgcolor=GAINSBORO width="60%" border=1>
<form method=post ENCTYPE="multipart/form-data">

<TR>
 <TD>File da caricare</TD>
 <TD><input type="file" name="SourceFile" size="60"></TD>
</TR>

 <TD>Dimensione massima (<%=(Form.SizeLimit \ 1024)\1024 %>Mb) </TD>
 <TD Align=Right><input type="submit" Name="Action" value="Upload file &gt;&gt;"></TD>
</TR>

</form></Table>




<HR COLOR=silver Size=1>

</Div>
</BODY></HTML>
