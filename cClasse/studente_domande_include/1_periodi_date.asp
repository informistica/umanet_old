
<%

daStud=Request.QueryString("daStud") ' chiamato da href della classifica
daForm=Request.QueryString("daForm") ' chiamato dal bottone invia di scelta periodi dal al
daMenu=Request.QueryString("daMenu")
centromsg=request.QueryString("centromsg") 

if daForm<>"" then ' se sono chiamato dal form dei Periodi
DataCla=request.form("txtData") 
DataCla2=request.form("txtData2")
end if

if daStud<>"" then 'per il quaderno se sono chiamato dalla classifica 
DataCla=request.QueryString("DataClaq") 
DataCla2=request.QueryString("DataClaq2")
end if

if daMenu<>"" then ' se sono chiamato dal menu left o navigation
DataCla=request.QueryString("DataClaq") 
DataCla2=request.QueryString("DataClaq2")
end if
'response.write("<br>centromsg="&centromsg)
if centromsg<>"" then
'response.write("SI")
DataCla=request.QueryString("DataCla") 
DataCla2=request.QueryString("DataCla2")
end if


  Session("DataClaq")=DataCla
  Session("DataClaq2")=DataCla2
  Session("DataCla")=DataCla
  Session("DataCla2")=DataCla2
  DataClaq=Session("DataClaq")
  DataClaq2=Session("DataClaq2")

'
'
' response.write()
' response.write("<br>DataCla2="&DataCla2)
' response.write("<br>DataClaq="&DataClaq)
' response.write("<br>DataClaq2="&DataClaq2)
' response.write("<br>SessionDataCla="&Session("DataCla"))
' response.write("<br>SessionDataCla2="&Session("DataCla2"))
'  response.write("<br>SessionDataClaq="&Session("DataClaq"))
' response.write("<br>SessionDataClaq2="&Session("DataClaq2"))
' 
' Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logPPP.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine("<br>DataCla="&DataCla)
'				objCreatedFile.WriteLine("<br>DataCla2="&DataCla2)
'				objCreatedFile.WriteLine("<br>DataClaq="&DataClaq)
'				objCreatedFile.WriteLine("<br>DataClaq2="&DataClaq2)
'				objCreatedFile.WriteLine("<br>SessionDataCla="&Session("DataCla"))
'				objCreatedFile.WriteLine("<br>SessionDataCla2="&Session("DataCla2"))
'				objCreatedFile.WriteLine("<br>SessionDataClaq="&Session("DataClaq"))
'				objCreatedFile.WriteLine("<br>SessionDataClaq2="&Session("DataClaq2"))
'				objCreatedFile.Close

%>