<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
 <!-- #include file = "../script/var_globali.inc" -->
<%

CodiceAllievo=request.QueryString("CodiceAllievo")
url=request.QueryString("url")
%>
<html>
   <head>
      <title>Grafico</title>
   <style type="text/css">
   #apDiv1 {
	position:absolute;
	left:11px;
	top:11px;
	width:935px;
	height:50px;
	z-index:2;
	background-color: #FFFFFF;
}
   </style>
   </head>
   <body bgcolor="#ffffff">
   <div id="apDiv1"></div>
   <div align="center">
<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="https://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="100%" height="300" id="Column3D" >
         <param name="movie" value="../FusionCharts/Column3D.swf" />
         <param name="FlashVars" value="&dataURL=data_<%=CodiceAllievo%>.xml">
         <param name="quality" value="high" />
         <embed src="Column3D.swf" flashVars="&dataURL=data_<%=CodiceAllievo%>.xml" quality="high" width="1011" height="300" name="MSArea" type="application/x-shockwave-flash" pluginspage="https://www.macromedia.com/go/getflashplayer" />
     </object>
   </div>
   <%
   	    Set objFSO = CreateObject("Scripting.FileSystemObject")
		 'homesito="/anno_2010-2011_ITC/"   
		url=Server.MapPath(homesito & "/Grafici")& "/Data_"&CodiceAllievo &".xml"   
		'lo tolgo per lo esegue primna della creazione del grafico quinda da errore
		'objFSO.DeleteFile(url)
   %>
</body>
</html>