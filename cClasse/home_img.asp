<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<title>Menu Interfaccia U-WWW</title>

<link href="../stile.css" rel="stylesheet" type="text/css" />
</head>

<body>
<%
' per sapere se sono in esecuzione sul server o in locale, serve per distinguere gli url per le risorse
  pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
  if (left(pathEnd1,10)="c:\inetpub") then
     locale=1
  else
     locale=0
  end if 
'response.write("LOcale = " &locale)
%>


		<div class="immagini" align="center">
		<img src="../../img/UWWW.jpg" border="0" usemap="#Map" />
		  <map name="Map" id="Map">
		   <% if locale = 1 then %>
		<area shape="rect" coords="299,64,623,85" href="UECDL/UWWW/Umanet/Pagine/Patente_Pratica.html" target="_blank" title="Domande per l'Eletto" />
		<%else%>
		<area shape="rect" coords="299,64,623,85" href="https://www.umanetexpo.net/informistica/UWWW/Umanet/Pagine/Patente_Pratica.html" target="_blank" title="Domande per l'Eletto" />
		<%end if %>
		  
		  
		   <% if locale = 1 then %>
		  <area shape="rect" coords="388,94,555,179" href="UECDL/UWWW/Umanet/Pagine/Patente_Teorica.html" target="_blank" title="Corrispondenze tra Metafore e Realtà"/>
		   <%else%>
		   <area shape="rect" coords="388,94,555,179" href="https://www.umanetexpo.net/informistica/UWWW/Umanet/Pagine/Patente_Teorica.html" target="_blank" title="Corrispondenze tra Metafore e Realtà"/>
		   <%end if %>
		   
		   <% if locale = 1 then %>
		  <area shape="rect" coords="687,99,903,225" href="UECDL/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html" target="_blank" title="Come si crea la metafora"/>
		   <%else%>
		     <area shape="rect" coords="687,99,903,225" href="https://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html" target="_blank" title="Come si crea la metafora"/>
		    <%end if %>
		  
		  
		  <% if locale = 1 then %>
		  <area shape="rect" coords="23,101,239,227" href="UECDL/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html" target="_blank" title="Come si crea la metafora"/>
		   <%else%>
		     <area shape="rect" coords="23,101,239,227" href="https://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html" target="_blank" title="Come si crea la metafora"/>
		    <%end if %>
		  
		  <% if locale = 1 then %>
		  <area shape="rect" coords="273,220,647,405" href="UECDL/UWWW/Metafore/Pagine/Umanet_Explorer.html" target="_blank" title="Navigazione per Immagini"/>
		  <%else%>
		  <area shape="rect" coords="273,220,647,405" href="https://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Umanet_Explorer.html" target="_blank" title="Navigazione per Immagini"/>
		  <%end if %>
		  <% if locale = 1 then %>
		   <area shape="rect" coords="15,251,245,411" href="UECDL/UWWW/Umanet/Pagine/Protocolli_di_Rete.html" title="Protocolli di Rete"/>
		  <%else%>
		  <area shape="rect" coords="15,251,245,411" href="https://www.umanetexpo.net/informistica/UWWW/Umanet/Pagine/Protocolli_di_Rete.html" title="Protocolli di Rete"/>
		   <%end if%>
		   <% if locale = 1 then %>
		   <area shape="rect" coords="678,251,905,413"href="UECDL/UWWW/Umanet/Pagine/Protocolli_di_Rete.html" title="Protocolli di Rete" />
		  <%else%>
		  <area shape="rect" coords="678,251,905,413" href="https://www.umanetexpo.net/informistica/UWWW/Umanet/Pagine/Protocolli_di_Rete.html"  title="Protocolli di Rete"/>
		   <%end if%>
		   
		  
		  
		  
		   <% if locale = 1 then %>
		  <area shape="rect" coords="252,413,653,561" href="UECDL/UWWW/Umanet/Pagine/Esempio_1_1.html" target="_blank" title="La Comunicazione tra Client e Server" />
		  <%else%>
		     <area shape="rect" coords="252,413,653,561" href="https://www.umanetexpo.net/informistica/UWWW/Umanet/Pagine/Esempio_1_1.html" target="_blank" title="La Comunicazione tra Client e Server" />
		  <%end if %>
		  <area shape="rect" coords="24,431,229,560" href="#connsx" />
		  
		  <area shape="rect" coords="685,431,893,558" href="#conndx" />
		  
		  <area shape="rect" coords="261,563,435,635" href="#connesinterna" />
		  
		  <area shape="rect" coords="481,563,661,636" href="#connesesterna" />
		  
		  <area shape="rect" coords="12,576,250,633" href="#codiceinterno" />
		  
		  <area shape="rect" coords="674,572,915,634" href="#codiceesterno" />
		  
		   <% if locale = 1 then %>
		  <area shape="rect" coords="261,639,431,705" href="UECDL/UWWW/Metafore/Pagine/Teoria_delle_Stringhe.html#2" target="_blank"  title="Macchina Regola Universo"/>
		  <%else%>
		   <area shape="rect" coords="261,639,431,705" href="https://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Teoria_delle_Stringhe.html#2" target="_blank"  title="Macchina Regola Universo"/>
		  <%end if%>
		  
		   <% if locale = 1 then %>
		  <area shape="rect"  coords="480,640,663,707" href="UECDL/UWWW/Metafore/Pagine/Teoria_delle_Stringhe.html#2" target="_blank"  title="Macchina Regola Universo" />
		  <%else%>
		   <area shape="rect"  coords="480,640,663,707" href="https://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Teoria_delle_Stringhe.html#2"  target="_blank" title="Macchina Regola Universo" />
		  <%end if%>
		  
	
		  
		  <area shape="rect" coords="10,638,254,708" href="#infointerne" />
		  
		  <area shape="rect" coords="673,638,911,707" href="#infoesterne" />
		  
		  <area shape="rect" coords="12,712,421,933" href="#ostesx" />
		  
		  <area shape="rect" coords="484,708,908,931" href="#ostedx" />
		  
		  
		   
		   <% if locale = 1 then %>
		  <area shape="rect" coords="15,936,207,1094" href="UECDL/UWWW/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%else%>
		   <area shape="rect" coords="15,936,207,1094" href="https://www.umanetexpo.net/informistica/UWWW/Navigazione/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%end if%>
		   
		   
		   
		   
		   
		 
		  <% if locale = 1 then %>
		  <area shape="rect" coords="214,937,425,1093" href="UECDL/UWWW/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%else%>
		   <area shape="rect" coords="214,937,425,1093" href="https://www.umanetexpo.net/informistica/UWWW/Navigazione/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%end if%>
		   
		    <% if locale = 1 then %>
		  <area shape="rect" coords="489,933,677,1089"  href="UECDL/UWWW/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%else%>
		   <area shape="rect" coords="489,933,677,1089"  href="https://www.umanetexpo.net/informistica/UWWW/Navigazione/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%end if%>
		   
		    <% if locale = 1 then %>
		  <area shape="rect" coords="685,932,896,1092" href="UECDL/UWWW/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%else%>
		   <area shape="rect" coords="685,932,896,1092" href="https://www.umanetexpo.net/informistica/UWWW/Navigazione/Navigazione.html" title="Navigazione in Video" / target="_blank">
		   <%end if%>
		   
		 
		 
		 
		 
		  </map>
</div>
	
</body>
</html>
