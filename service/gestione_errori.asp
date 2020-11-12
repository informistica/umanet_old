<%
'Sub GestisciErrore(NumeroErrore, DescrizioneErrore,Pagina,Tipo)
Sub GestisciErrore(DescrizioneErrore,Spiegazione1,Pagina,Riga)
  'compongo il messaggio
  	Response.Write("<div class=contenuti><br><font color=red>Si &egrave; verificato un ERRORE in fase di esecuzione.<br>") 
	Response.Write("Puoi inviare una mail all'amministratore del sistema (informistica@umanet.net).<br>")
	Response.Write("fornendo le seguenti indicazioni circa l'errore : <br>")
  	'Response.write("<br>Numero&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "&NumeroErrore)
  	Response.write("<br>Descrizione : "&DescrizioneErrore)
	Response.write("<br>Dettagli : "&Spiegazione1)
  	Response.write("<br>Pagina&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "&Pagina)
  	Response.write("<br>Riga&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "&Riga)
  	Response.write("<br><br>Grazie per la collaborazione!</font></div>")
  	
  'ripulisco Err
  'Err.Clear 
 End Sub
%>
