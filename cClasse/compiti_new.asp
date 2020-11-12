



 <!-- #include file = "studente_domande_include/4_quaderno.asp" -->

		<%
		 DataClaN = DataCla
		DataCla2N = DataCla2
		

		 if DataClaN="" then 
  DataClaN=Session("DataCla")
 end if
DataCla = DataClaN
DataClaq = DataClaN


 if DataCla2N="" then 
  DataClaN2=Session("DataCla2")
 end if
 
DataCla2 = DataCla2N
DataClaq2 = DataCla2N
 
		 %>


		<!-- #include file = "../var_globali.inc" -->


 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

		<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->
        <!-- #include file = "../cClasse/studente_domande_include/1_periodi_date.asp" -->

        <!-- #include file = "../extra/test_server.asp" -->

 <!-- #include file = "../cUtenti/adovbs.inc" -->

 <% 
DataCla = DataClaN
DataClaq = DataClaN
DataCla2 = DataCla2N
DataClaq2 = DataCla2N
		%>

 		<!-- #include file = "../include/formattaDataCla.inc" -->

<% 'response.write DataCla
%>

 <%

 CodiceAllievo = Request.querystring("cod")

 'per le store procedure
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd1.activeconnection = conn
set cmd2.activeconnection = conn
set cmd3.activeconnection = conn
umanet=request.querystring("umanet")
%>



 	 

 <%
if strcomp(umanet,"0")=0 then
 QuerySQL="SELECT * FROM MODULI_CLASSE " &_
 " WHERE Id_Classe='" & id_classe  &"'"
 '" WHERE Id_Classe='" & id_classe &"'" & superIdClasse   ' carica i titoli dei moduli ma non il contenuto
else
QuerySQL="SELECT * FROM MODULI_CLASSE_UMANET " &_
 " WHERE Id_Classe='" & id_classe  &"'"

end If 

  Set rsTabellaModuli = ConnessioneDB.Execute(QuerySQL)
   '  response.write(QuerySQL)
 %>
 <% k=0
 p=0
   compiti=0 ' serve per mettere il box se non ci sono compiti inseriti
		     do while not rsTabellaModuli.EOF
			 ' calcolo i punteggi frase per quel modulo
			 %>

			  <!-- #include file = "studente_domande_include/3_statistica_frasi.asp" -->
              <!-- #include file = "studente_domande_include/3_statistica_nodi.asp" -->
              <!-- #include file = "studente_domande_include/3_statistica_domande.asp" -->
			   <%
			   numrsMetafore=0
			   if strcomp(umanet,"1")=0 then%>
              <!-- #include file = "studente_domande_include/3_statistica_metafore.asp" -->
              <%end if%>
              <%
'
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\745.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 82"
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
			  %>


				<% ' numrsFrasi = numero di compiti inseriti da stud
                   ' numrsFrasi2 = punti ottenuti ; Pb =numrsFrasi2/numrsFrasi
                  '  numrsPreFrasi= compiti totali inseriti dal prof


                 %>




 <%
  idCap=rsTabellaModuli("ID_Mod")  ' parametro per il loader ajax
' response.write(numrsFrasi& " --" & numrsNodi & "---" & numrsDomande)
 ' se è stato svolto almeno un compito mostro il capitolo

 
 if (numrsFrasi<>0) or (numrsNodi<>0) or (numrsDomande<>0) or (numrsMetafore<>0)then  ' devo fare anche per nodi e domande mostro solo dove ci sono compiti svolti
 compiti=compiti+1
 %>

               <div class="accordion-group" id="accordionnew<%=k%>">
                  <div class="accordion-heading">
                    <a onclick="loader('<%=idCap%>',<%=k%>)" style="text-decoration:none" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%>" href="#collapsenew<%=k%>"  id="toggleCapitolo<%=k%>" title="<%=k%>">
                        <%=rsTabellaModuli("Titolo") %><small> (<% Response.write(numrsFrasi2+numrsNodi2+numrsDomande2+numrsMetafore2)%>)</small> &nbsp;&nbsp;
						<% if numrsFrasi2 > 0 then %><small><i class="icon-reply"></i>(<% Response.write(numrsFrasi2)%>)</small><%end if%>&nbsp;&nbsp;
						<% if numrsNodi2 > 0 then %><small><i class="glyphicon-snowflake"></i></small>(<% Response.write(numrsNodi2)%>)<%end if%>&nbsp;&nbsp;
						<% if numrsDomande2 > 0 then %><small><i class="icon-question-sign"></i>(<% Response.write(numrsDomande2)%>)</small><%end if%> 
						<% if numrsMetafore2 > 0 then %><small><i class="icon-picture"></i>(<% Response.write(numrsMetafore2)%>)</small><%end if%>
                    </a>
                  </div>
                 <div id="collapsenew<%=k%>" class="accordion-body collapse">
                    <div class="accordion-inner" id="accordion-inner<%=k%>">
                     
                    </div><!-- fine accordion inner-->
                  </div>

                </div> <!-- aggiunto solo qui-->
                
    <% end if  ' if (numrsFrasi<>0) or (numrsNodi<>0) or (numrsDomande<>0)then
	 %>
			<% k=k+1
			   rsTabellaModuli.movenext()
			Loop
			%>

            <% if compiti=0 then %>
            <span class="alert-error"><h5>Nessun compito inserito nel periodo dal
            <%response.write(cdate(DataClaq)&" al ")%>
            <%response.write(cdate(DataClaq2))%>
            </h5></span>
            <%
			end if
			%>

<script>
function loader(idmod,k){
    console.log(idmod+" "+k);


var query = window.location.search.substring(1)+"&idmod="+idmod;
	console.log(query);
	var url = "https://www.umanetexpo.net/expo2015Server/UECDL/script/cClasse/compiti_modulo.asp?"+query;
	//alert(url);

	var testo;
	var stato1, stato2;

//$("#accordion-inner"+k).html("Attendere prego...");
	//$("#compitispec").html("Attendere prego...");
	$("#accordion-inner"+k).html("<img src='taoloader.gif'> Caricamento...");

	//eseguo chiamata http
					var xhttp = new XMLHttpRequest();
					xhttp.onreadystatechange = function() {

						stato1=xhttp.readyState;
						stato2=xhttp.status;

						if(stato1==4 && stato2==200){

						testo = xhttp.responseText;

						$("#accordion-inner"+k).empty().append(testo);

						}


					};

					xhttp.open("GET", url, true);
					xhttp.send();





	}
function cancella_frase(CodiceFrase,riga,Modulo,Paragrafo,Cartella,CodiceAllievo) {
	if (window.confirm('Vuoi veramente cancellare la frase?')) {
	  	 	var url="../cFrasi/cancella_frase_ajax.asp?CodiceFrase="+CodiceFrase+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&Cartella="+Cartella+"&CodiceAllievo="+CodiceAllievo;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								if (risposta=="Cancellazione avvenuta!")
									$('#riga_'+riga).remove();
								else
									alert(risposta);
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 }

}
</script>





               </div> <!--<div class="bs-docs-example"> fino blocco compiti -->
