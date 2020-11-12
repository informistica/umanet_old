<%@ Language=VBScript %>

<% if Session("Admin") <> true then
Response.redirect "../../home.asp"
end if
%>

<!doctype html>
<html>
<head>

   <title>Modifica premetafora</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />


	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
<meta charset="utf-8">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	 <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />



    <script language="javascript" type="text/javascript">
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
    </script>
    <script type="text/javascript" src="../js/selezionatutti.js"></script>

<script language="javascript" type="text/javascript">
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"

 }
    </script>

  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
 <script src="../../js/datapicker_it.js"></script>

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<%
  Response.Buffer = true
  'On Error Resume Next
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  <% end if %>


	<div id="navigation">



		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->



	</div>

 <%
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
  CodiceTest = Request.QueryString("CodiceTest")
 %>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h3> <i class="icon-comments"></i> <%=Capitolo%>: <%=Paragrafo%></h3>

					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>
						<li>
							<a href="#">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Modifica premetafora</a>

						</li>

					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>







				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  METAFORE DISPONIBILI</h3>
			          </div>
				      <div class="box-content">


 	<%
 Modifica=Request.QueryString("Modifica")
  BoxApro=Request.QueryString("BoxApro")
  modifica_scadenze_classe=Request.QueryString("modifica_scadenze_classe")

  tutto=Request.QueryString("tutto")  ' se è settatto devo modificare tutte le f del capitolo
  modulo=Request.QueryString("Modulo")
Elimina=Request.QueryString("Elimina")
NumRec=Request.QueryString("NumRec")
ID=Request.QueryString("ID")
 Id_Stud=Request.QueryString("Id_Stud")   ' se è settato vuol dire che aggiungo eccezioni per singolo stud

if Elimina<>"" then



Select Case CodiceTest%>
   <% Case Cartella&"_U_2_3" 'Topolino%>
    <% Case Cartella&"_U_2_5" 'Navigazione%>
<%
	QuerySQL="Delete  " &_
"FROM preNavigazione WHERE ID_Premetafora=" & ID & ";"

%>

	 <% Case Cartella&"_U_2_8" 'Navigazione 
	' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
  end Select

response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)

'response.write(QuerySQL)



		if Request.ServerVariables("HTTP_REFERER") <>"" then
				response.Redirect request.serverVariables("HTTP_REFERER")
		end if

elseif Modifica="" then %>




  <%Cartella=Request.QueryString("Cartella")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")

  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")

  if Id_Stud<>""then
     QuerySQL="SELECT *  FROM Allievi WHERE CodiceAllievo='" & Id_Stud & "'"
  	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  	Cognome1=rsTabella("Cognome")
  	Nome1=rsTabella("Nome")%>
  	<p align="center"><font color="#FF0000" size="3">Modifica Scadenze per <%=Cognome1&" "%> <%=Nome1%> </font></p>
    <%
    end if
    if modifica_scadenze_classe<>"" then%>
    	<p align="center"><font color="#FF0000" size="3">Modifica Scadenze per classe </font></p>
<%  end if


'response.write(CodiceTest&"="&Cartella)
Select Case CodiceTest%>
   <% Case Cartella&"_U_2_3" 'Topolino%>
    <% Case Cartella&"_U_2_5" 'Navigazione%>
<%
	 
 
 QuerySQL="SELECT * FROM preNavigazione WHERE  Id_Paragrafo='" & CodiceTest & "' and Id_Sottoparagrafo='"&CodiceSottopar&"' order by Posizione"
QuerySQL2="SELECT count(*) FROM preNavigazione WHERE  Id_Paragrafo='" & CodiceTest & "' and Id_Sottoparagrafo='"&CodiceSottopar&"'"
  
 


%>

	 <% Case Cartella&"_U_2_8" 'Navigazione 
	' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
  end Select



 

Set rsTabella = ConnessioneDB.Execute(QuerySQL2)
'response.write("258:"&QuerySql2)
NumRec=rsTabella(0)
Numero=clng(NumRec)
Dim dom()
Redim dom(Numero)

 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)




'response.write("Numero ="&NumRec)

i=0
'paragrafo=rsTabella(2)
if rsTabella.eof and rsTabella.bof then%>
<span class="alert-error"><%=response.write("Non ci sono compiti assegnati")%></b></span>
<%end if%>
<form method="POST" name="dati" id="dati" class="form-vertical" >


    <p>Scadenza: <input type="text" name="txtDataVal" id="datepicker" class="input-medium datepick" /></p>
    <input type="button" class="btn"  onClick="selezionatutti('datepicker')" value="Applica a tutti">
    <hr>


<%do while not rsTabella.eof
	'if (i=0) or (StrComp(capitolo, rsTabella(0)) <> 0) then'

	dom(i)=rsTabella.fields("Quesito")
	'response.write(rsTabella("Img") & " " & rsTabella("Files"))

	if (rsTabella.fields("Img")=1)  then
	dom(i)=dom(i)& " $"
	end if

	if rsTabella("Files")<>0 then
	  dom(i)=dom(i)& " #"
	end if

	 %>
			<input type="text" class="hidden" value="<%=rsTabella.fields("ID_Premetafora")%>" name="txtIdFrase<%=i%>" size="3" >
			<fieldset><legend><%=i+1%> Frase 	</legend>
           							 <div class="control-group">

										<div class="controls">
											<input type="text" value="<%=rsTabella.fields("Quesito")%>" name="txtFrase<%=i%>" class="input-xxlarge"> &nbsp;&nbsp; <img src="../../img/elimina.jpg" width="16" height="16"  onClick="elimina(<%=rsTabella.fields("ID_Premetafora")%>);" title="Elimina"><br>
										</div>
									</div>

            <div class="control-group">

										<div class="controls">
                                        <b> <span title="Prevede la votazione ?">Vota</span> </b>
                                        <% if (rsTabella.fields("Img")=1)  then  %>
											 <INPUT TYPE="RADIO" name="txtImg<%=i%>" checked="true" value="1">Si
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>"  value="0">No
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>" value="1">Si
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>"   checked="true" value="0">No
										<% end if %>
                                        &nbsp;&nbsp;&nbsp;
                                         <span title="Cancella i voti assegnati ?"><b>Resetta</b></span>

                                             <% if (rsTabella.fields("Files")=1)  then  %>

											 <INPUT TYPE="RADIO" name="txtFile<%=i%>" checked="true" value="1">Si
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>"  value="0">No
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>" value="1">Si
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>"   checked="true" value="0">No

										<% end if %>
                                        

                                             <span title="Posizione nella lista"><b><span onclick="incrementa('txtPos<%=i%>',<%=i%>)"> +Pos</span></b></span>
                                             <input  class="input-mini" title="Numero d'ordine" type="text" value="<%=rsTabella.fields("Posizione")%>" id="txtPos<%=i%>" name="txtPos<%=i%>" size="1"  >
                                 &nbsp;&nbsp;&nbsp;
                                         <i title="Chiusura del compito" class="icon-calendar"></i>     <input type="text" value="<%=rsTabella.fields("Scadenza")%>" name="txtScadenza<%=i%>" id="scad<%=i%>"  class="input-small datepick"  />

										</div>


										<div id="divVerifica<%=i%>"  style="display:none"><b>Inserisci la risposta modello</b>
										 <textarea class="input-block-level" rows="3" name="txtModello<%=i%>">
		   								</textarea></p>
                           
                        				 </div><br>


									</div>







			</fieldset>

<%
	i=i+1
	rsTabella.movenext


		loop%>

		<hr>

        <div class="accordion" id="accordion2">
          <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse4">
												<center>Copia domande</center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
                                            <textarea rows=<%=NumRec%> class="input-block-level">
 <%
 for i=0 to NumRec-1
   response.write(dom(i)&chr(13))
 next

 %>
 </textarea>
                                            </div>
                                         </div>
                                      </div>
                                   </div>





   <input type="button" onClick="invia(0);" value="Modifica metafore" class="btn"><br><hr><br>
   

        </form>
			<br><br>
          <h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>#<%=BoxApro-3%>"> Torna al Libro... </a></h5>

		   <% if len(Id_Stud) > 0 then %>

			<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>&Id_Stud=<%=Id_Stud%>&modifica_scadenze_classe=<%=modifica_scadenze_classe%>">Torna al Libro (per modifica eccezioni)</a></h5>

			<% end if %>




 <br>

<% else ' aggiorno i campi ' aggiungo il test per capire se devo aggiungere eccezioni per Id_Stud
%>

 <%
 ' NumRec=Request.QueryString("NumRec")
  for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
   ID=Request.Form("txtIdFrase"&k)
   ID_Premetafora=ID
   Quesito = Request.Form("txtFrase"&k)

   Img=Request.Form("txtImg"&k)
   Verifica=Request.Form("txtVerifica"&k)
   cFile=Request.Form("txtFile"&k)
   if cFile="" then
      cFile=0
   end if
   if cFile<>0 then ' resetto i voti
		Select Case CodiceTest%>
		<% Case Cartella&"_U_2_3" 'Topolino%>
		<% Case Cartella&"_U_2_5" 'Navigazione%>
		<%
		QuerySQL="delete from [VotiMetaforaNavigazione] where Id_Premetafora=" & ID  & ";"
		%>
		<% Case Cartella&"_U_2_8" 'Navigazione 
		' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
		end Select
	 ConnessioneDB.Execute(QuerySQL)  
   end if
   Pos=Request.Form("txtPos"&k)
   Scadenza=Request.Form("txtScadenza"&k)
   if Scadenza="" then
      Scadenza=fine_anno
   end if
      Quesito = Replace(Quesito, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Quesito=  Replace(Quesito,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql

'response.write("modifica_scadenze_classe"&modifica_scadenze_classe)
   'TestoDomandaPlus=Request.Form("TestoDomandaPlus")
      if len(Id_Stud)>0 or modifica_scadenze_classe<>"" then ' aggiungo eccezioni

		   if DateDiff("D", Date(), Scadenza)>=0 then
           if modifica_scadenze_classe<>"" then
            QuerySQL="SELECT count(*) FROM Eccezioni_Frasi  WHERE  Id_Classe='"&session("Id_Classe")&"' and Id_Prefrase='"&ID&"';"
           else
  			   QuerySQL="SELECT count(*) FROM Eccezioni_Frasi  WHERE  Id_Stud='"&Id_Stud&"' and Id_Prefrase='"&ID&"';"
           end if
			   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			   ris=rsTabella(0)

				if ris=0 then
          if modifica_scadenze_classe<>"" then
        QuerySQL="INSERT INTO Eccezioni_Frasi (Id_Prefrase,Id_Classe,Scadenza) SELECT '" & ID  & "','" & session("Id_Classe") & "','" & Scadenza & "';"
          else
  					QuerySQL="INSERT INTO Eccezioni_Frasi (Id_Prefrase,Id_Stud,Scadenza) SELECT '" & ID  & "','" & Id_Stud & "','" & Scadenza & "';"
            end if
        	ConnessioneDB.Execute(QuerySQL)
				else
           if modifica_scadenze_classe<>"" then
            	QuerySQL = "UPDATE Eccezioni_Frasi SET Scadenza = '" &Scadenza&"' WHERE Id_Prefrase = '"&ID&"' and Id_Classe = '"& session("Id_Classe")&"';"
           else
					QuerySQL = "UPDATE Eccezioni_Frasi SET Scadenza = '" &Scadenza&"' WHERE Id_Prefrase = '"&ID&"' and Id_Stud = '"&Id_Stud&"';"
          end if
					ConnessioneDB.Execute(QuerySQL)
				end if
		   end if

	  else
	'response.write("<br>466"&CodiceTest&"="&Cartella)
			Select case CodiceTest
			Case Cartella&"_U_2_3"
			Case Cartella&"_U_2_5" 
				  QuerySQL ="UPDATE preNavigazione SET Quesito = '" & Quesito & "', Scadenza = '" & Scadenza & "', Img = " & Img & ", Posizione = " & Pos& ", Files = " & cFile &" WHERE ID_Premetafora =" &ID&";"

			Case Cartella&"_U_2_8"	

			end Select

	    
      ConnessioneDB.Execute(QuerySQL)
	   end if
	 ' response.Write("<br>"&QuerySQL)





   next
 %>

			<p><p>
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->

<h5>Modifica Effettuata...<%if modifica_scadenze_classe<>"" then response.write ("per tutta la classe") end if%></h5><br><br>
<h5><a href="../cClasse/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>#<%=BoxApro%>">Torna al Libro</a></h5>

<% if len(Id_Stud) > 0 then %>
	<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>&Id_Stud=<%=Id_Stud%>&modifica_scadenze_classe=<%=modifica_scadenze_classe%>">Torna al Libro (per modifica eccezioni)</a></h5>
<% end if %>

<% end if %>



                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->


		</div> <!--fine main-->
        </div>


 <script language="javascript" type="text/javascript">

function invia(pagina) {
	
 if (pagina==0)
	{ //modifica scadenze
	 document.dati.action="2modificapremetafora.asp?modifica_scadenze_classe=<%=modifica_scadenze_classe%>&BoxApro=<%=BoxApro%>&Modifica=1&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=CodiceTest%>&NumRec=<%=NumRec%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&SottoParagrafo=<%=CodiceSottopar%>";
	document.dati.submit();
	}
	else


	{//creo verifica
 
   
	if (confirm("Sei sicuro di voler creare una verifica")){
	
		document.dati.action = "../cAdmin/inserisci_verifica.asp?Id_Classe=<%=Id_Classe%>&ID_Mod=<%=ID_Mod%>&Titolo=<%=Titolo%>&classe=<%=classe%>&cartella=<%=cartella%>&Num=<%=NumRec%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>";  
		document.dati.submit();}
	else
	return; 
    }
	

}


function inserisci_inverifica(s,id)
{
if (s==0)
 document.getElementById("divVerifica"+id).style.display='block';
 else
   document.getElementById("divVerifica"+id).style.display='none';
}



function incrementa(ids,idx) {

   //var i = document.getElementById(ids).value;
   //document.getElementById(ids).value = parseInt(i) + 1;


	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		if (elements[i].id == 'txtPos'+idx)
		    {
		     var val=elements[i].value;
			 elements[i].value=parseInt(val)+1;
			 idx=idx+1;
			}
		}
	}


}

 function elimina(id) {
	 
	 document.dati.action="2modificapremetafora.asp?ID="+id+"&Elimina=1&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=CodiceTest%>&NumRec=<%=NumRec%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&SottoParagrafo=<%=CodiceSottopar%>";
	
		//document.dati.action = "../../home.asp"
		document.dati.submit();



}
 </script>

	</body>

 </html>
