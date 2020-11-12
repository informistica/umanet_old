<% CodiceAllievo = Request.querystring("cod") %>
<% 'response.write "Stringa: " & Session("stringalink") 

stringalink = Session("stringalink")
stringaproperty = Session("stringaproperty")
stringastud = Session("stringastud")
Session.Contents.Remove(Session("stringalink"))
Session.Contents.Remove(Session("stringaproperty"))
Session.Contents.Remove(Session("stringastud"))

stringalink = Left(stringalink, Len(stringalink)-1)
stringaproperty = Left(stringaproperty, Len(stringaproperty)-1)
stringastud = Left(stringastud, Len(stringastud)-1)

'Dim vlink
'Dim vproperty
'vlink = Split(stringalink,",")
'vproperty = Split(stringaproperty,",")

'response.write(vlink(0))
'response.write("<br>"&vproperty(0))

%>

<!DOCTYPE html>
<html lang="it-IT">

<head>
   
   
   <meta http-equiv='cache-control' content='no-cache'>
<meta http-equiv='expires' content='0'>
<meta http-equiv='pragma' content='no-cache'>
    <meta name="author" content="Vincent Link, Steffen Lohmann, Eduard Marbach, Stefan Negru, Vitalis Wiens" />
    <meta name="keywords" content="webvowl, vowl, visual notation, web ontology language, owl, rdf, ontology visualization, ontologies, semantic web" />
    <meta name="description" content="WebVOWL - Web-based Visualization of Ontologies" />
    <meta name="robots" content="noindex,nofollow" />
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=1">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <link rel="stylesheet" type="text/css" href="css/webvowl.css" />
    <link rel="stylesheet" type="text/css" href="css/webvowl.app.css" />
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
    <link rel="icon" href="favicon.ico" type="image/x-icon" />
	<meta http-equiv="Content-Type" content="text/html;charset=UTF-8">

    <title>Mappa</title>
	
</head>

<body>

    <main>

        <section id="canvasArea">
            <div id="browserCheck" class="hidden">
                WebVOWL does not work properly in Internet Explorer and Microsoft Edge. Please use another browser, such as
                <a href="http://www.mozilla.org/firefox/">Mozilla Firefox</a> or
                <a href="https://www.google.com/chrome/">Google Chrome</a>, to run WebVOWL.
                <label id="killWarning">Hide warning</label>
            </div>

            <!-- <div id="logo">


                <h2>WebVOWL <br/><span>1.0.6</span></h2>

            </div>-->
            <div id="loading-info">
                <div id="loading-error" class="hidden">
                    <div id="error-info"></div>
                    <div id="error-description-button" class="hidden">Show error details</div>
                    <div id="error-description-container" class="hidden">
                        <pre id="error-description"></pre>
                    </div>
                </div>
                <div><span id="sidebarExpandButton"> > </span></div>

                <div id="loading-progress" class="hidden">
                    <span>Loading ontology... </span>
                    <div class="spin">&#8635;</div><br>
                    <div id="myProgress">
                        <div id="myBar"></div>
                    </div>

                </div>
            </div>


            <div id="graph"></div>
        </section>
        <aside id="detailsArea">
	        
	        <% collegamento = Request.QueryString("collegamento")	
			
			if Session("CodiceAllievo") <> "" then
			
				
		        if collegamento <> 1 then collegamento = 0 end if
				
				%>
				<script>
				var collegamento = <%=collegamento%>;
				var stringacollegamento = '<br><center><span style="color:white">Per aggiungere, modificare o eliminare collegamenti</span></center><br><center><a href="javascript:void(0)" onclick="entracollegamento()"><button>Entra in modalità collegamento</button></a></center><br>';
				var stringavisualizzazione = '<br><center><a href="javascript:void(0)" onclick="annullamodifiche(1)"><button>Annulla</button></a>&nbsp;<a href="javascript:void(0)" onclick="escicollegamento()"><button>Esci</button></a></center><br><br><center><span style="color:white">Clicca sul pulsante qui sotto<br>per caricare tutte le modifiche effettuate</span></center><br><center><a onclick="refresh()" href="javascript:void(0)"><button>Ricarica mappa</button></a><br><br></center>';
				</script>		
				<div id="stringainfo">
				<script>
				document.write(stringacollegamento);
				</script>
				</div>
				<%else %>
				
				<div id="stringainfo">
				<script>
				var collegamento = <%=collegamento%>;
				var stringacollegamento = '<br><center><span style="color:white">La mappa è stata condivisa: non puoi effettuare modifiche, ma solo visualizzare.</span></center><br>';
				
				document.write(stringacollegamento);
				</script>
				</div>
				
				<%end if%>
            <section id="generalDetails">
                <h1 id="title"></h1>
                <!--<span><a id="about" href=""></a></span>
                <h5>Version: <span id="version"></span></h5>
                <h5>Author(s): <span id="authors"></span></h5>
                <h5>
                    <label>Language: <select id="language" name="language" size="1"></select></label>
                </h5>-->
                <h3 class="accordion-trigger accordion-trigger-active">Descrizione</h3>
                <div class="accordion-container scrollable">
                    <p id="description"></p>
                </div>
                <!--<h3 class="accordion-trigger">Metadata</h3>
                <div id="ontology-metadata" class="accordion-container"></div>-->
                <h3 class="accordion-trigger">Statistiche</h3>
                <div class="accordion-container">
                    <p class="statisticDetails">Nodi: <span id="classCount"></span></p>
                    <p class="statisticDetails">Connessioni: <span id="objectPropertyCount"></span></p>
                </div>
                <h3 class="accordion-trigger" id="selection-details-trigger">Dettagli</h3>
                <div class="accordion-container" id="selection-details">
                    <div id="classSelectionInformation" class="hidden">
                        <p class="propDetails">Name: <span id="name"></span></p>
						
                        <!--<p class="propDetails">Type: <span id="typeNode"></span></p>-->
                        <p class="propDetails">Equiv.: <span id="classEquivUri"></span></p>
                        <p class="propDetails">Disjoint: <span id="disjointNodes"></span></p>
                        <p class="propDetails" style="display:none">Charac.: <span id="classAttributes"></span></p>
                        <p class="propDetails">Individuals: <span id="individuals"></span></p>
                        <p class="propDetails">Description: <span id="nodeDescription"></span></p>
                        <p class="propDetails">Comment: <span id="nodeComment"></span></p>
                    </div>
                    <div id="propertySelectionInformation" class="hidden">
                        <p class="propDetails">Name: <span id="propname"></span></p>
                        <!--<p class="propDetails">Type: <span id="typeProp"></span></p>-->
                        <p id="inverse" class="propDetails">Inverse: <span></span></p>
                        <p class="propDetails">Domain: <span id="domain"></span></p>
                        <p class="propDetails">Range: <span id="range"></span></p>
                        <p class="propDetails">Subprop.: <span id="subproperties"></span></p>
                        <p class="propDetails">Superprop.: <span id="superproperties"></span></p>
                        <p class="propDetails">Equiv.: <span id="propEquivUri"></span></p>
                        <p id="infoCardinality" class="propDetails">Cardinality: <span></span></p>
                        <p id="minCardinality" class="propDetails">Min. cardinality: <span></span></p>
                        <p id="maxCardinality" class="propDetails">Max. cardinality: <span></span></p>
                        <p class="propDetails">Charac.: <span id="propAttributes"></span></p>
                        <p class="propDetails">Description: <span id="propDescription"></span></p>
                        <p class="propDetails">Comment: <span id="propComment"></span></p>
                    </div>
                    <div id="noSelectionInformation">
                        <p><span>Seleziona un elemento nella mappa.</span></p>
                    </div>
                </div>
            </section>
									
        </aside>
        <nav id="optionsArea">
            <ul id="optionsMenu">
                <li id="aboutMenu"><a href="#">About</a>
                    <ul class="toolTipMenu aboutMenu">
                        <li><a href="license.txt" target="_blank">MIT License &copy; 2014-2017</a></li>
                        <li id="credits" class="option">WebVOWL Developers:<br/> Vincent Link, Steffen Lohmann, Eduard Marbach, Stefan Negru, Vitalis Wiens
                        </li>

                        <li><a href="http://vowl.visualdataweb.org/webvowl.html#releases" target="_blank">Version: 1.0.6<br/>(release history)</a></li>

                        <li><a href="http://purl.org/vowl/" target="_blank">VOWL Specification &raquo;</a></li>
                    </ul>
                </li>
                <li id="pauseOption"><a id="pause-button" href="#">Pause</a></li>
                <li id="resetOption"><a id="reset-button" href="#" type="reset">Reset</a></li>
                <li id="moduleOption"><a href="#">Modes</a>
                    <ul class="toolTipMenu module">
                        <!--<li class="toggleOption" id="helicopterZoom"></li>-->
                        <!--<li class="toggleOption" id="grassHopperZoom"></li>-->
                        <li class="toggleOption" id="dynamicLabelWidth"></li>
                        <li class="toggleOption" id="pickAndPinOption"></li>
                        <li class="toggleOption" id="nodeScalingOption"></li>
                        <li class="toggleOption" id="compactNotationOption"></li>
                        <li class="toggleOption" id="colorExternalsOption"></li>
                    </ul>
                </li>
                <li id="filterOption"><a href="#">Filter</a>
                    <ul class="toolTipMenu filter">
                        <li class="toggleOption" id="datatypeFilteringOption"></li>
                        <li class="toggleOption" id="objectPropertyFilteringOption"></li>
                        <li class="toggleOption" id="subclassFilteringOption"></li>
                        <li class="toggleOption" id="disjointFilteringOption"></li>
                        <li class="toggleOption" id="setOperatorFilteringOption"></li>
                        <li class="slideOption" id="nodeDegreeFilteringOption"></li>
                    </ul>
                </li>
                <li id="gravityOption"><a href="#">Gravity</a>
                    <ul class="toolTipMenu gravity">
                        <li class="slideOption" id="classSliderOption"></li>
                        <li class="slideOption" id="datatypeSliderOption"></li>
                    </ul>
                </li>
                <li id="export"><a href="#">Export</a>
                    <ul class="toolTipMenu export">
                        <li><a href="#" download id="exportJson">Export as JSON</a></li>
                        <li><a href="#" download id="exportSvg">Export as SVG</a></li>
                        <li class="option">
                            <div>
                                <form class="converter-form" id="url-copy-form">
                                    <label for="exportedUrl2">Export as URL:</label>
									<% if Request.QueryString("condivisione") = 1 then
											linkcondivisione = Session("urlmappa")
										else 
											linkcondivisione = Session("urlmappa")&"&condivisione=1&DB="&Session("DB")&"&Materia="&Session("ID_Materia")
										end if
										%>
										<% if Request.QueryString("condivisione") = 1 then %>
										<input type="text" id="exportedUrl2" placeholder="Clicca su Link" value="<%=Session("UrlCondivisione")%>">
										<span id="copyBt" title="Copy to Clipboard">Copy</span>
										<% else %>
                                    <input type="text" id="exportedUrl2" placeholder="Clicca su Link">
									<span id="copyBt" title="Copy to Clipboard">Link</span>
									<%end if %>
									
                                    <input type="hidden" id="exportedUrl" placeholder="an URL">
                                    
                                </form>
                            </div>
                        </li>
                    </ul>
                </li>
                <li id="select"><a href="#">Mappe</a>
                    <ul class="toolTipMenu select">
                       

                        <li class="option" id="converter-option">
                            <form class="converter-form" id="iri-converter-form">
                                <label for="iri-converter-input">Mappa personalizzata:</label>
                                <input type="text" id="iri-converter-input" placeholder="Inserisci IRI della mappa">
                                <button type="submit" id="iri-converter-button" disabled>Visualizza</button>
                            </form>
                            <div class="converter-form">
                                <input class="hidden" type="file" id="file-converter-input" autocomplete="off">
                                <label class="truncate" id="file-converter-label" for="file-converter-input">Seleziona file della mappa</label>
                                <button type="submit" id="file-converter-button" autocomplete="off" disabled>
								Upload
							</button>
                            </div>
                        </li>
                    </ul>
                </li>
                <li id="li_locationSearch"> <a title="Nothing to locate, enter search term." href="#" id="locateSearchResult">&#8853;</a></li>
                <li class="searchMenu" id="searchMenuId">
                    <input class="searchInputText" type="text" id="search-input-text" placeholder="Search">
                    <ul class="searchMenuEntry" id="searchEntryContainer">
                    </ul>
                </li>
                <li id="li_right" style="float:left"><a href="#" id="RightButton"></a></li>
                <li id="li_left" style="float:left"><a href="#" id="LeftButton"></a></li>

                </li>
            </ul>
        </nav>
    </main>
	
	<script>
	var utente = "mappeutenti/<%=CodiceAllievo%>";
	var stringalink = "<%=stringalink%>";
	var stringaproperty = "<%=stringaproperty%>";
	var stringastud = "<%=stringastud%>";
	</script>
	
    <script src="js/d3.min.js"></script>
    <script src="js/webvowl.js"></script>
    <script src="js/webvowl.app.js"></script>
    <script>
	
	
        window.onload = webvowl.app().initialize;	
		$(document).ready(function() { $("#gravityOption").click(); });
		
		var ref = window.localStorage.getItem("refresh");
		var dab = window.localStorage.getItem("dabutton");
		
		/*alert(ref);
		alert(dab);*/
		
		if(ref == 1 && dab != 1){
			//var z = confirm("Hai eseguito un refresh della pagina: in questo modo la mappa non viene ricaricata. La prossima volta utilizza il tasto a destra \"Ricarica mappa\", entrando in modalità collegamento.\nVuoi ricaricare la mappa ora?");
			z=true;
			if(z){
				window.location.href = "<%=Session("urlmappa")%>";
			}
		}
		window.localStorage.removeItem("refresh");
		window.localStorage.removeItem("dabutton");
		
    </script>
				  
	<script>
	<% if not Session("CodiceAllievo") = "" then %>
	
	if(collegamento==1){
		entracollegamento();
		}
	
	var vlink = stringalink.split(",");
	var vproperty = stringaproperty.split(",");
	var vstud = stringastud.split(",");
	
	/*alert(vlink[0]);
	alert(vproperty[0]);*/
	
	var primonodo = null;
	var secondonodo = null;
	var descrizione = null;
	
	function colleganodi(id){
	
		if(primonodo == null){
			primonodo = id;
			//alert("Seleziona il secondo nodo");
			var colleg = "<center>Seleziona il secondo nodo</center><br>";
			$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
		}else if(secondonodo == null){
			secondonodo = id;
		}
		
		
		if(primonodo != null && secondonodo != null){
			
			var colleg = "<center>Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.</center><br>"
			
			if(primonodo==secondonodo){
				alert("Il primo nodo non può essere uguale al secondo");
				annullamodifiche(0);
				
				$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
			}else{
			descrizione = prompt("Inserisci la descrizione del collegamento");
			if(descrizione=="" || descrizione==null || descrizione=="null" || descrizione=="undefined"){
				alert("Non puoi lasciare vuota la descrizione del collegamento");	
				annullamodifiche(0);
				$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
			}else{
			
					$.ajax({
					  method: "POST",
					  url: "../cNodi/inserisci_collegamento_interattivo.asp?Id_n1="+primonodo+"&Id_n2="+secondonodo+"&L1=1&L2=1&T2="+descrizione,
					  dataType: "html",
					  data: {  }
					}) /* .ajax */
					 .done(function( ans ) {
					 //alert(ans);
						if(ans.trim() == "inserito"){
							alert("Inserimento effettuato correttamente, ricarica la mappa per vedere le modifiche");
							var colleg = "<center>Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.</center><br>"
	$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
							primonodo = null;
							secondonodo = null;
							descrizione = null;
						}else{
							alert("Errore nella modifica");
							window.location.href = "<%=Session("urlmappa")%>";
						}
					 }); /* .done */
			}
			}
		}
	}
	
	<% if Session("Admin") = true then %>
	
	function modificalink(id){
		
		if(primonodo!=null){
			alert("Avevi già selezionato un nodo da collegare: operazione annullata");
			annullamodifiche(0);
		}else{
			var i = vproperty.indexOf(id);
			var nlink = vlink[i];
			
			var modifica = confirm("Vuoi modificare la descrizione del collegamento selezionato?");
				if(modifica){
					var testo = prompt("Inserisci la nuova descrizione")
					//window.location.href="../cNodi/modifica_collegamento_interattivo.asp?idlink="+nlink+"&T2="+testo;
					
					if(testo=="" || testo==null || testo=="null" || testo=="undefined"){
					alert("Non puoi lasciare vuota la descrizione del collegamento");
						annullamodifiche(0);
						
						var colleg = "<center>Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.</center><br>"
						$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
						
					}else{
					
					$.ajax({
					  method: "POST",
					  url: "../cNodi/modifica_collegamento_interattivo.asp?idlink="+nlink+"&T2="+testo,
					  dataType: "html",
					  data: {  }
					}) /* .ajax */
					 .done(function( ans ) {
					 //alert(ans);
						if(ans.trim() == "modificato"){
							alert("Modifica effettuata, ricarica la mappa per prenderne visione");
						}else{
							alert("Errore nella modifica");
							window.location.href = "<%=Session("urlmappa")%>";
						}
					 }); /* .done */
				
				}
					
				}else{
					var eliminazione = confirm("Vuoi eliminare il collegamento selezionato?");
			if(eliminazione){
				//window.location.href="../cNodi/elimina_collegamento_interattivo.asp?idlink="+nlink;
				
				$.ajax({
					  method: "POST",
					  url: "../cNodi/elimina_collegamento_interattivo.asp?idlink="+nlink,
					  dataType: "html",
					  data: {  }
					}) /* .ajax */
					 .done(function( ans ) {
					 //alert(ans);
						if(ans.trim() == "eliminato"){
							alert("Eliminazione effettuata, ricarica la mappa per vedere le modifiche");
						}else{
							alert("Errore nella modifica");
							window.location.href = "<%=Session("urlmappa")%>";
						}
					 }); /* .done */
					 
					 }else{
						annullamodifiche(0);
						}
						
						}
			
		}
		
		}
		
	<% else %>
	
		function modificalink(id){
		
		if(primonodo!=null){
			alert("Avevi già selezionato un nodo da collegare: operazione annullata");
			annullamodifiche(0);
		}else{
			var i = vproperty.indexOf(id);
			
			var username = "<%=Session("CodiceAllievo")%>";
			
			if(vstud[i] != username){
			
				alert("Non puoi modificare i nodi di altri utenti");
				annullamodifiche(0);
			
			}else{
			
			var nlink = vlink[i];
			
			var modifica = confirm("Vuoi modificare la descrizione del collegamento selezionato?");
				if(modifica){
					var testo = prompt("Inserisci la nuova descrizione")
					//window.location.href="../cNodi/modifica_collegamento_interattivo.asp?idlink="+nlink+"&T2="+testo;
					
					if(testo=="" || testo==null || testo=="null" || testo=="undefined"){
					alert("Non puoi lasciare vuota la descrizione del collegamento");
						annullamodifiche(0);
						
						var colleg = "<center>Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.</center><br>"
	$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
	
	
					}else{
					
					$.ajax({
					  method: "POST",
					  url: "../cNodi/modifica_collegamento_interattivo.asp?idlink="+nlink+"&T2="+testo,
					  dataType: "html",
					  data: {  }
					}) /* .ajax */
					 .done(function( ans ) {
					 //alert(ans);
						if(ans.trim() == "modificato"){
							alert("Modifica effettuata, ricarica la mappa per prenderne visione");
						}else{
							alert("Errore nella modifica");
							window.location.href = "<%=Session("urlmappa")%>";
						}
					 }); /* .done */
				
				}
					
				}else{
					var eliminazione = confirm("Vuoi eliminare il collegamento selezionato?");
			if(eliminazione){
				//window.location.href="../cNodi/elimina_collegamento_interattivo.asp?idlink="+nlink;
				
				$.ajax({
					  method: "POST",
					  url: "../cNodi/elimina_collegamento_interattivo.asp?idlink="+nlink,
					  dataType: "html",
					  data: {  }
					}) /* .ajax */
					 .done(function( ans ) {
					 //alert(ans);
						if(ans.trim() == "eliminato"){
							alert("Eliminazione effettuata, ricarica la mappa per vedere le modifiche");
						}else{
							alert("Errore nella modifica");
							window.location.href = "<%=Session("urlmappa")%>";
						}
					 }); /* .done */
					 
					 }else{
						annullamodifiche(0);
						}
						
						}
			
		}
		
		}
		
		}
	
	<% end if %>
		
	function entracollegamento(){
	collegamento = 1;
	var colleg = "<center>Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.</center><br>"
	$("#stringainfo").html(stringavisualizzazione+"<br>"+colleg);
	//alert("Seleziona il primo nodo o clicca sul collegamento da modificare/eliminare.\nPer annullare la selezione clicca su \"Annulla\".\nPer uscire dalla modalità collegamento clicca sul pulsante \"Esci\" in alto a destra.");
	}
	
	function escicollegamento(){
	collegamento = 0;
	$("#stringainfo").html(stringacollegamento);
	alert("Modalità di collegamento terminata");
	}
	
	function annullamodifiche(al){
	primonodo = null;
	secondonodo = null;
	descrizione = null;
	
	if(al==1){
		alert("Modifiche non ancora confermate annullate");
	}
	}
	
	<% end if %>
	
	$(window).on('beforeunload',function(){

		localStorage.setItem("refresh", 1);

	});
	
	function refresh(){
	
	localStorage.setItem("dabutton", 1);
	window.location.href = "<%=Session("urlmappa")%>";
	
	}
	
	$( document ).ready(function() {
	var t = setTimeout(function(){
		document.getElementById("compactnotationModuleCheckbox").click();
		clearTimeout(t);
		}, 300);
	});
	
	document.querySelector("#copyBt").onclick = function() {
	if($("#exportedUrl2").val().trim() == ""){
		var x = null;
		$.ajax({
			method: "POST",
			url: "condividi.asp",
			dataType: "html",
			data: { url: "<%=linkcondivisione%>" }
		}) /* .ajax */
		.done(function( ans ) {
			
			x = ans;
			$("#exportedUrl2").val(x);
			$("#copyBt").html("Copy");
			
		}); /* .done */
	
	}else{
	
	// selezione del contenuto
			document.querySelector("#exportedUrl2").select();
			// copia negli appunti
			document.execCommand('copy');
			
			}
	
	};
	
	</script>
	
</body>

</html>