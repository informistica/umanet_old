<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script type="text/javascript" src="../../js/google.js"></script><title>Inserisci modulo </title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 <meta charset="UTF-8">

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>

    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>


<!-- Easy pie -->
    <link rel="stylesheet" href="../../css/plugins/easy-pie-chart/jquery.easy-pie-chart.css">
	<script src="../../js/plugins/easy-pie-chart/jquery.easy-pie-chart.min.js"></script>

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
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />

 

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%

		 Response.Buffer=True 
   
   Dim cartelle(4)
   cartelle(0)="_Domande"
   cartelle(1)="_Frasi"
   cartelle(2)="_Nodi"
   cartelle(3)="_Spiegazioni"
   cartelle(4)="_Esercizi"

		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo

' QuerySQL="SELECT Verifica from Paragrafi where ID_Paragrafo='"&CodiceTest&"'"
 'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'Verifica=rsTabella(0) 

 ID_Mod=Request("txtID_Mod")
  Titolo=Request("txtTitolo")
	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserimento modulo </h1>

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
							<a href="#">Admin</a>
							<i class="icon-angle-right"></i>
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
				        <h3> <i class="icon-reorder"></i> <%=Titolo%>
                         </h3>
			          </div>
				      <div class="box-content">
							<div class="row-fluid">
								<div class="span12">
				
									 <%

function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"e"&Chr(96))
'   sAns = Replace(sAns,chr(133),"a"&Chr(96))
'   Ans = Replace(sAns,chr(160),"a"&Chr(96))
'   sAns = Replace(sAns,chr(232),"e"&Chr(96))
'   sAns = Replace(sAns,chr(233),"e"&Chr(96))
'   sAns = Replace(sAns,chr(236),"i"&Chr(96))
'   sAns = Replace(sAns,chr(237),"i"&Chr(96))
'   sAns = Replace(sAns,chr(242),"o"&Chr(96))
'   sAns = Replace(sAns,chr(243),"o"&Chr(96))
'   sAns = Replace(sAns,chr(249),"u"&Chr(96))
'   sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
  sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
  'sAns=  Replace(sAns,"&","e") 
  sAns=  Replace(sAns,"/","-") 
  sAns=  Replace(sAns,"\","-") 
  sAns=  Replace(sAns,"?",".") 
  sAns=  Replace(sAns,"*","x") 
  sAns=  Replace(sAns,"<","_")
  sAns=  Replace(sAns,">","_") 
  
  ReplaceCar = sAns
end function



'*****Inserisco il MODULO
'  QuerySQL="  INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione,URL_OL)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & Classe & "', " & cint(right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))) &",'"&urlCopertina&"';" ' dave errore nel cint
  cartella=request.querystring("cartella")
  Id_Classe=Request.querystring("Id_Classe")
  umanet=Request.querystring("umanet")
  if umanet<>"" then
     QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&cartella&"';"
  else
   QuerySQL="SELECT max(posizione) FROM MODULI_NOT_UMANET where Cartella='"&cartella&"';"
end if
 response.write(QuerySQL) 
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1

Titolo = ReplaceCar(Titolo) 
  	    
  QuerySQL="INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione,URL_OL, Visibile)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & cartella & "', " & posizione &",'"&urlCopertina&"',1;"  
   response.write("<br>"&QuerySQL) 
   ConnessioneDB.Execute QuerySQL %>
	
   <%
	
    Set fso = CreateObject("Scripting.FileSystemObject") 
    for i=0 to 4 
		url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/"&Id_Mod&cartelle(i) 
		url=Replace(url,"\","/")
		 response.Write("<br>"&url)
		if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste gi�.<br>")
		else
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/Img")
			response.Write( "La cartella " & url&"/Img" & " � stata creata.<br>") 
		end if
    next 
	' creo la cartella per il modulo dentro la cartella risorse del corso  
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste gi�.<br>")
		else
		    response.Write( "Creazione della cartella :" & url&"/img" & "....") 
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/img")
	     	 
		 
		end if

		' creo la cartella per il modulo dentro la cartella verifiche del corso  
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			 response.Write( "<br>La cartella " & url & " esiste gi�.<br>")
		else
		    response.Write( "<br>Creazione della cartella :" & url) 
			fso.CreateFolder (url) 
	 
		end if
	
 'response.write(url)
'---Fine inserimento titolo  modulo
 '*** inizio paragrafi sottoparagrafi
 
    strText = Request.Form("MyTextArea")
    arrLines = Split(strText, vbCrLf)
    i=0
    j=0
    cont=0
    
    response.write("Numero elementi-1:"&ubound(arrLines))
 



        for a=0 to ubound(arrLines)-1
            s=split(arrLines(a),"&&&")
            if ubound(s)>0 then
                titpar=s(1)
				titpar = ReplaceCar(titpar)' sostituisco gli apici " con l'apice singolo
			 
                'inserisco titolo paragrafo
                i=i+1
                 Id_Paragrafo=ID_Mod&"_"&i
                 QuerySQL="  INSERT INTO Paragrafi (ID_Paragrafo, Titolo,Posizione,URL_L,URL_O)  SELECT '" & Id_Paragrafo  & "','" & titpar & "','" & i & "','"& ulrRisorsa1 &"','"& ulrRisorsa1 &"';"
                 response.write("<br>"&QuerySQL&"<br>") 
            	 ConnessioneDB.Execute QuerySQL 
				   QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod & "', '" & Id_Paragrafo  & "';"
				   ConnessioneDB.Execute QuerySQL 
			      response.Write(QuerySQL&"<br>")
                j=1
                a=a+1
                
            end if   
                'inserisco stottoparagrafo
                Sottoparagrafo=arrLines(a)
				Sottoparagrafo = ReplaceCar(Sottoparagrafo) 
                urlSotPar=arrLines(a+1)
                ID_Sottoparagrafo=Id_Paragrafo&"_"&j

                QuerySQL="  INSERT INTO Sottoparagrafi (ID_Sottoparagrafo, Titolo,Posizione,URL) SELECT '" & ID_Sottoparagrafo  & "','" &  Sottoparagrafo & "'," & j & ",'"&urlSotPar&"';"
				   response.Write("<br>"&QuerySQL&"<br>" )
				   ConnessioneDB.Execute QuerySQL 
				    
			    QuerySQL="  INSERT INTO ParagrafiSottoparagrafi (Id_Paragrafo,Id_Sottoparagrafo) SELECT '" & Id_Paragrafo  & "','" & ID_Sottoparagrafo & "';"
				   response.Write(QuerySQL&"<br>" )
				   ConnessioneDB.Execute QuerySQL   

            

           j=j+1
           a=a+1
        next 
   %>						<br><hr>			 
						<div class="alert alert-success">
                     	<b><%=response.write("Inserimento effettuato correttamente")%></b>
                     </div>		 
									 
								 

								</div>
							</div>
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->
 
   
 

	</body>

 </html>
