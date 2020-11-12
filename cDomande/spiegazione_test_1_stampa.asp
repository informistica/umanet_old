<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
     <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">
    
    


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
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
      
 
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
</head>
        

<%Function domandaplus()
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 Cartella=rsTabella.fields("Cartella")
	 Modulo=rsTabella.fields("ID_Mod")
	 'Paragrafo=rsTabella(15)
	 Paragrafo=rsTabella.fields("Titolo")
	' response.write("PARAGRAFO="&Paragrafo)
	 Id=rsTabella.fields("CodiceDomanda")
	'homesito="/anno_2010-2011_ITC/ECDL"
	 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	'response.write(url)
	objTextFile.Close
End Function %>
 
 
 <%Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    

  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

  
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
'homesito="/anno_2010-2011_ITC/ECDL"
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceTest=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
   Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if

  
  
  
  
  Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  		'response.write("Stato"&stato)				
%>

 <%'response.write("Stati :  " & stato & " " & stato0) 
 if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY class='theme-<%=session("stile")%>' onLoad="showText2();"> </BODY>
  <% else %>
     <body>
  <% end if %>
	 
     
        <% 
		
   'per il copia incolla
  ' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabella.fields("CIAbilitato") 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close
	  
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		 
  		<!-- #include file = "../service/controllo_sessione.asp" -->
	 
          
         
	 
    
    
    
    
	  	<%=Capitolo & ":"&Paragrafo%>
                        <% if Sottoparagrafo<>"" then
						response.write("/"&Sottoparagrafo)
						end if%>
                        
                        


                      
                      
 
  
                            <br>
                            
                                                  <%   
  
' essendo costante per tutte le query ...
'Visualizza tutte le domande di un modulo
costQuerySQL1 =	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE   Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and ID_Predomanda not in (Select Id_Predomanda from Domande WHERE Domande.Id_predomanda<>0 and Id_Mod='" &Modulo  & "') " &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" 
	'Visualizza tutte le domande di un modulo dello stud loggato 
	costQuerySQL11 =	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and Domande.Id_predomanda<>0 and ID_Predomanda not in (Select Id_Predomanda from Domande WHERE  Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "') " &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Id_Stud,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" 
	 

'Visualizza tutte le domande di un paragrafo
costQuerySQL2 =	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and Domande.Id_Arg='" & CodiceTest & "' and ID_Predomanda not in (Select Id_Predomanda from Domande WHERE Domande.Id_predomanda<>0 and Domande.Id_Arg='" & CodiceTest & "' and Id_Mod='" &Modulo  & "') " &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" 
	 'Visualizza tutte le domande di un paragrafo solo quelle dello stude loggato
costQuerySQL22 =	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and Domande.Id_predomanda<>0 and Domande.Id_Arg='" & CodiceTest & "' and ID_Predomanda not in (Select Id_Predomanda from Domande WHERE Domande.Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "') " &_
	" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Id_Stud,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" 
	 
	 
	 
' sottoparagrafo tutto 
costQuerySQL3=" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE  Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and (((Domande.Id_Sottoparagrafo)='" & CodiceSottoPar & "') AND ((Domande.[ID_Predomanda]) Not In (Select Id_Predomanda from Domande WHERE Domande.Id_predomanda<>0 and Domande.Id_Sottoparagrafo='" & CodiceSottoPar & "' and Id_Mod='" &Modulo  & "')))"&_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod, Domande.Tipo, Domande.In_Quiz, Domande.Cartella, Domande.Id_Predomanda, Domande.Id_Sottoparagrafo, Domande.Multiple, Domande.VF, Domande.Img"

' sottoparagrafo solo stud loggato
costQuerySQL33=" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Moduli.ID_Mod = Domande.Id_Mod) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
" WHERE Domande.Segnalata=0 and Domande.Lingua='"&Lingua&"' and (((Domande.Id_Sottoparagrafo)='" & CodiceSottoPar & "') AND ((Domande.[ID_Predomanda]) Not In (Select Id_Predomanda from Domande WHERE Domande.Id_predomanda<>0 and Domande.Id_Sottoparagrafo='" & CodiceSottoPar & "' and Id_Stud='"&Session("CodiceAllievo")&"') and  Id_Mod='" &Modulo  & "')))"&_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod, Domande.Tipo, Domande.In_Quiz, Domande.Cartella, Domande.Id_Predomanda, Domande.Id_Sottoparagrafo, Domande.Multiple, Domande.VF, Domande.Img " 
	 
	
	 
	
 
 
 
if (clng(Stato)=0) or (clng(Stato0)=0) then 
' 'Definzione codice SQl della query per ricercare le domande del paragrafo 
	
   if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le domande del PARAGRAFO altrimenti solo quelle dello       studente loggato  
  
	QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" &_
	costQuerySQL2 &_
	" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Domande.In_Quiz<>0" &_   
	" ORDER BY Paragrafi.ID_Paragrafo,Domande.VF,Domande.Multiple;"
   else
	QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Id_Stud,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" &_
	costQuerySQL22 &_
	" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Domande.In_Quiz<>0 and Domande.Id_Stud='"& Session("CodiceAllievo")& "'" &_
	"' ORDER BY Paragrafi.ID_Paragrafo,Domande.VF,Domande.Multiple;"
   end if 



else 
  if (clng(Stato)=1) or (clng(Stato0)=1) then 
  if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le domande del MODULO altrimenti solo quelle dello       studente loggato  
   	
   						'0					1				2					3			
	QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" &_
	costQuerySQL1 &_
" HAVING Moduli.ID_Mod='" & Modulo & "' AND Domande.In_Quiz<>0" &_ 
" ORDER BY Paragrafi.ID_Paragrafo,Domande.VF,Domande.Multiple;"
 else
   QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Id_Stud,Domande.Cartella,Domande.Id_Predomanda, Domande.Multiple, Domande.VF, Domande.Img" &_
	costQuerySQL11 &_
" HAVING Moduli.ID_Mod='" & Modulo & "' AND Domande.In_Quiz<>0 and Domande.Id_Stud='"& Session("CodiceAllievo") & "'" &_
"' ORDER BY Paragrafi.ID_Paragrafo,Domande.VF,Domande.Multiple;"
 end if
 
  else  'if (clng(Stato)=2) or (clng(Stato0)=2) then
  ' sottoparagrafo
  
		  if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le domande del MODULO altrimenti solo quelle dello       studente loggato  		
			QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Cartella,Domande.Id_Predomanda,Domande.Id_Sottoparagrafo, Domande.Multiple, Domande.VF, Domande.Img" &_
			costQuerySQL3 &_
		" HAVING Moduli.ID_Mod='" & Modulo & "'and  Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Domande.In_Quiz<>0" &_ 
		" ORDER BY Allievi.Cognome,Domande.Multiple, Domande.VF, Domande.Img;"
		 else
		   QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Domande.Quesito, Domande.CodiceDomanda, Moduli.ID_Mod,Domande.Tipo,Domande.In_Quiz,Domande.Id_Stud,Domande.Cartella,Domande.Id_Predomanda,Domande.Id_Sottoparagrafo, Domande.Multiple, Domande.VF, Domande.Img" &_
			costQuerySQL33 &_
		" HAVING  Moduli.ID_Mod='" & Modulo & "' and Paragrafi.ID_Paragrafo='" & Codice_Test & "' AND Domande.In_Quiz<>0 and Domande.Id_Stud='"& Session("CodiceAllievo") & "'" &_
		"' ORDER BY Allievi.Cognome,Domande.Multiple, Domande.VF, Domande.Img;"
		 end if
  
  
  end if
end if    
  '  response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 
  capitolo=rsTabella("Titolo")
	   titoloParagrafo=rsTabella(0)
' la inserisco per i moduli condivis, devoprendere la cartella dalla domanda anziche dalla classe
cartella=rsTabella("Cartella")      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then %>
  <div class="alert alert-error">
                    Domande del Test non ancora disponibili!
                   
 </div>
 <%else%>
 
 
 <%
 
  i=1 'inizializza la variabile i (contatore delle domande)
  Do until rsTabella.EOF
  	' if i>1 then
  	if strcomp(titoloParagrafo,rsTabella(0))<>0 then
	    titoloParagrafo=rsTabella(0)%>
		 <b><center> <font size="+2">  <%=rsTabella(0)%></font> </center></b>
	<hr>
	<%
	'else 
	
	'end if
	
	 if StrComp(Sottoparagrafo, rsTabella("SotPar")) <> 0 then
			  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
			   Sottoparagrafo=rsTabella("SotPar")
                %>
                <b><center> <font size="+2"><%=rsTabella("SotPar")%></font> </center></b> 
			 <%end if%>    
	
  <% else%>
		<% if i=1 then%>
		   <b><center> <font size="+2"><%=titoloParagrafo%></font> </center></b> 
		   <%end if%>
     
	<% end if	 
 
    ID=rsTabella(4)
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
   url=Replace(url,"\","/")
 
              
 
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
 
sReadAll = objTextFile.ReadAll
'sReadAll = url
objTextFile.Close   ' la soluzione seguente la rimuovo e dirò di copiare ed incollare la domanda plus nella spiegazione
' così da avere il livello di apprendimento comprensibile , diversamente dovrei prevedere il modo di far apparire il testo della domanda plus 
' anche nell'approfondimento di fine quiz.
'if clng(rsTabella.fields("Tipo"))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
'	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
'		url=Replace(url,"\","/")
'		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
'		sReadAll1 = objTextFile.ReadAll
'		objTextFile.Close
'end if
			 
%>
                              
  
  
    
  
 <table  class="table table-hover table-nomargin table-condensed table-bordered">
		 
        <tr>
			<th  class='hidden-350' width='75%'>
           <%
		 '  if stato<>0 then
		'    response.write(rsTabella("Titolo")&": ")
		'  end if
		  
            if rsTabella("VF")=1 then
			 
			  response.write("Vero/Falso")
			  else
			  	  if rsTabella("Multiple")=1 then
					  response.write("Risposta multipla")
				   else
				      response.write("Risposta singola")
				   end if
			  
			  end if
			  
			
			%>
              <a title="<%=response.write(rsTabella("CodiceDomanda"))%>">.</a>
            </th>
			 
			<th>  </th>
		</tr>
		<tr>
			<td colspan=3>
			<p align="center"><b><%=rsTabella(3)%></b></td>
			 
		</tr>
		
		<% if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
	    <tr><td colspan="3"><p align="center">
 <textarea rows="<%=1+round((len(domandaplus()))/50)%>" name="TestoDomandaPlus0" value="ciao" class="input-block-level"><%
			 
			 
			 Response.write(domandaplus())%> </textarea><br></td></tr><br>
        <%end if %>
   
		<tr>
			<td colspan=3>
			
			<p align="center">
			 
			 
			<% Response.write(sReadAll)%>  
             </p>
		      </td>
		 
		</tr>
 
     </tbody>
	</table>
    
    
    
	<br>
<%    

       i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
							 
							</div>
						</div>
					</div>
				</div>								 
	 
	 
				 
			         
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

