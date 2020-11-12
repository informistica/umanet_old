
<%


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")



		%>
        <!-- #include file = "../../var_globali.inc" -->
 		<!-- #include file = "../../stringhe_connessione/stringa_connessione.inc" -->






<h4>&nbsp; <i class="icon-group"></i>  Modifica scadenze per la classe</h4>

      <%



 if session("admin") = true then %>

 <form method="POST" action="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&modifica_scadenze_classe=1" name="frmDocument1"  class="form-horizontal form-bordered form-validate">
            <%

            QuerySQL="SELECT count(*)  FROM Eccezioni_Frasi  WHERE  Id_Classe='"&Session("Id_Classe")&"';"

  						 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'response.write(QuerySQL)
  						 eccFrasi = rsTabella(0)
  						 QuerySQL="SELECT count(*)  FROM Eccezioni_Nodi  WHERE  Id_Classe='"&Session("Id_Classe")&"';"
  						  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  						 eccNodi=rsTabella(0)
  						 QuerySQL="SELECT count(*)  FROM Eccezioni_Domande  WHERE Id_Classe='"&Session("Id_Classe")&"';"
            eccDomande=rsTabella(0)

            ' 25/03/19 da implementare il resetta eccezioni per classe '
            %>

                                    <div class="control-group">
                                       <label for="textfield" class="control-label">Frasi </label>
                   <div class="controls">
                     <input type="text" style="height: auto;" disabled class="input-mini" value="<%=eccFrasi%>" rel="tooltip"  title="Numero di proroghe concesse" >
                    &nbsp; &nbsp;
                   <a href="2_rimuovi_eccezioni.asp?frasi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta e Rimuovi Eccezioni" rel="tooltip"  title="Consulta ed Elimina" ></a>
                   &nbsp; &nbsp;
                   <a href="studente_domande_include/2_resetta_eccezioni.asp?frasi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Resetta" rel="tooltip"  title="Azzera" ></a>
                   </div>
                 </div>

                                    <div class="control-group">
                                       <label for="textfield" class="control-label">Domande </label>
                   <div class="controls">
                     <input type="text" style="height: auto;" disabled class="input-mini"  value="<%=eccDomande%>" rel="tooltip"  title="Numero di proroghe concesse">
                                            &nbsp; &nbsp;
                   <a href="2_rimuovi_eccezioni.asp?domande=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta e Rimuovi Eccezioni" rel="tooltip"  title="Consulta ed Elimina" ></a>
                   &nbsp; &nbsp;
                   <a href="studente_domande_include/2_resetta_eccezioni.asp?domande=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Resetta" rel="tooltip"  title="Azzera" ></a>
                   </div>
                 </div>

                                    <div class="control-group">
                                       <label for="textfield" class="control-label">Nodi</label>
                   <div class="controls">
                     <input type="text" style="height: auto;" class="input-mini" disabled  value="<%=eccNodi%>" rel="tooltip"  title="Numero di proroghe concesse">
                                            &nbsp; &nbsp;
                   <a href="2_rimuovi_eccezioni.asp?nodi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta e Rimuovi Eccezioni" rel="tooltip"  title="Consulta ed Elimina" ></a>
                   &nbsp; &nbsp;
                   <a href="studente_domande_include/2_resetta_eccezioni.asp?nodi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Resetta" rel="tooltip"  title="Azzera" ></a>
                   </div>


                 </div>
                                   <br />
                                    <input type="submit" class="btn" value="Modifica scadenze" />
                                 <br /><br>


               </form>

 <% else

   Response.Redirect "../../../../../"

 end if

   %>
