 <%
            response.write("<br><b>"&rsTabella("Cognome") & " " &rsTabella("Nome")&" : </b>")
            If (fso.FileExists(urlRisposte)) Then    ' se esiste e quindi lo stud ha consegnato
                ' creo file per salvare la correzione
                urlRisorsaRisposteCorrezione=paragrafo&"_correzione_"&rsTabella("CodiceAllievo")&".xml"
                urlCorrezione=urlRisRisposte&urlRisorsaRisposteCorrezione
                urlCorrezione=Replace(urlCorrezione,"\","/")
        
                 Set objFileCorrezioni = fso.CreateTextFile(urlCorrezione, True)
	             objFileCorrezioni.WriteLine("<Correzioni>")

                ' apro file delle risposte
               ' response.write(urlRisposte)
                objXMLDocR.load urlRisposte  
                Set RootR = objXMLDocR.documentElement
                Set NodeListR = RootR.getElementsByTagName("Domanda")
                totale=0
                For n = 0 to NodeListR.length -1
                    Set IdPrefrase = objXMLDocR.getElementsByTagName("IdPrefrase")(n)
                    Set RispostaM = objXMLDocM.getElementsByTagName("Risposta")(n)
                    Set RispostaR = objXMLDocR.getElementsByTagName("Risposta")(n)
                    Set TestoM = objXMLDocM.getElementsByTagName("Testo")(n)
                    'risposta ideale
                    readAll=Replace(RispostaM.text,".","")
                    readAll=Replace(readAll,",","")
                    readAll=Replace(readAll,chr(13)," ")
                    readAll=Replace(readAll,vbCr," ")
                    readAll=Replace(readAll,vbLf," ")
                    risposta_ideale_pre = Split(readAll," ")
                    'risposta data
                    readAll=Replace(RispostaR.text,".","")
                    readAll=Replace(readAll,",","")
                    readAll=Replace(readAll,vbCr," ")
                    readAll=Replace(readAll,vbLf," ")
                    risposta_pre = Split(readAll," ")
                    
                    ri=0
                    for each x in risposta_ideale_pre
                            if (len(x)>5) then
                                
                                risposta_ideale(ri)=Replace(Lcase(Trim(x)),","," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),chr(13)," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),vbCr," ")' ***** FORSE RISOLVE IL
								risposta_ideale(ri)=Replace(risposta_ideale(ri),vbLf," ")											 
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"."," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"-","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"-"," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"("," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),")"," ")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"perche`","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"quindi","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"quando","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"infatti","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"dell`","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"l`","")
                                risposta_ideale(ri)=Replace(risposta_ideale(ri),"d`","")
                                ' per il problema dei token formati da due parti
                                    risposta_ideale_pre2 = Split(risposta_ideale(ri)," ")
                                    if ubound(risposta_ideale_pre2)>0 then
                                        for each y in risposta_ideale_pre2
                                            if y<>"" then
                                            risposta_ideale(ri)=y
                                            ri=ri+1
                                            end if
                                        next 
                                    end if        
                                    ri=ri+1
        
                            end if
                    next
                            
                    r=0
                    for each x in risposta_pre
                            if (len(x)>5) then
                                risposta(r)=Replace(Lcase(Trim(x)),",","")
                                risposta(r)=Replace(risposta(r),"."," ")
                                risposta(r)=Replace(risposta(r),"-","")
                                risposta(r)=Replace(risposta(r),chr(13)," ")
                                'risposta(r)=Replace(risposta(r),"<br>"," ")
                                risposta(r)=Replace(risposta(r),vbCr," ")' ***** FORSE RISOLVE IL
								risposta(r)=Replace(risposta(r),vbLf," ")
                                risposta(r)=Replace(risposta(r),"-"," ")
                                risposta(r)=Replace(risposta(r),")"," ")
                                risposta(r)=Replace(risposta(r),"("," ")
                                risposta(r)=Replace(Lcase(Trim(x)),";"," ")
                                risposta(r)=Replace(Lcase(Trim(x)),":"," ")
                                risposta(r)=Replace(risposta(r),"perche`","")
                                risposta(r)=Replace(risposta(r),"quindi","")
                                risposta(r)=Replace(risposta(r),"quando","")
                                risposta(r)=Replace(risposta(r),"infatti","")
                                risposta(r)=Replace(risposta(r),"dell`","")
                                risposta(r)=Replace(risposta(r),"l`","")
                                risposta(r)=Replace(risposta(r),"d`","")
                               'response.write("?"&risposta(r))
                               'if (strcomp(rtrim(ltrim(risposta(r))),"trasmissivo rappresentare")=0) then
                                '  response.write("*****ECCOLO!")
                                'end if
                                 ' per il problema dei token formati da due parti
                                risposta2 = Split(risposta(r)," ")
                                    if ubound(risposta2)>0 then
                                ' response.write("<br>SONO DENTRO *****")
                                        for each y in risposta2
                                            if y<>"" then
                                            risposta(r)=y
                                            r=r+1
                                            end if
                                        next 
                                    end if        



                                r=r+1
                            end if
                    next
               


                    Response.write("<br><br>"&(n+1) & ") "&TestoM.text)
                    objFileCorrezioni.WriteLine("<Domanda>")
                    objFileCorrezioni.WriteLine("<Testo>")
                    objFileCorrezioni.WriteLine("   "&TestoM.text)
                    objFileCorrezioni.WriteLine("</Testo>")
                    
                    okTotale=0
                    
                    'per ogni parola del modello vedo se appartiene alla risposta data
                    for j=0 to ri-1
                        trovata=0
                        okParziale=0
                        if risposta_ideale(j)<>"" then
                            objFileCorrezioni.WriteLine("<Modello>")
                            for i=0 to r-1
                                ' controllo puntale
                                '  if (strcomp(trim(risposta(i)),trim(risposta_ideale(j)))=0) and (strcomp("",trim(risposta_ideale(j)))<>0) then
                                ' accetto valida una corrispondenza con il 60% della risposta, vedo se il 60% della risposta ideale è contenuta nella risp data 
                                if (instr(risposta(i),left(trim(risposta_ideale(j)),cint(len(risposta_ideale(j))*0.6)))<>0) and (strcomp("",risposta_ideale(j))<>0) then 
                                    okTotale=okTotale+1
                                    okParziale=okParziale+1
                                    trovata=1
                                    risposta(i)=""
                                    Exit For                    
                                end if
                            next  
                            if trovata=1 then
                                    
                                    objFileCorrezioni.WriteLine("   "&risposta_ideale(j)&"(1)")
                                    risposta_ideale(j)=""
                            Else
                                objFileCorrezioni.WriteLine("   "&risposta_ideale(j)&"(0)")
                            end if  
                            objFileCorrezioni.WriteLine("</Modello>")
                        end if
                        ' risposta_ideale(j)=""
                    next
            
                    
                    fitness=(okTotale/ri)*100
                    
                    objFileCorrezioni.WriteLine("<Corrispondenza>")
                    objFileCorrezioni.WriteLine("   "&Fix(fitness))
                    'objFileCorrezioni.WriteLine("   "&okTotale &"/" & ri)
                    objFileCorrezioni.WriteLine("</Corrispondenza>")
                    objFileCorrezioni.WriteLine("</Domanda>")
                '	response.write("<br>Trovate "&ok&" corrispondenze su "&ri&" parole") 
                    response.write("<br><b>Corrispondenza "& Fix(fitness) &"%</b>") 
                '	response.write("<br>Rapporto len risposta/modello= "& r/ri &"%<br>") 
                    totale=totale+fitness

                Next
                    media=totale/NodeListR.length
                    Response.write("<br><i class='icon-bar-chart'></i><b>Sentiment : "&Fix(media)&"%</b><br>")
                    objFileCorrezioni.WriteLine("<Sentiment>")
                    objFileCorrezioni.WriteLine("   "&Fix(media))
                    objFileCorrezioni.WriteLine("</Sentiment>")
                    objFileCorrezioni.WriteLine("</Correzioni>")
                    objFileCorrezioni.Close
                Else ' il file non esiste, lo stud non ha fatto la verifica
                    response.write("<code>Non ha consegnato</code>")
            End If
            response.write("<hr>")
    
 %>