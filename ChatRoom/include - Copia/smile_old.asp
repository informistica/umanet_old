<%function SMILEFormat(sInput)
' non utilizzata da nessuno 
	dim sAns
	
	
		'Smilies
		sAns = Replace(sInput, ":huh?", "<img src=smilies/on_1.gif align=absmiddle>")
		
	'	 QuerySQL="Select * from TUTTESMILES where ID_Categoria=1 order by Posizione;"
'   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
'   rsTabellaS.movefirst
'   do while not rsTabellaS.eof 
'        sAns = Replace(sAns,rsTabellaS("Codice"), "<img src=" & rsTabellaS("Cartella_Cat")&"/"&rsTabellaS("Url")&" align=absmiddle>")
'  	    rsTabellaS.movenext
'   loop	
	sAns = Replace(sInput, ":huh?", "<img src=smilies/on_1.gif align=absmiddle>")
	sAns = Replace(sAns, ":s", "<img src=smilies/on_2.gif align=absmiddle>")
	sAns = Replace(sAns, ":PP", "<img src=smilies/on_3.gif align=absmiddle>")
	sAns = Replace(sAns, "}:)", "<img src=smilies/on_4.gif align=absmiddle>")
	sAns = Replace(sAns, ":DD", "<img src=smilies/on_5.gif align=absmiddle>")
	sAns = Replace(sAns, "}:|", "<img src=smilies/on_6.gif align=absmiddle>")
	sAns = Replace(sAns, ":)", "<img src=smilies/on_7.gif align=absmiddle>")
	sAns = Replace(sAns, ":oops", "<img src=smilies/on_8.gif align=absmiddle>")
	sAns = Replace(sAns, ";)", "<img src=smilies/on_9.gif align=absmiddle>")
	sAns = Replace(sAns, ":pff", "<img src=smilies/on_10.gif align=absmiddle>")
	sAns = Replace(sAns, ":_P", "<img src=smilies/on_11.gif align=absmiddle>")
	sAns = Replace(sAns, ":0", "<img src=smilies/on_12.gif align=absmiddle>")
    
	sAns = Replace(sAns, ":b;", "<img src=smilies/on_13.gif align=absmiddle>")
    sAns = Replace(sAns, ":xx", "<img src=smilies/on_14.gif align=absmiddle>")
	sAns = Replace(sAns, ":gg", "<img src=smilies/on_15.gif align=absmiddle>")
	sAns = Replace(sAns, ":nn", "<img src=smilies/on_16.gif align=absmiddle>")
	sAns = Replace(sAns, ":pp", "<img src=smilies/on_17.gif align=absmiddle>")
	sAns = Replace(sAns, ":kk", "<img src=smilies/on_18.gif align=absmiddle>")
	sAns = Replace(sAns, ":yy", "<img src=smilies/on_19.gif align=absmiddle>")
	sAns = Replace(sAns, ":zz", "<img src=smilies/on_20.gif align=absmiddle>")
	
	
	' Connessioni
	
QuerySQL="Select * from TUTTESMILES where ID_Categoria<>1 order by Posizione;"
   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
    rsTabellaS.movefirst
   do while not rsTabellaS.eof 
        sAns = Replace(sAns,rsTabellaS("Codice"), "<img src=../img_social/" & rsTabellaS("Cartella_Cat")&"/"&rsTabellaS("Url")&" align=absmiddle>")
  	    rsTabellaS.movenext
   loop	
   set rsTabellaS=nothing
 
	'sAns = Replace(sInput, ":;0_00", "<img src=../img_social/connessioni_percezioni/0_0_MatrixOmino_2.jpg  align=absmiddle>")
'	 
'	sAns = Replace(sAns, ":;0_01", "<img src=../img_social/connessioni_percezioni/0_1_incredibile_solo.jpg align=absmiddle>")
'	sAns = Replace(sAns, " :;0_02", "<img src=../img_social/connessioni_percezioni/0_2_vedoilsole.jpg  align=absmiddle>")
'	sAns = Replace(sAns, ":;0_03", "<img src=../img_social/connessioni_percezioni/0_3_NavigaOcchioSi.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_04", "<img src=../img_social/connessioni_percezioni/0_4__MondoCoerenzaVerde.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_05", "<img src=../img_social/connessioni_percezioni/0_5_LampadinaAccesa.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_06", "<img src=../img_social/connessioni_percezioni/0_6_TestaAccesa.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_07", "<img  src=../img_social/connessioni_percezioni/0_7_vedopioggia.jpg  align=absmiddle>")
'	sAns = Replace(sAns, ":;0_08", "<img src=../img_social/connessioni_percezioni/0_8_NavigaOcchioNo.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_09", "<img src=../img_social/connessioni_percezioni/0_9_MondoParadossoRosso.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_10", "<img  src=../img_social/connessioni_percezioni/0_10_LampadinaSpenta.jpg align=absmiddle>")
'	sAns = Replace(sAns, ":;0_11", "<img  src=../img_social/connessioni_percezioni/0_11_TestaSpenta.jpg  align=absmiddle>")
'    
'
'  
'
'	sAns = Replace(sAns, "[B]", "<strong>", 1, -1, 1)
'	sAns = Replace(sAns, "[/B]", "</strong>", 1, -1, 1)
'	sAns = Replace(sAns, "[I]", "<em>", 1, -1, 1)
'	sAns = Replace(sAns, "[/I]", "</em>", 1, -1, 1)
'	sAns = Replace(sAns, "[U]", "<u>", 1, -1, 1)
'	sAns = Replace(sAns, "[/U]", "</u>", 1, -1, 1)
'	
	
	SMILEFormat = sAns
end function%>

