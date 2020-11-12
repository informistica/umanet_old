<% 
		
		if session("Loggato")="" then
		
		  'Response.AddHeader "Refresh", "600"
		  Session("Loggato") =Request.Cookies("Dati")("Loggato")  
		  Session("Cognome")=Request.Cookies("Dati")("Cognome") 
		  Session("Nome")=Request.Cookies("Dati")("Nome") 
		  Session("CodiceAllievo")= Request.Cookies("Dati")("CodiceAllievo")
		  Session("Username")=Request.Cookies("Dati")("Username")   
		  Session("DataTest") = Request.Cookies("Dati")("DataTest")
		  Session("Id_Classe")= Request.Cookies("Dati")("Id_Classe") 
		  id_classe=Session("Id_Classe")
		  Session("cartella") = Request.Cookies("Dati")("cartella")
		  cartella=Session("cartella")
		  Session("Cartella") = Request.Cookies("Dati")("Cartella")
		  Session("CartellaAdmin")= Request.Cookies("Dati")("CartellaAdmin")
	      Session("In_Quiz")= Request.Cookies("Dati")("In_Quiz")
	      Session("CodAdmin")= Request.Cookies("Dati")("CodAdmin")
		  Session("Admin")= Request.Cookies("Dati")("Admin")
		   Session("Admin2")= Request.Cookies("Dati")("Admin2")
		  ' impostate in home.asp
		    if  Session("Admin2")<>"" then
		   Session("Admin")=true
		  end if
		  Session("stile")= Request.Cookies("Dati")("stile")
		  
		  
		'  response.write("<br>Loggato" & Session("Loggato"))
		 ' response.write("<br>CodiceAllievo" & Session("CodiceAllievo"))
		  '  response.write("<br>Username" & Session("Username"))
		  'response.write("<br>DataTest" & Session("DataTest"))
		  '  response.write("<br>Id_Classe" & Session("Id_Classe"))
		  'response.write("<br>cartella" & Session("cartella"))
		  	  
      Session("Materia") =Request.Cookies("Dati")("Materia")
	  Session("ID_Materia")=Request.Cookies("Dati")("ID_Materia") 
	  Session("ID_Matsint")= Request.Cookies("Dati")("ID_Matsint")   
	  Session("idxMat") =Request.Cookies("Dati")("idxMat")
	   
	  Session("DBCopiatestonline") = Request.Cookies("Dati")("DBCopiatestonline")
	  Session("DBClassifica") = Request.Cookies("Dati")("DBClassifica")
	  Session("DBForum") = Request.Cookies("Dati")("DBForum")
	  Session("DBLavagna") = Request.Cookies("Dati")("DBLavagna")
	  Session("DBDiario") = Request.Cookies("Dati")("DBDiario")
	  Session("DBDesideri") = Request.Cookies("Dati")("DBDesideri")
	  
	  Session("DB") = Request.Cookies("Dati")("DB")
		Session("DataCla") = Request.Cookies("Dati")("DataCla")
	    Session("DataCla2") = Request.Cookies("Dati")("DataCla2")
		Session("DataClaq") = Request.Cookies("Dati")("DataCla")
	    Session("DataClaq2") = Request.Cookies("Dati")("DataCla2")
		
		session("categoria")= Request.Cookies("Dati")("categoria")
		session("id_categoria")= Request.Cookies("Dati")("id_categoria")
		
		session("id_as")= Request.Cookies("Dati")("id_as")
		 
  
		Session.LCID=1040
	   response.write("<br>(Ripristinata sessione)" & Session("DBCopiatestonline"))
		
		end if   
		  
		  'Request.Cookies("Dati").Expires = Date() + 1%>