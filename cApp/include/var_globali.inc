<% dim inizio_anno,homesito,DataClaFine, fine_anno 
homeserver="c:/inetpub/umanetroot"
homesito="/expo2015Server/UECDL"  
dominio="www.umanetexpo.net" 
inizio_anno="12/09/2013"
fine_anno="08/06/2016"
DataClaDefault="12/09/2013" 'inizio a/s
'DataCla2Default="08/06/2014" ' fine a/s
' DataCla2Default=left(FormatDateTime(now(),0),10)
	giorno=day(date())
	mese= month(date())
	anno=year(date())
	if len(giorno)=1 then
	   giorno="0" &	day(date())
	end if
	if len(mese)=1 then
	   mese="0" &	month(date())
	end if
	DataCla2Default=giorno&"/"&mese&"/"&year(date())
	
	
DataClaFine="12/12/2112"
datafine="12/12/2112"


Const pwdAdmin="300aa334c9bb655d5c3ecd644a06612404d2020f882462d34b6a9f3e33644bab"
              
Const AdminPassword = "maurosho" ' per la chat 
Const codAdmin="informistica"
Const eMailAdmin="prof.spinarelli.mauro@gmail.com"

dim vetCodici(),vetImg(500,4)


%>