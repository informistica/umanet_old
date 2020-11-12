<% dim inizio_anno,homesito,DataClaFine, fine_anno
homeserver="c:/inetpub/umanetroot"
homeserverlocal="https://localhost:8081"
homesito="/expo2015Server/UECDL"
homesitoecdl="/expo2015/"
dominio="www.umanetexpo.net"
protocollo="https://"
inizio_anno="12/09/2013"
fine_anno="08/06/2019"

anno_scolastico="1920"
 

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
'pathServer=""

Const pwdAdmin="ab95bb5fbdfcd054fe769154a06f9fadabc4c067cde0bcce1083f4249951bc16"  


'Session.Timeout=60
Const AdminPassword = "maurosho" ' per la chat
Const codAdmin="informistica"
Const eMailAdmin="prof.spinarelli.mauro@gmail.com"

dim vetCodici(),vetImg(500,4)

Session.CodePage = 65001


%>
