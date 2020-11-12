<%@ Language=VBScript %>

<%Response.charset="utf-8"%>

<% 
Call Response.AddHeader("Access-Control-Allow-Origin", "*") 
paragrafo = Request.QueryString("paragrafo")
%>

<%

id = request.QueryString("id")
tipo = request.QueryString("tipo")

Dim t(14)

if paragrafo <> 1 then

t(0) = "50,Expo_9,SLEGALITALIA,Vero/Falso,Questionari sulla cultura della legalità per liberare l'Italia da mafie e malaffare"
t(1) = "49,Expo_9,SLEGALITALIA,Singola,Questionari sulla cultura della legalità per liberare l’Italia da mafie e malaffare"
t(2) = "61,Expo_U_2,SCOPRI SE SEI UN ELEXPO,Vero/Falso,Sulla consapevolezza che deve avere un nuovo essere più responsabile e sostenibile"
t(3) = "62,Expo_U_2,SCOPRI SE SEI UN ELEXPO,Singola,Sulla consapevolezza che deve avere un nuovo essere più responsabile e sostenibile"
t(4) = "59,Expo_7,TU PUOI CAMBIARE IL MONDO,Vero/Falso,Ogni persona può agire. Se ognuno di noi fa la sua parte insieme potremo ottenere ciò che è necessario"
t(5) = "60,Expo_7,TU PUOI CAMBIARE IL MONDO,Singola,Ogni persona può agire. Se ognuno di noi fa la sua parte insieme potremo ottenere ciò che è necessario"
t(6) = "51,Expo_1,FOOD FOR SUSTAINABLE GROWTH,Vero/Falso,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e crescita sostenibile"
t(7) = "53,Expo_1,FOOD FOR SUSTAINABLE GROWTH,Singola,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e crescita sostenibile"
t(8) = "55,Expo_2,FOOD FOR ALL,Vero/Falso,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e giustizia sociale"
t(9) = "52,Expo_2,FOOD FOR ALL,Singola,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e giustizia sociale"
t(10) = "54,Expo_3,FOOD FOR HEALTH,Vero/Falso,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e salute"
t(11) = "56,Expo_3,FOOD FOR HEALTH,Singola,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e salute"
t(12) = "57,Expo_4,FOOD FOR CULTURE,Vero/Falso,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e cultura"
t(13) = "58,Expo_4,FOOD FOR CULTURE,Singola,Questionari sulle pubblicazioni del BCFN sul rapporto tra cibo e cultura"

else

t(0) = "1063,Expo_6_1,LE PROSPETTIVE DEL WHISTLEBLOWER,Vero/Falso, X è  il nostro particolare whistleblower che interviene ogni volta che percepisce un'irregolarità"
t(1) = "1064,Expo_6_1,LE PROSPETTIVE DEL WHISTLEBLOWER,Singola,X è  il nostro particolare whistleblower che interviene ogni volta che percepisce un'irregolarità"
t(2) = "1065,Expo_6_2,PIATTAFORMA SOCIAL GAME ELEXPO,Vero/Falso,Il Social Game degli X ricompensa e promuove la cittadinanza attiva e responsabile"
t(3) = "1066,Expo_6_2,PIATTAFORMA SOCIAL GAME ELEXPO,Singola,Il Social Game degli X ricompensa e promuove la cittadinanza attiva e responsabile"
t(4) = "1067,Expo_6_3,PIATTAFORMA WHISTLEBLOWING ANAC,Vero/Falso,Il Whistleblowing in Italia: La piattaforma tecnologica dell'ANAC"
t(5) = "1068,Expo_6_3,PIATTAFORMA WHISTLEBLOWING ANAC,Singola,Il Whistleblowing in Italia: La piattaforma tecnologica dell'ANAC"
t(6) = "1069,Expo_6_4,PROTEZIONE DEI WHISTLEBLOWER NEI PAESI OCSE,Vero/Falso,Protezione del whistleblower nei Paesi OCSE: un confronto tra le legislazioni"
t(7) = "1070,Expo_6_4,PROTEZIONE DEI WHISTLEBLOWER NEI PAESI OCSE,Singola,Protezione del whistleblower nei Paesi OCSE: un confronto tra le legislazioni"
t(8) = "1071,Expo_6_5,ITALIA INVESTE NEL WHISTLEBLOWING,Vero/Falso,Italia investe nel WhistleBlowing: importante strumento di prevenzione della corruzione"
t(9) = "1072,Expo_6_5,ITALIA INVESTE NEL WHISTLEBLOWING,Singola,Italia investe nel WhistleBlowing: importante strumento di prevenzione della corruzione"

end if


Response.Write(cercaquiz(id))

function cercaquiz(id)

if tipo=0 then
cercaquiz = t(id*2)
else
cercaquiz = t(id*2+1)
end if

end function



%>