<%
Dim Inizio_Timer
Inizio_Timer = Timer

ora = cDate(Request("h"))
fuso = clng(Request("fuso"))

'ora = cDate("00:00:01")
'fuso = 1

if fuso <> 1 then
	toadd = (fuso)
else
	toadd = 0
end if

ora = DateAdd("h",toadd,ora)
attuale = Time()
'attuale = cDate("23:59:59")
'response.write attuale

differenza = Abs(DateDiff("s",attuale,ora))
'response.write differenza

Dim Fine_Timer
Fine_Timer = Timer

Dim Tempo_Totale
Tempo_Totale = FormatNumber(Fine_Timer - Inizio_Timer, 2)
Tempo_Totale = clng(Tempo_Totale)

'Response.Write "<center>Tempo necessario al caricamento della pagina: " & Tempo_Totale & " sec.</center>"

differenza = differenza - Tempo_Totale

' disabilito il controllo 
if differenza <= 10 or differenza >= 86390 then
	'response.write "sync"
else
	'response.write "notsync"
	'response.write toadd
end if
' rispondo che è sempre sincronizzato, per il problema con il cambio del fuso orario
response.write "sync"


%> 