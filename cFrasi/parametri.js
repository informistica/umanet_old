//"/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp" == l && (frmDocument.S1.value = "");
var temp = new Array,
    i = 1,
    caratteri_iniz = 0;
"/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp" != l && (caratteri_iniz = frmDocument.txtSpiegazione.value.length);
var caratteri_attuali = 0,
    eccessotime = !1;
data = Date.now(), temp[0] = parseInt(data / 1e3);
var t = setInterval(function() {
    i % 1 == 0 && ((caratteri_attuali = "/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp" == l ? frmDocument.S1.value.length : frmDocument.txtSpiegazione.value.length) - caratteri_iniz > 13 ? eccessotime = !0 : caratteri_iniz = caratteri_attuali), data = Date.now(), temp[i] = parseInt(data / 1e3), i++
}, 1e3);

function getParametri(e) {
    clearInterval(t);
    for (var a = !1, i = 1; i < temp.length && !a; i++)(temp[i] - temp[i - 1] < -1 || temp[i] - temp[i - 1] > 4) && ((0 != temp[i] || 57 != temp[i - 3] && 58 != temp[i - 2] && 59 != temp[i - 1]) && (a = !0), 1 == temp[i] && temp[i - 1] < 57 && (a = !0));
    if (a) alert("Hai disabilitato JavaScript! Inserimento annullato."), window.location.href = window.location.href;
    else if (eccessotime && 0 == ci) alert("Rilevato uno standard input. Inserimento annullato."), window.location.href = window.location.href;
    else {
        var r = (new Date).getTimezoneOffset() / -60,
            n = new Date,
            o = new Date(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate(), n.getUTCHours(), n.getUTCMinutes(), n.getUTCSeconds()),
            s = 0;
        s = o.getHours() < 10 ? "0" + o.getHours() : o.getHours(), o.getMinutes() < 10 ? s += ":0" + o.getMinutes() : s += ":" + o.getMinutes(), o.getSeconds() < 10 ? s += ":0" + o.getSeconds() : s += ":" + o.getSeconds(), $.ajax({
            method: "POST",
            url: "verificaparametri.asp",
            dataType: "html",
            data: {
                fuso: r,
                h: s
            }
        }).done(function(t) {
            "notsync" == t ? (alert("Il tuo orologio non è sincronizzato: inserimento annullato. Sincronizza l'orario del tuo PC con quello di Internet prima di riprovare."), window.location.href = window.location.href) : "/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp" == l ? 0 == e ? document.getElementById("frmDocument").submit() : invio() : 0 == e ? document.getElementById("frmDocument").submit() : checkImg()
        })
    }

}
