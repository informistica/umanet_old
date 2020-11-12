function creaGrafico(labels, dati) {
  let c = colori();
  let ctx = document.getElementById("grafico").getContext('2d');
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: "Punteggi",
        data: dati,
        backgroundColor: c,
        borderColor: c,
        borderWidth: 1
      }]
    },
    options: {
      scales: {
        yAxes: [{
          ticks: {
            beginAtZero: true
          }
        }]
      }
    }
  });
}

function colori() {
  let a = new Array();
  for (let i = 0; i < 23; i++)
    a.push("rgb(" + parseInt(Math.random() * 255) + "," + parseInt(Math.random() * 255) + "," + parseInt(Math.random() * 255) + ")");
  return a;
}

function combacia(s) {
  //La data deve essere nel formato gg_mm_aaaa
  //////////////////////////////////0123456789
  if (s.length != 10)
    return false;
  //Controllo che i caratteri del giorno siano cifre
  if (isNaN(parseInt(s.substring(0, 2))))
    return false;
  //Controllo che i caratteri del mese siano cifre
  if (isNaN(parseInt(s.substring(3, 5))))
    return false;
  //Controllo che i caratteri dell'anno siano cifre
  if (isNaN(parseInt(s.substring(6))))
    return false;

  //Se arriva a questo punto probabilmente è una data
  return true;

}

function prendiDate(s) {
  //Divido la stringa ad ogni & e prendo solo le parti che combaciano con le date
  let divisione = s.split("&");
  let date = new Array();
  for (let i = 0; i < divisione.length; i++)
    if (combacia(divisione[i]))
      date.push(divisione[i]);
  return date;
}

function prendiPunti(r, pi, pf) {
  //Creo la matrice dei voti
  let punti = new Array();
  //Aggiungo solo i punti
  let s;
  let indice = 0;
  for (let i = 0; i < r.length; i++) {
    s = 0;
    indice = 0;
    for (let j = 3; j < r.length; j++)
      if ((j - 3) % 3 == 0) {
        if (indice > pi && indice <= pf)
          s += Number(r[i][j]);
        else if (pi == pf && indice == pi)
          s += 1;
        indice++;
      }
    punti.push(s);
  }
  return punti;
}

function prendiNomi(r) {
  //Prendo i nomi degli studenti
  let nomi = new Array();
  for (let i = 0; i < r.length; i++)
    nomi.push(r[i][0]);
  return nomi;
}

function bubbleSort(punti, nomi) {
  let indice = punti.length - 1;
  let fine = false;
  while (!fine && indice >= 0) {
    fine = true;
    for (let i = 0; i < indice; i++)
      if (punti[i] < punti[i + 1]) {
        scambia(punti, i);
        scambia(nomi, i);
        fine = false;
      }
    indice--;
  }
}

function scambia(vettore, i) {
  let m = vettore[i];
  vettore[i] = vettore[i + 1];
  vettore[i + 1] = m;
}

function resetCanvas() {
  $('#grafico').remove();
  $('#contenitore').append('<canvas id="grafico"><canvas>');
};