function creaGrafico(date, datasets, massimo) {
  let ctx = document.getElementById("grafico").getContext('2d');
  new Chart(ctx, {
    type: 'line',
    data: {
      labels: date,
      datasets: datasets
    },
    options: {
      maintainAspectRatio: false,
      scales: {
        yAxes: [{
          display: true,
          ticks: {
            beginAtZero: true,
            steps: massimo,
            stepValue: 1,
            max: massimo
          }
        }]
      },
      legend: {
        display: true,
        // position:innerWidth>1000?"right":"bottom",
        // align:"center"
      }
    }
  });
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

function prendiDateN(s, inizio, f) {
  //Divido la stringa ad ogni & e prendo solo le parti che combaciano con le date
  let divisione = s.split("&");
  let date = new Array();
  let indice = 0;
  for (let i = 0; i < divisione.length && date.length < f - inizio + 1; i++)
    if (combacia(divisione[i])) {
      if (indice >= inizio && indice <= f)
        date.push(divisione[i]);
      indice++;
    }
  return date;
}

function prendiVoti(r, inizio, f) {
  //Creo la matrice dei voti
  let voti = new Array();
  for (let i = 0; i < r.length; i++)
    voti.push(new Array());
  let indice;
  let s;
  //Aggiungo solo i voti
  for (let i = 0; i < r.length; i++) {
    indice = 0;
    s = 0;
    for (let j = 2; j < r.length; j++) {
      if ((j - 2) % 3 == 0) {
        if (indice >= inizio && indice <= f)
        voti[i].push( Number(r[i][j]));
        indice++;
      }
    }
    
  }
  return voti;
}


function prendiNomi(r) {
  //Prendo i nomi degli studenti
  let nomi = new Array();
  for (let i = 0; i < r.length; i++)
    nomi.push(i + " " + r[i][0]);
  return nomi;
}

function creaDatasets(nomi, voti, alunni) {
  let datasets = new Array();
  let rosso, verde, blu;
  if (alunni.length == 0)
    for (let i = 0; i < voti.length; i++) {
      rosso = parseInt(Math.random() * 255);
      verde = parseInt(Math.random() * 255);
      blu = parseInt(Math.random() * 255);
      datasets.push({
        label: nomi[i],
        backgroundColor: "rgb(" + rosso + "," + verde + "," + blu + ")",
        borderColor: "rgb(" + rosso + "," + verde + "," + blu + ")",
        borderWidth: 1,
        data: voti[i],
        fill: false
      });
    }
  else {
    for (let i = 0; i < alunni.length; i++) {
      rosso = parseInt(Math.random() * 255);
      verde = parseInt(Math.random() * 255);
      blu = parseInt(Math.random() * 255);
      datasets.push({
        label: nomi[alunni[i]],
        backgroundColor: "rgb(" + rosso + "," + verde + "," + blu + ")",
        borderColor: "rgb(" + rosso + "," + verde + "," + blu + ")",
        data: voti[alunni[i]],
        fill: false
      });
    }
  }
  return datasets;
}

function resetCanvas() {
  $('#grafico').remove(); // Tolgo la canvas
  $('#perIlGrafico').append('<canvas id="grafico"><canvas>'); //Creo di nuovo la canvas
};

function calcolaMassimo(v) {
  let max = 0;
  for (let i = 0; i < v.length; i++)
    for (let j = 0; j < v[i].length; j++)
      if (v[i][j] > max)
        max = Number(v[i][j]);
  return max;
}