function creaGrafico(date, nome, voti) {
  let ctx = document.getElementById("grafico").getContext('2d');
  new Chart(ctx, {
    type: 'line',
    data: {
      labels: date,
      datasets: [{
        label: nome,
        backgroundColor: "rgb(" + 255 + "," + 0 + "," + 0 + ")",
        borderColor: "rgb(" + 255 + "," + 0 + "," + 0 + ")",
        data: voti,
        fill: false
      }]
    },
    options: {
      scales: {
        yAxes: [{
          display: true,
          ticks: {
            beginAtZero: true,
            steps: 10,
            stepValue: 1,
            max: 10
          }
        }]
      },
      legend: {
        display: true,
        position: "right",
        align: "center"
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

function prendiVoti(r) {
  //Creo la matrice dei voti
  let voti = new Array();
  for (let i = 0; i < r.length; i++)
    voti.push(new Array());
  //Aggiungo solo i voti
  for (let i = 0; i < r.length; i++)
    for (let j = 2; j < r.length; j++)
      if ((j - 2) % 3 == 0)
        voti[i].push(r[i][j]);
  return voti;
}

function prendiNomi(r) {
  //Prendo i nomi degli studenti
  let nomi = new Array();
  for (let i = 0; i < r.length; i++)
    nomi.push(r[i][0]);
  return nomi;
}

function resetCanvas() {
  $('#grafico').remove(); // Tolgo la canvas
  $('#contenitore').append('<canvas id="grafico"><canvas>'); //Creo di nuovo la canvas
};