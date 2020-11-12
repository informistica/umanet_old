 

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
		{
            target: '.box-title',
              content: ' <a target=_blank href=https://www.umanet.net/expo2015/UECDL/U-ECDL/home_img.asp>Interfaccia Umanet World Wide Web</a> per rappresentare la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html><b>Navigazione nella Rete della Vita.</b></a><br>Questa esercitazione consiste nel verificare la corretta configurazione della scena di gioco del  Soggetto che vuole raggiungere un obiettivo attraverso la metafora del <a target=_blank href=https://www.umanet.net/informistica/UWWW/Umanet/Pagine/Connessioni_UWWW.html><b>Client/Server</b></a>',
            direction: 'left'
          },
		 {
            target: '#btnAggiorna',
           content: 'Bottone per aggiornare le modifiche ai parametri della metafora per meglio integrarli nella narrazione prodotta dallo script',
            direction: 'left',
           // shadow: false
          },
          {
            target: '#idClient',
            content: 'Questi sono i parametri attuali con i quali &egrave; configurata la metafora dal lato Client: Il Soggetto che effettua la  Domanda, la sua  Aspettativa, il Motivo di questa aspettativa, cosa Desidera ottenere e per soddisfare quale Bisogno?',
            direction: 'left'
          },
		    {
            target: '#idServer',
              content: 'Questi sono i parametri attuali con i quali &egrave; configurata la metafora dal lato SERVER: Il soggetto che risponde alla richiesta,la sua Risposta che pu&ograve; soddisfare la richiesta oppure deluderla, comunicazione <b>Coerente o Paradossale</b>? ... le altre domande sono speculari al Client.',
            direction: 'left'
          },
		  {
            target: '#tipoEvento',
              content: '<b>Indica se &egrave; stata configurata una comunicazione <b>Coerente</b>, in cui Aspettativa del Client e Realt&agrave; del Server si incontrano, oppure <b>Paradossale</b> in cui si producono significati in contrasto e divergenti accumulando tensione, come in un elastico teso tra due estremi.',
            direction: 'left'
          },
		  {
            target: '#idInizio',
             content: 'Il bottone per accendere il modello e mettere in moto la simulazione della realt&agrave;</b>',
            direction: 'left'
          },
          {
            target: '#btnSxDx',
            content: 'I bottoni per rispondere alle domande poste dallo script...</b>',
             direction: 'left',
          },
		  
          
            {
            target: '#Boxtext',
           content: 'La narrazione prodotta dallo script combinando i parametri statici con le risposte che fornirai alle domande.',
            direction: 'left',
           // shadow: false
          },
          
		  
		  
        ]
      }

      // Execute on DOM ready
      $(function() {
        $.pageguide(defaultGuide);      
      });