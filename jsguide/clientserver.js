 

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
          {
            target: '.box-title',
              content: ' <a target=_blank href=https://www.umanet.net/expo2015/UECDL/U-ECDL/home_img.asp>Interfaccia Umanet World Wide Web</a> per rappresentare la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html><b>Navigazione nella Rete della Vita.</b></a><br>Questa esercitazione consiste nel configurare la <a target=_blank href=https://www.umanet.net/informistica/UWWW/Umanet/Pagine/Connessioni_UWWW.html><b>Connessione</b></a> tra un <b>Client</b> che chiede ed un <b>Server</b> che risponde. Se la risposta &egrave; deludente potrebbe prodursi un <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html> Terremoto nella Relazione Umana</a>',
            direction: 'left'
          },
		   {
            target: '#idClient',
              content: 'Questi sono i parametri attuali con i quali &egrave; configurata la metafora dal lato Client in ordine : Il Soggetto che effettua la  Domanda, la sua  Aspettativa, il Motivo di questa aspettativa, cosa Desidera ottenere e per soddisfare quale Bisogno?',
            direction: 'left'
          },
		  
		   {
            target: '#idTolleranza',
              content: '<b>Un numero da 1 a 5 per configurarare la <b>Soglia Critica</b> di sopportazione del Client, oltre la quale la tensione accumulata per la delusione della aspettativa si scarica in un <b>Terremoto nella Relazione</b>.',
            direction: 'left'
          },
		   {
            target: '#idServer',
              content: 'Configurazione del <b>SERVER</b> : Il soggetto che risponde alla richiesta,la sua Risposta che pu&ograve; soddisfare la richiesta oppure deluderla, comunicazione <b>Coerente o Paradossale</b>? ... le altre domande sono speculari al Client.',
            direction: 'left'
          },
		   {
            target: '#tipoEvento',
              content: '<b>Stabilisci se hai configurato una comunicazione <b>Coerente</b>, in cui Aspettativa del Client e Realt&agrave; del Server si incontrano, oppure <b>Paradossale</b> in cui si producono significati in contrasto e divergenti accumulando tensione, come in un elastico teso tra due estremi.',
            direction: 'left'
          },
          
           {
            target: '#Boxtext',
           content: 'La narrazione prodotta dallo script combinando i parametri statici con le risposte che fornirai alle domande.',
            direction: 'left',
           // shadow: false
          },
		  
          
        
		   {
             target: '#btnSxDx',
            content: 'I bottoni per rispondere alle domande poste dallo script...</b>',
          
            direction: 'left',
           // shadow: false
          }
		  
		  
		  
        ]
      }

      // Execute on DOM ready
      $(function() {
        $.pageguide(defaultGuide);      
      });