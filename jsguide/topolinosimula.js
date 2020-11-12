 

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
		{
            target: '.box-title',
              content: ' <a target=_blank href=https://www.umanet.net/expo2015/UECDL/U-ECDL/home_img.asp>Interfaccia Umanet World Wide Web</a> per rappresentare la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html><b>Navigazione nella Rete della Vita.</b></a><br>Questa esercitazione consiste nel verificare la corretta configurazione della scena di gioco del  Soggetto che vuole raggiungere un obiettivo attraverso la metafora del <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html><b>Topolino nel Labirinto</b></a>',
            direction: 'left'
          },
		
          {
            target: '#Parametri',
            content: 'Questi sono i parametri attuali con i quali &egrave; configurata la metafora, in ordine rappresentano:<b> Soggetto, Obiettivo, Motivazione, Contesto, Comportamento,Azione Virtuosa, Azione Viziosa, Esito negativo.</b>',
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
           
            margin: {
              bottom: 500
            },
            events: {
              select: function(e) {
                $('a.view-source-link').on('click', viewsource);
              },
              deselect: function(e) {
                $('a.view-source-link').off('click', viewsource);
              }
            }
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