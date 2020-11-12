 

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
		{
            target: '.box-title',
              content: ' <a target=_blank href=https://www.umanet.net/expo2015/UECDL/U-ECDL/home_img.asp>Interfaccia Umanet World Wide Web</a> per rappresentare la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html><b>Navigazione nella Rete della Vita.</b></a><br>Questa esercitazione consiste nel configurare la scena di gioco per un Soggetto che vuole raggiungere un obiettivo attraverso la metafora del <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html><b>Topolino nel Labirinto</b></a>',
            direction: 'left'
          },
		
          {
            target: '#Parametri',
            content: 'Configura la metafora inserendo in ordine :<b> Soggetto, Obiettivo, Motivazione, Contesto, Comportamento,Azione Virtuosa, Azione Viziosa, Esito negativo.</b>',
            direction: 'left'
          },
		  {
            target: '#idDistanza',
             content: 'Un numero da 1 a 5 per indicare il grado di difficolt&agrave; e una misura di quanto bisogna perseverare nella scelta della strada giusta prima di raggiungere questo Obiettivo.</b>',
            direction: 'left'
          },
          {
            target: '#Boxtext',
            content: 'Qui dentro <b>scrivi in forma discorsiva una spiegazione</b> in cui colleghi i vari significati della metafora, come <b>un filo del discorso che unisce varie perle per farne una collana di significati...</b>',
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
            target: '.accordion-inner',
           content: 'Puoi mettere in moto il modello e fare una <b>Simulazione</b>, oppure applicare un algoritmo per generare una <b>Narrazione multimediale</b>, infine puoi zoomare e <b>Sviluppare in profondit&agrave;</b> la riflessione collegando una nuova metafora.',
            direction: 'left',
           // shadow: false
          },
          {
            target: '#btnSxDx',
           content: '<br>I bottoni servono per <b>Navigare</b> avanti ed indietro attraverso le diverse metafore collegate, sono come i tasti per muovere un ascensore che si sposta su e gi&ugrave; nello spettro dei significati.',
            direction: 'left',
           // shadow: false
          }
		  
		  
        ]
      }

      // Execute on DOM ready
      $(function() {
        $.pageguide(defaultGuide);      
      });