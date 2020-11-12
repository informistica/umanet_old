 

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
		 {
            target: '.box-title',
              content: ' <a target=_blank href=https://www.umanet.net/expo2015/UECDL/U-ECDL/home_img.asp>Interfaccia Umanet World Wide Web</a> per rappresentare la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Metafore/Pagine/Umanet_World_Wide_Web.html><b>Navigazione nella Rete della Vita.</b></a><br>Questa esercitazione consiste nel configurare la scena di gioco per un Soggetto che vuole raggiungere un obiettivo attraverso la metafora della <a target=_blank href=https://www.umanet.net/informistica/UWWW/Umanet/Pagine/Umanet_Evolution.html><b>Navigazione verso un Obiettivo</b></a>',
            direction: 'left'
          },
          {
            target: '#Parametri',
             content:' Questi sono i parametri attuali con i quali &egrave; configurata la metafora, in ordine :<br><b> Soggetto, Obiettivo, Motivazione, Contesto, Comportamento,Azione Virtuosa, Azione Viziosa,Segnali di pericolo, Esito negativo, ed infine Comportamenti sbagliati da lasciare andare.</b>',
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
            direction: 'left'
          },
		   {
            target: '#idImg',
            content: 'La rappresentazione metaforica dei vari momenti che scandiscono la qualit&agrave; della navigazione</b>',
             direction: 'left',
            direction: 'left'
          },
		  
		  
          {
            target: '#Boxtext',
              content: 'La narrazione prodotta dallo script combinando i parametri statici con le risposte che fornirai alle domande.',
            direction: 'left',
           
          },
		  
           
           
		  
        ]
      }

      // Execute on DOM ready
      $(function() {
        $.pageguide(defaultGuide);      
      });