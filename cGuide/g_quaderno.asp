         <!-- Placed at the end of the document so the page loads faster -->
   <!-- <script src="../../../guida/docs/lib/jquery.js"></script>-->
    <script src="../../../guida/docs/lib/bootstrap/js/bootstrap-dropdown.js"></script>
    <script src="../../../guida/docs/lib/google-code-prettify/prettify.js"></script>

    <script src="../../../guida/js/jquery.pageguide.js"></script>
    <script language="javascript">
      /**
       * Helper Functions
       */

      // View source of current page in a new window
      function viewsource(e){
        window.open("view-source:" + window.location, 'jquery.pageguide.source');
      }

      // Smooth scroll to anchor
      function scrollTo(e) {
        e.preventDefault();

        var anchor = e.currentTarget.hash.slice(1);
            $t = $('a[name=' + anchor + ']');

        if (!$t.size()) return;

        var dvh = $(window).height(),
            dvtop = $(window).scrollTop(),
            eltop = $t.offset().top,
            mgn = {top: 100, bottom: 100};

        var scrollTo = eltop - mgn.top;

        $('html,body').animate({
          scrollTop: scrollTo
        }, {
          duration: 400
        });
      }

      // Example guides
     

      

      // Load an example guide
      

      // Guide definition for the guide loaded on dom ready
      var defaultGuide = {
        id: 'jQuery.PageGuide',
        title: 'Fai un giro veloce ',
        steps: [
          
		   {
            target: '#diario',
            content: 'Compiti assegnati e altre comunicazioni di servizio. ',
            direction: 'right'
          },
		   {
            target: '#lavagna',
            content: 'Bacheca per organizzare attivit&agrave;risorse extra, consegne e altro.',
            direction: 'right'
          },
		  {
            target: '#libro',
            content: 'Risorse per svolgere i compiti, libri di testo, pubblicazioni, siti, ecc..',
            direction: 'right'
          },
          {
            target: '#compiti',
            content: 'Le ultime comunicazioni dal diario, bacheca, forum pi&ugrave; i compiti svolti',
            direction: 'right'
          },
          
          {
            target: '#mappe',
            content: 'Mappe concettuali create per riassumere argomenti complessi  ',
            direction: 'right'
          },
        {
            target: '#interrogazioni',
            content: 'Punteggi ottenuti durante le sessioni orali del flusso di lavoro ',
            direction: 'right'
          },
      
		  {
            target: '#classifica',
            content: 'Riporta tutti i punteggi delle attivit&agrave; della classe in forma di graduatoria ',
            direction: 'right'
          },
		  {
            target: '#calendario',
            content: 'Eventi e notifiche dal calendario per non dimenticare le scadenze',
            direction: 'right'
          },
		   {
            target: '#librou',
            content: 'Le risorse su cui svolgere attivit&agrave; con le metafore interattive',
            direction: 'right'
          },
		   {
            target: '#quadernou',
            content: 'Le attivit&agrave; svolte nel Libro Umanet sulle metafore interattive, con accesso in lettura e modifica ',
            direction: 'right'
          },
		  
          {
            target: '#forum',
            content: 'Spazio social per discussioni su argomenti di istruzione e formazione',
            direction: 'right'          
          },
		   {
            target: '#chat',
            content: 'Spazio per registrare conversazioni multimediali su argomenti di istruzione e formazione',
            direction: 'right'          
          },
		   {
            target: '#periodo',
            content: 'Seleziona le attivit&agrave; da mostrare, e decidi se mostrare i Punti Social',
            direction: 'right'          
          },
		  
		   {
            target: '#bacheca',
            content: 'Gli ultimi tre post pi&ugrave; recenti dal social network',
            direction: 'right'          
          },
		  
		  {
            target: '#adiario',
            content: 'Spazio per appunti personali e gruppi di discussione privata',
            direction: 'top'          
          },
		   {
            target: '#apost',
            content: 'Il registro di tutte le attivit&agrave; che hai svolto nella parte social (Diario,Lavagna,Forum,Chat)',
            direction: 'top'          
          },
		   {
            target: '#report',
            content: 'Il riepilogo dei punteggi ottenuti e il tuo andamento nel tempo;',
            direction: 'right'          
          },
           {
            target: '#acompiti',
            content: 'I compiti svolti, raggruppati per capitoli, paragrafi e sottoparagrafi',
            direction: 'left'          
          },
           
		  
		  
        ]
      }

      // Execute on DOM ready
      $(function() {
        $.pageguide(defaultGuide);      
      });
    </script>
      