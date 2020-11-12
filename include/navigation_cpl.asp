 
<div class="container-fluid">


				<% Select Case id_app
					  Case 1
						response.write("<a href='#' id='brand'>Umanet per il CPL   </a>")
					  Case 2
						response.write("<a href='#' id='brand'>Umanet per la CNV   </a>")
					
					End Select %>

			
			<a href="#" class="toggle-nav" rel="tooltip" data-placement="bottom" title="Toggle navigation"><i class="icon-reorder"></i></a>
			<ul class='main-nav'>     
                <li><a href="#"> <i class="icon-home"></i>
						<span>Home </span>
					</a>
				</li>
				
                
                <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Gestione</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">
                      <li><a   href="#"><span></span><i class="glyphicon-log_book"></i>Voce 1</a></li>
                    


								
					</ul>
				</li>
               

			</ul>
            
              
            
        </div>     
     
			
        
        <script type="text/javascript">
		
		$('.red').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=red"
			 event.stopPropagation();
		});
			$('.green').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=green"
			 event.stopPropagation();
		});
			$('.brown').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=brown"
			 event.stopPropagation();
		});
			$('.lime').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=lime"
			 event.stopPropagation();
		});
			$('.purple').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=purple"
			 event.stopPropagation();
		});
		$('.pink').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=pink"
			 event.stopPropagation();
		});
		
		$('.magenta').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=magenta"
			 event.stopPropagation();
		});
		
			$('.grey').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=grey"
			 event.stopPropagation();
		});
			$('.darkblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=darkblue"
			 event.stopPropagation();
		});
			$('.lightgrey').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=lightgrey"
			 event.stopPropagation();
		});
		
		 
			$('.satblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.orange').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=orange"
			 event.stopPropagation();
		});
			$('.blue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=blue"
			 event.stopPropagation();
		});
		$('.satblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.satgreen').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satgreen"
			 event.stopPropagation();
		});
			 $('.teal').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=teal"
			 event.stopPropagation();
		});
		
		</script>