 <%%>
  <script type="text/javascript" src="../js/personalizza.js"></script>
		<script type="text/javascript">
	
		 
$(window).load(function () {
	   
	   $('#<%=box_apri%>').click();
	   $('#FissaTopBar').click();
	   $("body").addClass("theme-"+"<%=session("stile")%>").attr("data-theme","theme-"+"<%=session("stile")%>");

	 
	  // event.stopPropagation();
	    
	});
	
</script>
 <%%>