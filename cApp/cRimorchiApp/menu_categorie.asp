<!-- richiama_test.asp -->
<%@ Language=VBScript %>

<% Session.CodePage = 65001 %>
<%
response.addHeader "Access-Control-Allow-Origin", "*"
response.addHeader "Access-Control-Allow-Credentials", "true"

%>

<button class="btn btn-success" style="width:40%; height:60px; background-color: #27ae60!important; border-color:#27ae60!important; color:white!important;" ontouchstart="categoriaselezionata(0, 'Frasi impossibili per rimorchiare', 'FRASE IMPOSSIBILE')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" class="menu-text"> Impossibili </span>
	<b class="arrow"></b>
</button>

<button class="btn btn-primary" style="width:40%; height:60px; background-color: #428bca!important; border-color:#428bca!important; color:white!important;" ontouchstart="categoriaselezionata(1, 'Frasi pessime', 'FRASE PESSIMA')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="1" class="menu-text"> Pessime </span>
	<b class="arrow"></b>
</button>

<button class="btn btn-yellow" style="width:40%; height:60px; margin-top:5px; background-color: #f1c40f!important; border-color:#f1c40f!important; color:white!important;" ontouchstart="categoriaselezionata(2, 'Frasi scontate', 'FRASE SCONTATA')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="2" class="menu-text"> Scontate </span>
	<b class="arrow"></b>
</button>

<button class="btn btn-danger" style="width:40%; height:60px ;margin-top:5px; color:white!important" ontouchstart="categoriaselezionata(3, 'Frasi carine', 'FRASE CARINA')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="3" class="menu-text"> Carine </span>
	<b class="arrow"></b>
</button>

