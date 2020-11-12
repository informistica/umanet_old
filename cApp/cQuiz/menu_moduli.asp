<!-- richiama_test.asp -->
<%@ Language=VBScript %>

<%
response.addHeader "Access-Control-Allow-Origin", "*"
response.addHeader "Access-Control-Allow-Credentials", "true"

paragrafo = Request.QueryString("paragrafo")

%>

<% if paragrafo <> 1 then %>

<li class="">
	<a href="#" ontouchstart="quizselezionato(0, 'Expo_9')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_9" class="menu-text"> SlegalItalia </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(1, 'Expo_U_2')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="1" value="Expo_U_2" class="menu-text"> Elexpo </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(2, 'Expo_7')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="2" value="Expo_7" class="menu-text"> Cambia il Mondo </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(3, 'Expo_1')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="3" value="Expo_1" class="menu-text"> Food For World </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(4, 'Expo_2')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="4" value="Expo_2" class="menu-text"> Food For All </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(5, 'Expo_3')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="5" value="Expo_3" class="menu-text"> Food For Health </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(6, 'Expo_4')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="6" value="Expo_4" class="menu-text"> Food For Culture </span>
	</a>

	<b class="arrow"></b>
</li>

<% else %>

<li class="">
	<a href="#" ontouchstart="quizselezionato(0, 'Expo_6_1')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_6_1" class="menu-text"> Prospettive </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(1, 'Expo_6_2')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_6_2" class="menu-text"> Social Game </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(2, 'Expo_6_3')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_6_3" class="menu-text"> Whistleblowing </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(3, 'Expo_6_4')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_6_4" class="menu-text"> Protezione </span>
	</a>

	<b class="arrow"></b>
</li>

<li class="">
	<a href="#" ontouchstart="quizselezionato(4, 'Expo_6_5')">
        <i class="menu-icon fa fa-arrow-right"></i>
        <span id="0" value="Expo_6_5" class="menu-text"> Investimento </span>
	</a>

	<b class="arrow"></b>
</li>



<% end if %>