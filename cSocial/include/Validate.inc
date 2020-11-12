<script LANGUAGE="JAVASCRIPT">
<!--
//
function Validate()
{
sAuthor = document.InputForm.Name.value
//sEmail = document.InputForm.Email.value
//iEmailLen = sEmail.length
iAuthorLen = sAuthor.length
sTopic = document.InputForm.Topic.value
iTopicLen = sTopic.length

if(iAuthorLen == 0) {

	alert("Non hai inserito il tuo Nome")
document.InputForm.Name.focus()
return false

}
/*if(iEmailLen == 0) {
alert("Email Address is a required field")
document.InputForm.Email.focus()
return false
}*/
//check to see that at least one char in author is non-blank
var bOK = false

for(var i = 0;i<iAuthorLen;i++)
{

if(sAuthor.charAt(i) != " ") 
	{
	bOK = true
	break
	}
} //end for

if(bOK == false) {
	alert("Non hai inserito il tuo Nome")
	document.InputForm.Name.focus()
	return false
}

//check to see that at least one char in email is non-blank, @, and .
bOK = true
var bAmp = true
var bPeriod = true

/*for(var i = 0;i<iEmailLen;i++)
{
if(bOK==false){
if(sEmail.charAt(i) != " " && sEmail.charAt(i) != "." && sEmail.charAt(i) != "@") 
	{
	bOK = true
	if(bAmp== true && bPeriod== true) break
	}
}//end if bok==false

if(sEmail.charAt(i)=="@")
	{
	bAmp = true
	if(bPeriod==true && bOK==true) break
	}

if(sEmail.charAt(i)==".")
	{
	bPeriod = true
	if(bAmp==true && bOK==true) break
	}
} //end for

*/

/*if(bOK == false || bAmp == false || bPeriod == false) {
	alert("The e-mail address you entered is not valid")
	document.InputForm.Email.focus()
	return false
}//end if
*/


if(iTopicLen == 0) {

alert("Non hai inserito l'Argomento")
document.InputForm.Topic.focus()
return false

} //end if

for(var i = 0;i<iTopicLen;i++)
{

if(sTopic.charAt(i) != " ") 
	{
	return true
	
	}
} //end for


	alert("Non hai inserito l'Argomento")
	document.InputForm.Topic.focus()
	return false

}//end function
-->

</script>
