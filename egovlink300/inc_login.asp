<script src="https://apis.google.com/js/platform.js" async defer></script>
<meta name="google-signin-client_id" content="1087616019738-p41a8s5a4hd9k7b6r4j27sto7d1e760d.apps.googleusercontent.com">
<%
'Process message from login attempt
If request.querystring <> "" Then 
	sStatus    = Decode(request.querystring(encode("STATUS")))
	sUserEmail = Decode(request.querystring(encode("USERLOGIN")))

	'Set message to user
	Select Case sStatus
		Case "FAILED"
			sMsg = "The logon and password entered are incorrect."
		Case Else 
			sMsg = ""
	End Select 
End If 
%>
	<script>
	<!--

		function validate()
		{
			if (document.frmLogin.frmsubjecttext.value != '')
			{
				document.frmLogin.frmsubjecttext.focus();
				alert("Please remove any input from the Internal Only field at the bottom of the form.");
				return false;
			}
			return true;
		}

	//-->
	</script>
	<%
response.write vbcrlf & "<form name=""frmLogin"" id=""frmLogin"" action=""login.asp"" method=""post"" autocomplete=""off"">"
response.write vbcrlf & "<input type=""hidden"" name=""token"" value=""" & request.querystring("token") & """ />"
response.write vbcrlf & "<table cellspacing=""2"" cellpadding=""0"" border=""0"">"
response.write vbcrlf & "<input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & iorgid & """ />"

If sMsg <> "" Then 
	response.write vbcrlf & "<tr>"
	response.write "<td colspan=""3"">"
	response.write vbcrlf & "<p><font style=""color:#ff0000; padding:5px 5px;"">" & sMsg & "</font></p><br />"
	response.write "</td>"
	response.write "</tr>"
End If 

response.write vbcrlf & "<tr>"
response.write "<td align=""right""><strong>Email:</strong>&nbsp;&nbsp;</td>"
response.write "<td align=""left""><input type=""text"" name=""email"" id=""email"" value=""" & sUserEmail & """ autocomplete=""off"" /></td>"
response.write "</tr>"

response.write vbcrlf & "<tr>"
response.write "<td align=""right""><strong>Password:</strong>&nbsp;&nbsp;</td>"
response.write "<td align=""left""><input type=""password"" name=""password"" id=""password"" value="""" autocomplete=""off"" /></td>"
response.write "</tr>"

response.write vbcrlf & "<tr>"
response.write "<td>&nbsp;</td>"
response.write "<td align=""left""><input type=""submit"" class=""actionbtn"" value=""Sign in"" onclick=""return validate()"" /></td>"
response.write "</tr>"


%>
<!--tr>
<td colspan="2">
<br />
<br />
<div id="g-signin" onclick="setClickTrue();" data-onsuccess="onSignIn"></div>
<a id="g-signout" href="#" onclick="signOut();" style="display:none;">Sign out</a>
</td>
</tr>
<script>
    function renderButton() {
      gapi.signin2.render('g-signin', {
        'scope': 'profile email',
        'width': 285,
        'height': 50,
        'longtitle': true,
        'theme': 'dark',
        'onsuccess': onSignIn,
        'onfailure': null
      });
    }
var click = false;
  function signOut() {
    var auth2 = gapi.auth2.getAuthInstance();
    auth2.signOut().then(function () {
	//document.getElementById("g-signout").style.display = 'none';
	document.getElementById("g-signin").style.display = '';
    });
  }
  function onSignIn(googleUser) {
        // The ID token you need to pass to your backend:
        var id_token = googleUser.getAuthResponse().id_token;

	//document.getElementById("g-signout").style.display = '';
	document.getElementById("g-signin").style.display = 'none';


	//IF NOT ALREAY LOGGED IN!
	if (readCookie("userid") == "" || readCookie("userid") == "-1" || click)
	{
		window.location = 'test_gauth.asp?id_token=' + id_token + '<%=querystring%>';
	}

}
function setClickTrue()
{
	click = true;
}

(function(){
    var cookies;

    function readCookie(name,c,C,i){
        if(cookies){ return cookies[name]; }

        c = document.cookie.split('; ');
        cookies = {};

        for(i=c.length-1; i>=0; i--){
           C = c[i].split('=');
           cookies[C[0]] = C[1];
        }

        return cookies[name];
    }

    window.readCookie = readCookie; // or expose it however you want
})();
</script>
  <script src="https://apis.google.com/js/platform.js?onload=renderButton" async defer></script-->
