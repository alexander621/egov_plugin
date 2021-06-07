<meta name="viewport" content="width=device-width, initial-scale=1" />
<!--#include file="includes/common.asp" //-->
<!--#include file="includes/start_modules.asp" //-->
<!--#include file="include_top_functions.asp" //-->
<script src="https://apis.google.com/js/platform.js" async defer></script>
<meta name="google-signin-client_id" content="1087616019738-p41a8s5a4hd9k7b6r4j27sto7d1e760d.apps.googleusercontent.com">
<style>
form {
    max-width: 330px;
    padding: 15px;
    margin: 0 auto;
}
#problemtextfield1
{
	display:none;
}
input[type="submit"], li
{margin-top:10px;}
input[type="text"], select, input[type="password"], input[type="button"], input[type="submit"]
{ 
	font-size:16px;
	width: 100% !important;
}
</style>
<%
Dim sError, oActionOrg
Dim iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP

assistant = "alexa"
if request.querystring("googleauth") = "true" then assistant = "googleauth"

querystring = "&" & request.servervariables("QUERY_STRING")

if request.cookies("userid") <> "" and request.cookies("userid") <> "-1" AND request.querystring(assistant) <> "true" then
	'Clear Cookie if logged in but token has expired
	sSQL = "SELECT authtokenid FROM authtokens WHERE userid = '" & track_dbsafe(request.cookies("userid")) & "' AND token = '" & track_dbsafe(request.querystring("token")) & "' and orgid = '" & iorgid & "' and daterecorded >= '" & DateAdd("n",-5,Now()) & "'"
	Set oAuth = Server.CreateObject("ADODB.RecordSet")
	oAuth.Open sSQL, Application("DSN"), 3, 1
	if oAuth.EOF then
		response.cookies("userid") = ""
	end if
	oAuth.Close
	Set oAuth = Nothing
elseif request.cookies("userid") <> "" and request.cookies("userid") <> "-1" AND request.querystring(assistant) = "true" then
		response.cookies("userid") = ""

end if


if request.cookies("userid") = "" or request.cookies("userid") = "-1" then%>
<form>
	Login to link your account for <%=sOrgName%>:
</form>
<!--#include file="inc_login.asp"-->
<tr>
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
	document.getElementById("g-signout").style.display = 'none';
	document.getElementById("g-signin").style.display = '';
    });
  }
  function onSignIn(googleUser) {
        // The ID token you need to pass to your backend:
        var id_token = googleUser.getAuthResponse().id_token;

	document.getElementById("g-signout").style.display = '';
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
  <script src="https://apis.google.com/js/platform.js?onload=renderButton" async defer></script>
<%
if request.querystring("state") <> "" then
	response.write "<input type=""hidden"" name=""" & assistant & """ value=""true"" />"
	response.write "<input type=""hidden"" name=""state"" value=""" & request.querystring("state") & """ />"
	response.write "<input type=""hidden"" name=""redirect_uri"" value=""" & request.querystring("redirect_uri") & """ />"
end if
response.write vbcrlf & "<tr>"
response.write "<td colspan=""2"" align=""left"">"
response.write "<p><br />"
response.write vbcrlf & "<ul>"
response.write vbcrlf & "<li><a href=""basic_forgot_password.asp"">Can't remember your password?</a></li>"
response.write vbcrlf & "<li><a href=""basic_register.asp?" & request.servervariables("QUERY_STRING") & """>Not registered yet?</a></li>"
response.write vbcrlf & "</ul>"
response.write "</p>"
response.write "</td>"
response.write "</tr>"
response.write vbcrlf & "<tr>"
response.write "<td colspan=""2"" align=""left"">" 
response.write vbcrlf & "<div id=""problemtextfield1"">"
response.write vbcrlf & "Internal Use Only, Leave Blank: <input type=""text"" name=""frmsubjecttext"" id=""problemtextinput"" value="""" size=""6"" maxlength=""6"" /><br />"
response.write vbcrlf & "<strong>Please leave this field blank and remove any <br />values that have been populated for it.</strong>"
response.write "</div>"
response.write "</td>"
response.write "</tr>"


response.write vbcrlf & "</table>"
response.write vbcrlf & "</form>"
elseif request.querystring("alexa") = "true" then
	response.redirect "https://pitangui.amazon.com/api/skill/link/M3B45EZT5APEYP"
elseif request.querystring("googleauth") = "true" then
	response.redirect "https://oauth-redirect.googleusercontent.com/r/e-gov-link-b19d5"
else%>
You're logged in!
<% end if %>

