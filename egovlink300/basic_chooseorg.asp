<meta name="viewport" content="width=device-width, initial-scale=1" />
<%
    response.cookies("userid") = ""
    %>
<!--#include file="includes/common.asp" //-->
<!--#include file="includes/start_modules.asp" //-->
<!--#include file="include_top_functions.asp" //-->
<%
	querystring = ""
	if request.servervariables("query_string") <> "" then querystring = "?" & request.servervariables("query_string")

assistant = "alexa"
if request.querystring("googleauth") = "true" then assistant = "google"
%>
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
<script src="https://apis.google.com/js/platform.js" async defer></script>
<meta name="google-signin-client_id" content="1087616019738-p41a8s5a4hd9k7b6r4j27sto7d1e760d.apps.googleusercontent.com">
<script>
 function init(){
    var auth2 = gapi.auth2.getAuthInstance();
    auth2.signOut().then(function () {
    	//window.location="<%=GetEGovDefaultPage(iorgid)%>";
    });
}
</script>
<script>
	function GoToSelection(e)
	{
		window.location = e + "/basic_login.asp<%=querystring%>"
	}
</script>
<%
Dim sError, oActionOrg
Dim iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP



sSQL = "SELECT DISTINCT o.orgid,OrgName,OrgEgovWebsiteURL " _
	& " FROM Organizations o " _
	& " INNER JOIN alexaconfigs ac ON ac.orgid = o.orgid"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
%>
<form>
Select your municipality below (if your's isn't listed, ask them to add E-Gov for <%= UCase(Left(assistant,1)) & LCase(Right(assistant, Len(assistant) - 1))%>):
<select name="orgs" onChange="GoToSelection(this.value)">
	<option value="">Choose...</option>
<%
Do While not oRs.EOF
	%><option value="<%=oRs("OrgEgovWebsiteURL")%>"><%=oRs("OrgName")%></option><%
	oRs.MoveNext
loop
%>
</select>
</form>
<div id="g-signin" class="g-signin2" data-onsuccess="init" style="display:none;"></div>
<%
oRs.Close
Set oRs = Nothing


%>
