<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->

<%
'NOTE: We need to destory both "userid" cookies for ASP and ASP.NET.
'If coming from the logout.aspx page we can destory the ASP cookie.
'If coming from any other page to logout we first need to redirect to the
'rd_logout.aspx page, destory the ASP.NET userid cookie, and then return here
'to destory the ASP userid cookie.
    response.cookies("userid") = ""
    %>
    <html>
    <head>
<script src="https://apis.google.com/js/platform.js" async defer></script>
<meta name="google-signin-client_id" content="1087616019738-p41a8s5a4hd9k7b6r4j27sto7d1e760d.apps.googleusercontent.com">
<body>
<script>
 function init(){
    var auth2 = gapi.auth2.getAuthInstance();
    auth2.signOut().then(function () {
    	window.location="<%=GetEGovDefaultPage(iorgid)%>";
    });
}
</script>
<meta http-equiv="refresh" content="2;url=<%=GetEGovDefaultPage(iorgid)%>" />
<div id="g-signin" class="g-signin2" data-onsuccess="init" style="display:none;"></div>
</body>
</html>
    <%
    response.end

 if request("from") = "ASPX" then
    response.cookies("userid") = ""

'dim objCookie
'loop through cookie collection

'For Each objCookie in Request.Cookies
	'delete the cookie by setting its expiration date
	'to a date in the past
'	Response.Cookies(objCookie).Expires = "1/1/2000"
'Next

    response.redirect GetEGovDefaultPage(iorgid)
 else
    response.Redirect "rd_logout.aspx"
 end if
 
%>

<!-- #include file="include_top_functions.asp" //-->
