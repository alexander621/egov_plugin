<script type="text/javascript" src="scripts/modules.js"></script>
<script type="text/javascript" src="scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
<!--

$(document).ready(function() {
   $('#buttonLogin').click(function() {
      location.href = 'login.asp';
   });
});

function openWin2(url, name) 
{
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function CheckForm()
{
//	alert(document.all.lookupemail.email.value);
	if (document.lookupemail.email.value == "")
	{
		alert("Please input an email address to search for.");
		document.lookupemail.email.focus();
		return false;
	}
	else
	{
		return true;
	}
}

//-->
</script>

<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf

  if request.servervariables("request_method") = "POST" then
     sSubmittedEmail = request("email")

    	setupSendMail iOrgID, _
                   sSubmittedEmail
  else
    	displayEmailLookup()
  end if
%>
<%
'------------------------------------------------------------------------------
sub setupSendMail(iOrgID, _
                  iEmail)

  dim sSql, oLogin, sOrgID, sEmail

  sOrgID = 0
  sEmail = "''"

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iEmail <> "" then
     sEmail = Track_DBsafe(iEmail)
     sEmail = "'%" & sEmail & "%'"
  end if

  sSQL = "SELECT userid, "
  sSQL = sSQL & " useremail, "
  sSQL = sSQL & " userpassword "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE useremail LIKE (" & sEmail & ") "
  sSQL = sSQL & " AND orgid = " & sOrgID
  sSQL = sSQL & " AND isdeleted = 0 "

 	set oLogin = Server.CreateObject("ADODB.Recordset")
	 oLogin.Open sSQL, Application("DSN"), 3, 1

  if not oLogin.eof then
   		sUserEmail      = oLogin("useremail")
   		sUserPassword   = oLogin("userpassword")
     sCC             = ""
     sFromEmail      = "noreply@eclink.com"
     sSubject        = "Password Assistance: " & sOrgName & " - E-Gov Services"
     sHTMLBody       = "Your " & sOrgName & " - E-Gov Services Password: (" & sUserPassword & ")"
     sTextBody       = "Your " & sOrgName & " - E-Gov Services Password: (" & sUserPassword & ")"
     sHighImportance = "Y"

     if sDefaultEmail <> "" then
        sFromEmail = sDefaultEmail  'The city's default email
     end if

     sendEmail sFromEmail, _
               sUserEmail, _
               sCC, _
               sSubject, _
               sHTMLBody, _
               sTextBody, _
               sHighImportance

     response.write "<fieldset class=""fieldset_doesnotexist"">" & vbcrlf
     response.write "  Your password has been sent to the email entered." & vbcrlf
     response.write "  <p>" & vbcrlf
     response.write "    <input type=""button"" name=""buttonLogin"" id=""buttonLogin"" value=""Return to Login"" />" & vbcrlf
     response.write "  </p>" & vbcrlf
     response.write "</fieldset>" & vbcrlf
  else
     response.write "<fieldset class=""fieldset_doesnotexist"">" & vbcrlf
     response.write "  The email address entered does not exist in the system." & vbcrlf
     response.write "  <p>" & vbcrlf
     response.write "    <input type=""button"" name=""buttonLogin"" id=""buttonLogin"" value=""Return to Login"" />" & vbcrlf
     response.write "  </p>" & vbcrlf
     response.write "</fieldset>" & vbcrlf
  end if

  oLogin.close
  set oLogin = nothing

end sub

'------------------------------------------------------------------------------
sub displayEmailLookup()
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Password Assistance</legend>" & vbcrlf
  response.write "  <form name=""lookupemail"" id=""lookupemail"" method=""post"" onsubmit=""return CheckForm();"">" & vbcrlf
  response.write "  <div id=""passwordText"">Please enter the email address that you used to register your account.</div>" & vbcrlf
  response.write "  <div>" & vbcrlf
  response.write "    <strong>Email:</strong>" & vbcrlf
  response.write "    <input type=""text"" name=""email"" id=""email"" value="""" />" & vbcrlf
  response.write "    <input type=""submit"" name=""buttonLookup"" id=""buttonLookup"" value=""Lookup"" />" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  </form>" & vbcrlf
end sub
%>
