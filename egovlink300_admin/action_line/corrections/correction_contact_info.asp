<!DOCTYPE HTML>
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/2/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	02/02/07	John Stullenberger - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Initialize and declare variables
 Dim sError
 sLevel = "../../"  'Override of value from common.asp

'Set timezone information into session
 session("iUserOffset") = request.cookies("tz")
%>
<html>
<head>
  <title>E-Gov Administration { Edit Request Contact Information }</title>
  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script type="text/javascript" src="../../scripts/modules.js"></script>
  <script type="text/javascript" src="../../scripts/jquery-1.7.2.min.js"></script>

<script type="text/javascript">
  <!--
	//Set timezone in cookie to retrieve later
	var d=new Date();
	if (d.getTimezoneOffset) {
		var iMinutes = d.getTimezoneOffset();
		document.cookie = "tz=" + iMinutes;
	}
  //-->
  </script>

<style>
  div.correctionsbox {
     border:  solid 1px #336699;
     padding: 4px 0px 0px 4px;
  }

  div.correctionsboxnotfound {
     background-color: #e0e0e0;
     border:           1px solid #000000;
     padding:          10px;
     color:            #ff0000;
     font-weight:      bold;
  }

  .correctionslabel {
     font-weight: bold;
     white-space: nowrap;
  }

  th.corrections {
     background-color: #93bee1;
     font-size:        12px;
     padding:          5px;
     color:#000000;
  }

  input.correctionstextbox {
     border: 1px solid #336699;
     width:  400px;
  }

  .savemsg {
     font-size:   12px;
     padding:     5px;
     color:       #0000ff;
     font-weight: bold;
  }

  .instructions {
     color: #ff0000;
  }
</style>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<h3>Edit Request Contact Information</h3>		" & vbcrlf
  response.write "<p><input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to Request"" class=""button"" onclick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';"" /></p>" & vbcrlf
  'response.write "<img align=absmiddle src=""../../../admin/images/arrow_2back.gif""> <a href=""../action_respond.asp?control=" & request("irequestid") & """>Return to Request</a> " & vbcrlf

	'Display to user that values were saved
		if request("r") = "save" then
  			response.write "<p><span class=""savemsg"">Saved " & Now() & ".</span></p>" & vbcrlf
		end if

 'Get contact information
		displayUserInfo(request("irequestid"))

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayUserInfo(iID)

	'Check for empty or missing userid
 	if IsNull(iID) OR iID = "" then
   		response.write "<p><div class=""correctionsboxnotfound"">No information available for this request.</div></p>" & vbcrlf
  else
			 'Get information for specified user
  			sSQL = "SELECT * "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL &      " INNER JOIN egov_users ON egov_actionline_requests.userid = egov_users.userid "
     sSQL = sSQL & " WHERE egov_actionline_requests.action_autoid = " & iID

   		set oUser = Server.CreateObject("ADODB.Recordset")
  			oUser.Open sSQL, Application("DSN"), 3, 1

  			if not oUser.eof then
    				sUserEmail = trim(oUser("useremail"))

    				response.write "<form name=""contactinfo"" id=""contactinfo"" action=""correction_contact_info_cgi.asp"" method=""post"">" & vbcrlf
				    response.write "  <input type=""hidden"" name=""status"" id=""status"" value=""" & request("status") & """ />" & vbcrlf
    				response.write "  <input type=""hidden"" name=""substatus"" id=""substatus"" value=""" & request("substatus") & """ />" & vbcrlf
				    response.write "  <input type=""hidden"" name=""irequestid"" id=""irequestid"" value=""" & iID & """ />" & vbcrlf

        displayButtons request("irequestid")

        response.write "<fieldset class=""fieldset"">" & vbcrlf
								response.write "<table>" & vbcrlf
     			response.write "  <tr><td colspan=""2""><p class=""instructions"">Please update the user contact information and press <strong>Save Changes</strong> when finished making changes.</p></td></tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">First Name:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userfname"" type=""text"" class=""correctionstextbox"" value=""" & oUser("userfname") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Last Name:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userlname"" type=""text"" class=""correctionstextbox"" value=""" & oUser("userlname") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Business Name:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userbusinessname"" type=""text"" class=""correctionstextbox"" value=""" & oUser("userbusinessname") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Email:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""useremail"" type=""text"" class=""correctionstextbox"" value=""" & oUser("useremail") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Daytime Phone:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userhomephone"" type=""text"" class=""correctionstextbox"" value=""" & FormatPhone(oUser("userhomephone")) & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Fax:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userfax"" type=""text"" class=""correctionstextbox"" value=""" & FormatPhone(oUser("userfax")) & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Address:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""useraddress"" type=""text"" class=""correctionstextbox"" value=""" & oUser("useraddress") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">City:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""usercity"" type=""text"" class=""correctionstextbox"" value=""" & oUser("usercity") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">State / Province:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userstate"" type=""text"" class=""correctionstextbox"" value=""" & oUser("userstate") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Zip / Postal Code:</td>" & vbcrlf
        response.write "      <td width=""100%""><input name=""userzip"" type=""text"" class=""correctionstextbox"" value=""" & oUser("userzip") & """ /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
								response.write "  <tr>" & vbcrlf
        response.write "      <td class=""correctionslabel"">Preferred Contact Method:</td>" & vbcrlf
        response.write "      <td width=""100%"">" & vbcrlf
        response.write "          <select name=""contactmethodid"" id=""contactmethodid"">" & vbcrlf
								                            subListContactMethods oUser("contactmethodid")
        response.write "          </select>" & vbcrlf
								response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
    				response.write "</table>" & vbcrlf
        response.write "</fieldset>" & vbcrlf
    				response.write "</form>" & vbcrlf
  			else
    				response.write "<p><div class=""correctionsboxnotfound"">No information available for this request.</div></p>" & vbcrlf
  			end if
  end if

end sub

'------------------------------------------------------------------------------
function FormatPhone( Number )
  if Len(Number) = 10 then
     FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  else
     FormatPhone = Number
  end if
end function

'------------------------------------------------------------------------------
sub subListContactMethods(iSelected)

	sSQL = "SELECT * FROM egov_contactmethods ORDER BY contactdescription"

	set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN"), 3, 1

	response.write "  <option value=""0"" selected=""selected"">Please select a contact method...</option>" & vbcrlf

	if not oMethods.eof then
	   do while not oMethods.eof
    			if iSelected = oMethods("rowid") then
      				sSelected = " selected=""selected"""
       else
      				sSelected = ""
       end if

    			response.write "  <option value=""" &  oMethods("rowid") & """" & sSelected & ">" & oMethods("contactdescription") & "</option>" & vbcrlf

    			oMethods.movenext
   	loop
	end if

	oMethods.close
	set oMethods = nothing 

end sub

'------------------------------------------------------------------------------
sub displayButtons(iRequestID)

  response.write "<p>" & vbcrlf
  response.write "  <input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='../action_respond.asp?control=" & iRequestID & "';"" />" & vbcrlf
  response.write "  <input type=""submit"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "</p>" & vbcrlf

end sub
%>