<!DOCTYPE html>
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="poolpass_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: select_members.asp
' AUTHOR: Steve Loar
' CREATED: 01/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/27/06 Steve Loar - INITIAL VERSION
' 1.1  09/11/08 David Boyer - Added Membership Renewals
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim oMembership, iUserId, iOrgId, sUserType, iRateId
 dim sMessage, sDescription, iMaxSignUps, nAmount, iFamilyCount, iPeriodId, iMembershipId

 set oMembership = New classMembership

	session("ManageURL")    = ""

'If they do not have a userid set, take them to the login page automatically
 if request.cookies("userid") = "" or request.cookies("userid") = "-1" then
	   session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
   	response.redirect "../user_login.asp"
 end if

'Determine if this is a purchase or a renewal.
'  - If the "poolpassid" is NULL then it is a "purchase".
'  - Otherwise, if there is a value then it is a "renewal".
 if trim(request("poolpassid")) <> "" then
    lcl_poolpassid   = trim(request("poolpassid"))
    lcl_renewal_text = "Renewal"

    getPoolPassInfo lcl_poolpassid, _
                    iUserId, _
                    iRateId, _
                    iMembershipId, _
                    iPeriodId, _
                    lcl_isSeasonal

   'If this is a renewal membership purchase then the startdate will be the expirationdate+1 of the membership the user is renewing from.
   'If this is a new membership purchase then startdate will be the current date.
    lcl_startdate      = oMembership.getMembershipStartDate(lcl_poolpassid)
    lcl_expirationdate = datevalue(getPoolPassExpirationDate(iorgid, lcl_poolpassid))

 else
    if request("userid")     = "" _
    or request("usertype")   = "" _
    or request("periodid")   = "" _
    or request("isSeasonal") = "" then
       response.redirect "poolpass_select.asp"
    else
       lcl_poolpassid     = ""
       lcl_renewal_text   = "Purchase"
      	iUserId            = request("userid")
      	sUserType          = request("usertype")
      	iRateId            = request("rateid")
      	iPeriodId          = request("periodid")
       lcl_isSeasonal     = request("isSeasonal")
       lcl_startdate      = date()
       lcl_expirationdate = oMembership.getExpirationDate(iPeriodID, lcl_startdate)

       if request("imembershipid") <> "" then
         	iMembershipId = CLng(request("imembershipid"))
       elseif request("membershipid") <> "" then
         	iMembershipId = CLng(request("membershipid"))
      	else

          iMembershipId = oMembership.getMembershipIdByMembership("pool")
       end if
    end if
 end if

 mtype = "pool"
 sSQL = "SELECT membership FROM egov_memberships WHERE membershipid = " & iMembershipId
 Set oM = Server.CreateObject("ADODB.RecordSet")
 oM.Open sSQL, Application("DSN"), 3, 1
 if not oM.EOF then mtype = oM("membership")
 oM.Close
 Set oM = Nothing

'Get the rate name and message
	getRateInfo iRateId, _
             nAmount, _
             sMessage, _
             sDescription, _
             iMaxSignUps, _
             iAttendanceTypeID, _
             lcl_rate_residenttype, _
             lcl_rate_residenttypedesc, _
             lcl_isPunchcard, _
             lcl_punchcard_limit

 if sUserType = "" then
    sUserType = lcl_rate_residenttype
 end if

'Default the attendancetypeid to Member (1) if for whatever reason it is empty
 if iAttendanceTypeID = "" OR isnull(iAttendanceTypeID) then
    iAttendanceTypeID = 1
 end if

 if iMaxSignUps = "" then
	   iMaxSignUps = 1
 end if

 if nAmount = "" then
   	nAmount = 001.00
 end if

	sRateName    = ""
	iFamilyCount = 0

 if DATEDIFF("d",lcl_expirationdate,date()) > 0 then
    lcl_startdate = date()
 end if

'Set up the TITLE tag
 lcl_title = "E-Gov Services " & sOrgName & " Membership Purchase"

 if iorgid = 7 then
    lcl_title = sOrgName
 end if

	session("RedirectPage") = "pool_pass/select_members.asp?iuserid="& iUserId & "&iorgid=" & iOrgId & "&usertype=" & sUserType & "&rateid=" & iRateId & "&periodid=" & iPeriodId & "&membershipid=" & iMembershipId & "&isseasonal=" & lcl_isSeasonal
	session("RedirectLang") = "Return to Membership Purchase"

%>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
  <title><%=lcl_title%></title>

  <link rel="stylesheet" href="../css/styles.css" />
  <link rel="stylesheet" href="../global.css" />
  <link rel="stylesheet" href="./style_pool.css" />
  <link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

<style>
  #membersTable {
     border:           1pt solid #000000;
     background-color: #ffffff;
  }

  th#membersTableBorders {
     border-bottom:    1pt solid #000000;
  }
</style>

<script src="../scripts/modules.js"></script>
<script src="../scripts/easyform.js"></script>

<script>
<!--
	function UpdateFamily(iUserId)	{
 		//location.href='../family_members.asp?userid=' + iUserId;
	 	location.href='../family_list.asp?userid=' + iUserId;
	}

	function ContinuePurchase()
	{

		// set var checkbox_choices to zero
		var checkbox_choices = 0;

		//alert(document.all.addPoolPass.passIncl.length);

		if (document.addPoolPass.passIncl.length > 1)
		{
			// Loop from zero to the one minus the number of checkbox button selections
			for (counter = 0; counter < document.addPoolPass.passIncl.length; counter++)
			{	

				// If a checkbox has been selected it will return true
				// (If not it will return false)
				if (document.addPoolPass.passIncl[counter].checked)
				{ 
					checkbox_choices = checkbox_choices + 1; 
				}
			}
		}
		else
		{
			if (document.addPoolPass.passIncl.checked)
			{
				checkbox_choices = checkbox_choices + 1;
			}
		}

		if (checkbox_choices == 0)
		{
			alert( "Please select at least one name to include on the pass.");
			return;
		}
		else if (checkbox_choices > document.addPoolPass.imaxsignups.value)
		{
			alert("You are limited to " + document.addPoolPass.imaxsignups.value + " person(s) for this type of pass. \n Please deselect some, or choose another type of pass.");
			return;
		}
		else if (document.addPoolPass.imaxsignups.value < 5 && checkbox_choices != document.addPoolPass.imaxsignups.value)
		{
			alert("This pass is for " + document.addPoolPass.imaxsignups.value + " person(s). \n Please select more, or choose another type of pass.");
			return;
		}
		else
		{
			//alert("Pass is OK.");
			document.addPoolPass.submit();
		}
	}

	function GoBack() {
 		location.href='poolpass_select.asp?mtype=<%=mtype%>'
	}
  //-->
 </script>
</head>

<!--#Include file="../include_top.asp"-->
<%
 'Set up the membership, rate, and price information
  lcl_membershiprate_details = "You have selected the&nbsp;"
  lcl_membershiprate_details = lcl_membershiprate_details & oMembership.GetMembershipPeriodName( iPeriodId )
  lcl_membershiprate_details = lcl_membershiprate_details & sRateName & "&nbsp;"
  lcl_membershiprate_details = lcl_membershiprate_details & oMembership.GetMembershipNameById( iMembershipId ) & "&nbsp;"
  lcl_membershiprate_details = lcl_membershiprate_details & "Membership at " & FormatCurrency(nAmount,2)

 'Setup "Start Date" and "Start Date Title"
  if NOT lcl_isSeasonal then
     lcl_displayStartDate = lcl_startdate
  else
     lcl_displayStartDate = year(lcl_startdate) & " Season"
  end if

  if lcl_poolpassid <> "" then
     if NOT lcl_isSeasonal then
        lcl_startdate_title = "Membership Renewal Start Date"
     else
        lcl_startdate_title = "Renew for Membership Season"
     end if
  else
     lcl_startdate_title = "Start Date"
  end if

  RegisteredUserDisplay( "../" )

  response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Back"" class=""reserveformbutton"" onclick=""GoBack();"" /><br />" & vbcrlf

  response.write "<form method=""post"" name=""addPoolPass"" id=""addPoolPass"" action=""pool_pass_add.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""sVirtualSite"" id=""sVirtualSite"" value="""       & sorgVirtualSiteName & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iuserid"" id=""iuserid"" value="""                 & iUserId             & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value="""                   & iOrgId              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""usertype"" id=""usertype"" value="""               & sUserType           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rateid"" id=""rateid"" value="""                   & iRateId             & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""imaxsignups"" id=""imaxsignups"" value="""         & iMaxSignUps         & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""membershipid"" id=""membershipid"" value="""       & iMembershipId       & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""periodid"" id=""periodid"" value="""               & iPeriodId           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""poolpassid"" id=""poolpassid"" value="""           & lcl_poolpassid      & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""amount"" id=""amount"" value="""                   & nAmount             & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""isSeasonal"" id=""isSeasonal"" value="""           & lcl_isSeasonal      & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""isPunchcard"" id=""isPunchcard"" value="""         & lcl_isPunchcard     & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""punchcard_limit"" id=""punchcard_limit"" value=""" & lcl_punchcard_limit & """ />" & vbcrlf

  response.write "<div id=""content"">" & vbcrlf

 'BEGIN: Membership Details, Rate, and Price info -----------------------------
  lcl_showMessage = checkToShowMessage(iMembershipID)

  response.write "	 <div class=""reserveformtitle"">" & lcl_membershiprate_details & "</div>" & vbcrlf
  response.write "	 <div class=""reserveforminputarea"">" & vbcrlf
  response.write "		  <p>Please select names from the list below to be included in the&nbsp;" & oMembership.GetMembershipNameById( iMembershipId ) & "&nbsp;Membership.</p>" & vbcrlf

  if lcl_showMessage AND trim(sMessage) <> "" then
     response.write "		  <p>" & sMessage & "</p>" & vbcrlf
  end if

  response.write " 	</div>" & vbcrlf
 'END: Membership Details, Rate, and Price info -------------------------------

 'BEGIN: Available Family Members ---------------------------------------------
  response.write "  <div class=""reserveformtitle"">Available Family Members</div>" & vbcrlf
  response.write "  	 <div class=""reserveforminputarea"">" & vbcrlf
  response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "        <tr>" & vbcrlf
  response.write "            <td><input type=""button"" name=""updateFamilyMembersButton"" id=""updateFamilyMembersButton"" value=""Update Family Members"" class=""reserveformbutton"" onclick=""UpdateFamily(" & iUserId & ");"" /></td>" & vbcrlf
  response.write "            <td align=""right"">" & vbcrlf
  response.write "                <strong>" & lcl_startdate_title & ": " & lcl_displayStartDate & "</strong>" & vbcrlf
  response.write "                <input type=""hidden"" name=""startdate"" id=""startdate"" value=""" & lcl_startdate & """ size=""10"" maxlength=""10"" />" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "        </tr>" & vbcrlf
  response.write "      </table>" & vbcrlf
  response.write "      <br />" & vbcrlf
  response.write "      <div class=""membersTableContain"">" & vbcrlf
  response.write "      <table id=""membersTable"" border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "        <tr>" & vbcrlf
  response.write "            <th id=""membersTableBorders"">Include</th>" & vbcrlf
  response.write "            <th id=""membersTableBorders"">First Name</th>" & vbcrlf
  response.write "            <th id=""membersTableBorders"">Last Name</th>" & vbcrlf

  if lcl_poolpassid <> "" then
     response.write "            <th id=""membersTableBorders"">Member ID</th>" & vbcrlf
  end if

  response.write "            <th id=""membersTableBorders"">Relation</th>" & vbcrlf
  response.write "            <th id=""membersTableBorders"">Birthdate</th>" & vbcrlf
  response.write "        </tr>" & vbcrlf

                          iFamilyCount = GetFamilyMembers(iUserId, iRateId, lcl_poolpassid)

  response.write "      </table>" & vbcrlf
  response.write "      </div>" & vbcrlf
  response.write "    </div>" & vbcrlf
 'END: Available Family Members -----------------------------------------------

  if iFamilyCount > 0 then
     response.write "<div id=""poolfooter"">" & vbcrlf
     response.write "  <input type=""button"" name=""continueButton"" id=""continueButton"" value=""Continue with Purchase"" class=""reserveformbutton"" onclick=""ContinuePurchase();"" />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../include_bottom.asp"-->  
<%
 set oMembership = nothing

'------------------------------------------------------------------------------
function getPoolPassExpirationDate(p_orgid, p_poolpassid)
  dim lcl_return, sOrgID, sPoolPassID

  lcl_return  = ""
  sOrgID      = 0
  sPoolPassID = 0

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_poolpassid <> "" then
     sPoolPassID = clng(p_poolpassid)
  end if

  sSQL = "SELECT expirationdate "
  sSQL = sSQL & " FROM egov_poolpasspurchases "
  sSQL = sSQL & " WHERE poolpassid = " & sPoolPassID
  sSQL = sSQL & " AND orgid = " & sOrgID

 	set oExpDate = Server.CreateObject("ADODB.Recordset")
	 oExpDate.Open sSQL, Application("DSN"), 3, 1

  if not oExpDate.eof then
     lcl_return = oExpDate("expirationdate")
  end if

  set oExpDate = nothing

  getPoolPassExpirationDate = lcl_return

end function

'------------------------------------------------------------------------------
function checkToShowMessage(iMembershipID)
  dim lcl_return, sMembershipID

  lcl_return    = false
  sMembershipID = 0

  if iMembershipID <> "" then
     sMembershipID = clng(iMembershipID)
  end if

  sSQL = "SELECT showMessage "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE membershipid = " & sMembershipID

 	set oGetShowMessage = Server.CreateObject("ADODB.Recordset")
	 oGetShowMessage.Open sSQL, Application("DSN"), 3, 1

  if not oGetShowMessage.eof then
     if oGetShowMessage("showMessage") then
        lcl_return = true
     end if
  end if

  oGetShowMessage.close
  set oGetShowMessage = nothing

  checkToShowMessage = lcl_return

end function
%>
