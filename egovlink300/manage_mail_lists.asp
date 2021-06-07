<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: manage_mail_lists.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Selection of email subscriptions by citizens.
'
' MODIFICATION HISTORY
' 2.0  02/04/08  David Boyer - Added Job/Bid Postings
' 2.1  2014-06-11  Jerry Felix - revised the email regex to be more permissive for new TLDs
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim sError

 if session("RedirectLang") <> "" then
	   sBackLang     = session("RedirectLang")
   	sRedirectPage = session("RedirectPage")
 else
    session("RedirectLang") = "Back"
    session("RedirectPage") = "manage_mail_lists.asp"
 end if

 if request.cookies("userid") = "" OR request.cookies("userid") = "-1" then
  		response.redirect "user_login.asp"
	end if

 if Request.ServerVariables("REQUEST_METHOD") = "POST" then
   	Call UpdateMailLists(request.cookies("userid"))
 end if

'Set the session for the family update form to come back here
' session("ManageURL")  = "manage_account.asp"
' session("ManageLang") = "Return to Manage Account"

' if session("RedirectLang") <> "" then
'	   sBackLang     = session("RedirectLang")
'   	sRedirectPage = session("RedirectPage")
' else
'	   sBackLang     = "Back"
'   	sRedirectPage = "manage_account.asp"
' end if

' session("RedirectLang") = sBackLang
' session("RedirectPage") = sRedirectPage

'Check for org features
 lcl_orghasfeature_subscriptions = orghasfeature(iorgid,"subscriptions")
 lcl_orghasfeature_job_postings  = orghasfeature(iorgid,"job_postings")
 lcl_orghasfeature_bid_postings  = orghasfeature(iorgid,"bid_postings")
 lcl_orghasfeature_subscriptions_distributionlist_showdesc = orghasfeature(iorgid,"subscriptions_distributionlist_showdesc")
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
<title>E-Gov Services <%=sOrgName%> - Manage Account</title>

<link rel="stylesheet" type="text/css" href="css/styles.css" />
<link rel="stylesheet" type="text/css" href="global.css" />
<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

<script language="javascript" src="scripts/modules.js"></script>
<script language="javascript" src="scripts/easyform.js"></script>

<script language=javascript>
<!--

	function openWin2(url, name) 
	{
	  popupWin = window.open(url, name,"resizable,width=500,height=450");
	}

	function UpdateFamily(iUserId)
	{
		location.href='family_members.asp?userid=' + iUserId;
	}

	function Validate() 
	{
		var msg="";
		if(document.register.egov_users_userpassword.value != document.register.skip_userpassword2.value)
		{
			msg+="The passwords you have entered do not match.\n";
		}

		//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz))$/;
		var rege = /.+@.+\..+/i;
		var Ok = rege.test(document.register.egov_users_useremail.value);

		if (! Ok)
		{
			msg+="The email must be in a valid format.\n";
		}

		// set the work phone
		if (document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "")
		{
			var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
			if (sPhone.length < 10)
			{
				msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userworkphone.value);
				if ( ! Ok )
				{
					msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the fax
		if (document.register.skip_fax_areacode.value != "" || document.register.skip_fax_exchange.value != "" || document.register.skip_fax_line.value != "" )
		{
			var sPhone = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
			if (sPhone.length < 10)
			{
				msg += "Fax must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_userfax.value = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userfax.value);
				if ( ! Ok )
				{
					msg += "Fax must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the home phone number
		document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;
		if (document.register.egov_users_userhomephone.value != "" )
		{
			var hPhone = document.register.egov_users_userhomephone.value;
			if (hPhone.length < 10)
			{
				msg += "The home phone must be a valid phone number, including area code.\n";
			}
			else
			{
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userhomephone.value);
				if ( ! Ok )
				{
					msg += "The home phone must be a valid phone number, including area code.\n";
				}
			}
		}
		else
		{
			msg+="The home phone cannot be blank.\n";
		}

		// Process the business address if one was chosen
		var bexists = eval(document.register["skip_Baddress"]);
		if(bexists)
		{
			//See if they picked from the business dropdown and put that in the address field 
			if (document.register.skip_Baddress.selectedIndex > -1)
			{
				var belement = document.register.skip_Baddress;
				var bselectedvalue = belement.options[belement.selectedIndex].value;

				//alert( bselectedvalue );
				//  0000 is the first pick that we do not want
				if (bselectedvalue != "0000")
				{
					document.register.egov_users_userbusinessaddress.value = bselectedvalue;
					document.register.egov_users_residenttype.value = "B";
				}
			}
		}

		// Process the resident address if one was chosen - this is second to set the local resident type
		var exists = eval(document.register["skip_Raddress"]);
		if(exists)
		{
			// See if they picked from the resident dropdown and put that in the address field 
			if (document.register.skip_Raddress.selectedIndex > -1)
			{
				var element = document.register.skip_Raddress;
				var selectedvalue = element.options[element.selectedIndex].value;

				//alert( selectedvalue );
				//  0000 is the first pick that we do not want
				if (selectedvalue != "0000")
				{
					document.register.egov_users_useraddress.value = selectedvalue;
					document.register.egov_users_residenttype.value = "R";
				}
			}
		}

		if(msg != "")
		{
			msg="Your form could not be submitted for the following reasons.\n\n" + msg;
			alert(msg);
			return;
		}
		else {	
			if (validateForm('register')) 
			{ 
				document.register.submit(); 
			}
		}
	}

	function GoBack(ReturnToURL)
	{
		if (ReturnToURL != "")
		{
			location.href=ReturnToURL;
		}
		else
		{
			history.go(-1);
		}
	}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e) 
	{
		var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
			input.value = input.value.slice(0, len);
		var addNdx = 1;

		while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
		{
			addNdx++;
			//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
		}

		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
	}

	function containsElement(arr, ele) 
	{
		var found = false, index = 0;

		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
		return found;
	}

	function getIndex(input) 
	{
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
	}
		return true;
	}

//-->
</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->

<div align="left" style="padding-bottom:20px;"> <% RegisteredUserDisplay("") %> </div>
<p>

<div class="indent20">
	 <a href="javascript:GoBack('<%=Session("RedirectPage")%>')"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=session("RedirectLang")%></a>
<p>

<div class="box_header4"><%=sOrgName%> - Subscribe to Email Communications: </div>

<div class="groupSmall2">
<form name="register" action="manage_mail_lists.asp" method="post">
	
<p>
  Check or uncheck the box next to the mail list you would like to subscribe\unsubscribe. 
  Then, press the <strong> Update Mailing Lists </strong> button.
</p>

<%
  'Display all distributionlists
   DisplayMaillists request.cookies("userid")
%>
</form>
</div>
</div>
</div>
<p>
   
<!-- #include file="include_bottom.asp" -->
<!-- #include file="includes\inc_dbfunction.asp" -->
<%
'------------------------------------------------------------------------------
sub DisplayMaillists(iuserid)
	Dim sSql, oList

'Retrieve all of the distribution lists, job postings, and bid postings.	
 sSQL = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
 sSQL = sSQL & " distributionlistdisplay, orgid, isnull(distributionlisttype,'') as distributionlisttype, parentid "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE orgid = '" & iorgid & "' "
 sSQL = sSQL & " AND distributionlistdisplay = 1 "
 sSQL = sSQL & " AND parentid is null "
 sSQL = sSQL & " ORDER BY distributionlisttype, distributionlistname "

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

 response.write "<fieldset style=""width: 95%"">" & vbcrlf
 response.write "  <legend>Subscriptions&nbsp;</legend><p>" & vbcrlf

 lcl_listtype_prev = "LIST"
 lcl_line_count    = 0

	if NOT oList.eof then
		  while NOT oList.EOF

       lcl_line_count = lcl_line_count + 1

       if oList("distributionlisttype") = "JOB" then
          lcl_listtype   = oList("distributionlisttype")
          lcl_list_title = getFeatureName("job_postings")
          'lcl_list_title = "JOB POSTINGS"
       elseif oList("distributionlisttype") = "BID" then
          lcl_listtype   = oList("distributionlisttype")
          lcl_list_title = getFeatureName("bid_postings")
          'lcl_list_title = "BID POSTINGS"
       else
          lcl_listtype   = "LIST"
          lcl_list_title = "DISTRIBUTION LISTS"
       end if

      'Determine which features the org has "turned-on"
       if lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions then
          lcl_show = "Y"
       elseif lcl_listtype = "JOB" AND lcl_orghasfeature_job_postings then
          lcl_show = "Y"
       elseif lcl_listtype = "BID" AND lcl_orghasfeature_bid_postings then
          lcl_show = "Y"
       else
          lcl_show = "N"
       end if

      'If the current parent (category) is different then the previous record then reset the variables
      'if (isnull(lcl_listtype_prev) OR lcl_listtype_prev <> oList("distributionlisttype")) AND lcl_line_count > 1 then
       if lcl_line_count > 1 then
          if isnull(lcl_listtype_prev) then
             lcl_listtype_prev = "LIST"
          end if

          if lcl_listtype_prev <> lcl_listtype then
             lcl_line_count = 1
          end if
       end if

       if lcl_line_count = 1 then
          if lcl_list_title <> "" then
'             if lcl_listtype <> "LIST" then
'                response.write "</div>" & vbcrlf
'             end if

             if lcl_show = "Y" then
                response.write "<p>" & vbcrlf
                response.write "<strong>" & UCASE(lcl_list_title) & "</strong>" & vbcrlf
                response.write "<hr size=""1"" width=""100%"">" & vbcrlf
                response.write "<div id=""" & lcl_listtype & """>" & vbcrlf
             end if
          end if
       end if

       lcl_listtype_prev = lcl_listtype

       if lcl_show = "Y" then
       			if IsMember( iuserid, oList("distributionlistid") ) then
         				sChecked = " checked=""checked"""
          else
             sChecked = ""
        		end if

       			response.write "<input type=""checkbox"" name=""maillist"" value=""" & oList("distributionlistid") & """" & sChecked & " />" & vbcrlf

         'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
          lcl_desc = ""

          if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
             if oList("distributionlistdescription") <> "" then
                lcl_desc = " <i>(" & oList("distributionlistdescription") & ")</i>"
             else
                lcl_desc = ""
             end if
          else
             lcl_desc = ""
          end if

          response.write "<strong>" & oList("distributionlistname") & "</strong>" & lcl_desc & vbcrlf
          response.write "<br />" & vbcrlf

         'Check for any sub-categories
         	sSQL2 = "SELECT * "
     	    sSQL2 = sSQL2 & " FROM egov_class_distributionlist "
        	 sSQL2 = sSQL2 & " WHERE orgid = " & iorgid
        	 sSQL2 = sSQL2 & " AND distributionlistdisplay = 1 "
  	       sSQL2 = sSQL2 & " AND UPPER(distributionlisttype) = '" & UCASE(oList("distributionlisttype")) & "' "
          sSQL2 = sSQL2 & " AND parentid = " & oList("distributionlistid")
          sSQL2 = sSQL2 & " ORDER BY UPPER(distributionlistname) "

        		set rs2 = Server.CreateObject("ADODB.Recordset")
        		rs2.Open sSQL2, Application("DSN"), adOpenForwardOnly, adLockReadOnly

          if not rs2.eof then
           		while not rs2.eof
                response.write "<span style=""margin-left:15pt"">" & vbcrlf
                response.write "<input name=""maillist"" type=""checkbox"" value=""" & rs2("distributionlistid") & """"

        	      	If IsMember( iuserid, rs2("distributionlistid") ) Then
               				response.write " checked=""checked"" "
             			End If

               'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
                lcl_desc = ""

    						      if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
                   if oList("distributionlistdescription") <> "" then
                      lcl_desc = " <i>(" & oList("distributionlistdescription") & ")</i>"
                   else
                      lcl_desc = ""
                   end if
                else
                   lcl_desc = ""
                end if

       			      response.write " />&nbsp;<strong>" & rs2("distributionlistname") & "</strong>" & lcl_desc & vbcrlf
                response.write "</span>" & vbcrlf
                response.write "<br />" & vbcrlf
             			rs2.movenext
           		wend
          end if
       end if

       lcl_listtype_prev = lcl_listtype

    			oList.MoveNext
  		wend

  		response.write "<p><input type=""submit"" value=""Update Mailing Lists"" class=""button""></p>" & vbcrlf
	else
		  response.write "<p>There are no mailing lists available at this time.</p>" & vbcrlf
	end If

 response.write "</fieldset><p>" & vbcrlf

	oList.close 
	Set oList = Nothing 

End Sub

'------------------------------------------------------------------------------
' SUB UPDATEMAILLISTS(IUSERID)
'------------------------------------------------------------------------------
Sub UpdateMailLists( iuserid )
	Dim sSql, oCmd

	' CLEAR CURRENT LISTS
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "DELETE FROM egov_class_distributionlist_to_user WHERE userid = '" & iuserid & "'"
		.Execute
	End With

	' RECREATE CURRENT LIST
	For Each list In request("maillist")
		oCmd.CommandText = "INSERT INTO egov_class_distributionlist_to_user (userid,distributionlistid) VALUES ('" & iuserid & "','" & list & "')"
		oCmd.Execute
	Next	
	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------
' FUNCTION ISMEMBER(USERID,LISTID)
'------------------------------------------------------------------------------
Function IsMember( iuserid, listid )
	Dim sSQL, oList

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_class_distributionlist_to_user "
 sSQL = sSQL & " WHERE (userid = '" & iuserid & "' "
 sSQL = sSQL & " AND distributionlistid='" & listid & "')"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If NOT oList.EOF Then
		IsMember = True
	Else
		IsMember = False 
	End If
	oList.close 
	Set oList = Nothing

End Function
%>
