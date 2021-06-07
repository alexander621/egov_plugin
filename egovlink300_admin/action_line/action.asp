<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_rye.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: action.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module is the New Action Line Requests.
'
' MODIFICATION HISTORY
' 1.0	10/16/06	 Steve Loar  - Security, Header and Nav changed
' 2.0	05/07/07  Steve Loar  - Changes to problem location to handle larger cities
' 3.0 11/05/07  David Boyer - Incorporated "Validate Address" feature
' 4.0 04/08/08  David Boyer - Added new address format to issue location
' 4.1 08/19/08  David Boyer - Added javascript field length check to textarea fields.
' 4.2 08/28/08  David Boyer - Added new form validation
' 4.3 05/29/09  David Boyer - Added check to see if "Additional Information" textarea is displayed or not.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../" ' Override of value from common.asp
 lcl_hidden = "hidden" 'Hides/Shows all hidden fields.  Hide=HIDDEN, Show=TEXT

 if NOT UserHasPermission( session("userid"), "create requests" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 Dim sError
 iorgid = session("orgid")

'Check for org features
 lcl_orghasfeature_large_address_list                    = orghasfeature("large address list")
 lcl_orghasfeature_issue_location                        = orghasfeature("issue location")
 lcl_orghasfeature_actionline_donotsend_submissionemails = orghasfeature("actionline_donotsend_submissionemails")

'Set the field lengths for the custom/internal fields
 lcl_text_field_length     = 1024
 lcl_textarea_field_length = 4000
%>
<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

 	<script type="text/javascript" src="../scripts/modules.js"></script>
  <script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="../scripts/removespaces.js"></script>
  <script type="text/javascript" src="../../scripts/easyform.js"></script>
  <script type="text/javascript" src="../scripts/setfocus.js"></script>
  <script type="text/javascript" src="../scripts/selectAll.js"></script>
 	<script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>
  <script type='text/javascript' src='../scripts/jquery.ui.core.min.js?ver=1.10.3'></script>
  <script type='text/javascript' src='../scripts/jquery.ui.position.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='../scripts/jquery.ui.widget.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='../scripts/jquery.ui.menu.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='../scripts/jquery.ui.autocomplete.min.js?ver=1.10.3'></script>

  <script type="text/javascript" src="../scripts/zebra_datepicker.js"></script>
  <link rel="stylesheet" href="../css/zebra_datepicker.css" type="text/css">
  <link rel='stylesheet' id='smoothness-css'  href='https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css?ver=3.6.1' type='text/css' media='all' />

<script type="text/javascript">
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

<%
'------------------------------------------------------------------------------
 if lcl_orghasfeature_large_address_list then
'------------------------------------------------------------------------------
%>
function checkAddress( sReturnFunction, sSave ) {
		// Remove any extra spaces
		document.frmRequestAction.residentstreetnumber.value = removeSpaces(document.frmRequestAction.residentstreetnumber.value);

		// check the number for non-numeric values
		var rege = /^\d+$/;
		var Ok = rege.exec(document.frmRequestAction.residentstreetnumber.value);

  if(document.frmRequestAction.ques_issue2.value=="") {
		   if ( ! Ok ) {
       //alert("The Resident Street Number cannot be blank and must be numeric.");
	 	   	setfocus(document.frmRequestAction.residentstreetnumber);
   			 inlineMsg(document.getElementById("residentstreetnumber").id,'The Resident Street Number cannot be blank and must be numeric',10,'residentstreetnumber');
   		 	return false;
   		}

   		// check that they picked a street name
   		if ( document.frmRequestAction.skip_address.value == '0000') {
       //alert("Please select a street name from the list first.");
   	 		setfocus(document.frmRequestAction.skip_address);
   			 inlineMsg(document.getElementById("skip_address").id,'Please select a street name from the list first',10,'skip_address');
   		 	return false;
   		}

   		// This is here because window.open in the Ajax callback routine will not work
   		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=frmRequestAction", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
   		//self.focus();
   		// Fire off Ajax routine

   		doAjax('checkaddress.asp', 'stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value, sReturnFunction, 'get', '0');
  }else{
     if(document.frmRequestAction.residentstreetnumber.value!="" || document.frmRequestAction.skip_address.value!="0000") {
        document.frmRequestAction.ques_issue2.value="";
      		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=frmRequestAction", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		doAjax('checkaddress.asp', 'stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&orgid=<%=iorgid%>', sReturnFunction, 'get', '0');
     }else{
        FinalCheck('NOT FOUND');
     }
  }
}

function CheckResults( sResults ) {
  // Process the Ajax CallBack when the validate address button is clicked
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
	 	  	document.frmRequestAction.ques_issue2.value = '';
      document.frmRequestAction.validstreet.value = 'Y';
  		 	alert("This is a valid address in the system.");
  }else{
      document.frmRequestAction.validstreet.value = 'N';
 	  		//winHandle.focus();
   			PopAStreetPicker('CheckResults', 'no');
  }
}

function FinalCheck( sResults ) {
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
      document.frmRequestAction.validstreet.value = 'Y';
      document.frmRequestAction.submit();
  }else{
      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {
     		    //if(winHandle != null && ! winHandle.closed) { 
        	  //			winHandle.close();
   			     //}

           if (sResults == 'FOUND SELECT') {
               document.frmRequestAction.validstreet.value = 'Y';
           }else{
               document.frmRequestAction.validstreet.value = 'N';
           }

           document.frmRequestAction.submit();
      }else{
           document.frmRequestAction.validstreet.value = 'N';
         		//if(winHandle != null && ! winHandle.closed) { 
           //   winHandle.focus();
           //}else{
           //   document.frmRequestAction.submit();
           //}
           if(document.frmRequestAction.ques_issue2.value!="") {
              document.frmRequestAction.submit();
           } else {
             	PopAStreetPicker('FinalCheck', 'yes');
           }
      }
  }
}

function PopAStreetPicker( sReturnFunction, sSave )	{
		// pop up the address picker
  winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=frmRequestAction", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}
<%
'------------------------------------------------------------------------------
 end if
'------------------------------------------------------------------------------
%>
function validateCheckbox(p_field) {
  //get the total options
  var lcl_total_options = document.getElementById("total_options_"+p_field).innerHTML;
  var lcl_total_checked = 0;

  for (i = 1; i <= lcl_total_options-1; i++) {
       if(document.getElementById(p_field+"_"+i).checked==true) {
          lcl_total_checked = lcl_total_checked + 1;
       }
  }
  if(lcl_total_checked == 0) {
     document.getElementById(p_field+"_"+lcl_total_options).checked=true;
  }else{
     document.getElementById(p_field+"_"+lcl_total_options).checked=false;
  }
}

function disableSelfSendChkbox() {

 if(document.getElementById("doNotSendAllEmail").checked==true) {
    document.getElementById("doNotSendSelfEmail").disabled=true;
    document.getElementById("doNotSendSelfEmail").checked=false;
 }else{
    document.getElementById("doNotSendSelfEmail").disabled=false;
 }

}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();">
  <%'DrawTabs tabRequests,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<%
 response.write "<div id=""content"">" & vbcrlf
 response.write "  <h3>Create Action Line Request</h3>" & vbcrlf
 response.write "  <div id=""centercontent"">" & vbcrlf
 response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"">" & vbcrlf

'------------------------------------------------------------------------------
 if trim(request("actionid")) <> "" then
 	'Display the requested action line form

	  Dim sDefaultCity,sDefaultState,sDefaultZip
	  GetDefaultValues()

	  Call subDisplayActionForm(request("actionid"),iorgid)

 else

	 'Display the list of available forms
   response.write "<table cellspacing=""1"" cellpadding=""1"" border=""0"">" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <td valign=""top"">" & vbcrlf
   response.write "          <table cellspacing=""1"" cellpadding=""1"" border=""0"">" & vbcrlf
   response.write "            <tr valign=""top"">" & vbcrlf
   response.write "                <td>" & vbcrlf

  'BEGIN: List Forms
   response.write "                    <form action=""action.asp?list=true"" method=""post"">" & vbcrlf
   response.write "                    <table cellspacing=""0"" border=""0"" class=""tablelist"">" & vbcrlf
   response.write "                      <tr><th>Request List</th></tr>" & vbcrlf
   response.write "                      <tr>" & vbcrlf
   response.write "                          <td nowrap=""nowrap"">" & vbcrlf
                                                 fnListForms()
   response.write "                          </td>" & vbcrlf
   response.write "                      </tr>" & vbcrlf
   response.write "                    </table>" & vbcrlf
   response.write "                    </form>" & vbcrlf
  'END: List Forms

   response.write "                </td>" & vbcrlf
   response.write "            </tr>" & vbcrlf
   response.write "          </table>" & vbcrlf
   response.write "      </td>" & vbcrlf
   response.write "  </tr>" & vbcrlf
   response.write "</table>" & vbcrlf

 end if

 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
Function fnListForms()
	
	sLastCategory = "NONE_START"

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM dbo.egov_form_list_200 "
 sSQL = sSQL & " WHERE (orgid=" & iorgID & ") "
 sSQL = sSQL & " AND (form_category_id <> 6) "
 sSQL = sSQL & " ORDER BY form_category_Sequence, action_form_name "

	Set oForms = Server.CreateObject("ADODB.Recordset")
	oForms.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForms.EOF Then
  		while NOT oForms.eof

   			'DETERMINE IF PUBLIC OR INTERNAL ONLY
    			if oForms("action_form_internal") then
      				sInternal = " - <b> (Internal Only) </b>"
    			else
      				sInternal = ""
			    end if

      'Set the variable for the currrent category
    			sCurrentCategory = oForms("form_category_name")

      'Set the variable to be used to display the category
       lcl_display_category = "<b><a name=""" & oForms("form_category_id")  & """>" & sCurrentCategory & "</a></b>" & vbcrlf

      'If the current category does NOT match the previous category then insert a paragraph and 
      'start the next list of forms for the new category.
    			if (sCurrentCategory <> sLastCategory) then

          'if this is NOT the first category in the list
           if sLastCategory <> "NONE_START" then
            		response.write "</ul>" & vbcrlf
           end if

           response.write "<p></p>" & vbcrlf
          	response.write "<ul>" & vbcrlf
           response.write lcl_display_category
    			end if

    			sTopic = Server.URLEncode(sCurrentCategory & " > " & oForms("action_form_name"))
		
    			response.write "<li><a href=""action.asp?actionid=" & oForms("action_form_id") & """>" & oForms("action_form_name") & sInternal&  "</a></li>" & vbcrlf

    			sLastCategory = sCurrentCategory

       oForms.MoveNext
    wend

    response.write "</p>" & vbcrlf
  		response.write "</ul>" & vbcrlf
	else
  		response.write "<p style=""padding-top:10px;""><center><font color=""red""><b><i>No action forms enabled.</i></b></font></p>" & vbcrlf
	end if

	oForms.close
	Set oForms = Nothing 

End Function

'------------------------------------------------------------------------------
Sub subDisplayActionForm(iFormID,iorgid)

'GET FORM GENERAL INFORMATION
	Dim sTitle
	Dim sIntroText
	Dim sFooterText
	Dim sMask
	Dim blnEmergencyNote
	Dim sEmergencyText
	Dim blnIssueDisplay
	Dim sIssueMask
	Dim	iStreetNumberInputType 
	Dim	iStreetAddressInputType 
	Dim sIssueName
	Dim sIssueDesc
	Dim sIssueQues
 Dim sHideIssueLocAddInfo

	' GET FORM INFORMATION	
	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id=" & iFormID

	Set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForm.EOF Then
   'POPULATE DATA FROM RECORDSET
  		sTitle                  = oForm("action_form_name")
		  sIntroText              = oForm("action_form_description")
  		sFooterText             = oForm("action_form_footer")
		  sMask                   = oForm("action_form_contact_mask")
  		blnEmergencyNote        = oForm("action_form_emergency_note")
		  sEmergencyText          = oForm("action_form_emergency_text")
  		blnIssueDisplay         = oForm("action_form_display_issue")
		  sIssueMask              = oForm("action_form_issue_mask")
  		iStreetNumberInputType  = oForm("issuestreetnumberinputtype")
		  iStreetAddressInputType = oForm("issuestreetaddressinputtype")
  		sIssueName              = oForm("issuelocationname")
    sHideIssueLocAddInfo    = oForm("hideIssueLocAddInfo")

		  If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
       sIssueName = "Issue/Problem Location:"
  		End If

		  sIssueDesc = oForm("issuelocationdesc")

  		If IsNull(sIssueDesc) Then
       'sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". Provide any additional information on problem location in the box below."
       sIssueDesc = "Please select the closest street number/streetname of problem location from list or select"
       sIssueDesc = sIssueDesc & " ""Choose street from dropdown"". Provide any additional information on problem location in the box below."
		  End If

  		sIssueQues = oForm("issuequestion")

    'If trim(sIssueQues) = "" OR IsNull(sIssueQues) Then
    '	  sIssueQues = "Provide any additional information on problem location in the box below."
    'End If

    'Determine if the Contact Info fields are displayed
     lcl_checkForContactFields = 0
     lcl_displayContactFields  = False
     lcl_displayContactEmail   = False

     for i = 1 to 10
        if isDisplay(sMASK,i) then
           lcl_checkForContactFields = lcl_checkForContactFields + 1
        else
           lcl_checkForContactFields = lcl_checkForContactFields
        end if
     next

    'If the value is greater than zero is there is at least one contact field that is displayed/required.
     if lcl_checkForContactFields > 0 then
        lcl_displayContactFields = True
     end if

	Else
 		 response.redirect("action.asp")
	End If

	Set oForm = Nothing 

 response.write "<p><a href=""action.asp""><strong><<< Return to E-Gov Request Form Entry List</strong></a></p>" & vbcrlf
 response.write "<form name=""frmRequestAction"" action=""action_cgi.asp?list=true"" method=""post"">" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""actionid"" value=""" & iFormID & """ />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""actiontitle"" value=""" & sTitle & """ />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""validstreet"" value="""" />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""control_field"" id=""control_field"" value="""" size=""20"" maxlength=""4001"" />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""showContactInfo"" id=""showContactInfo"" value=""" & lcl_displayContactFields & """ size=""5"" maxlength=""10"" />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""showContactEmail"" id=""showContactEmail"" value=""" & isDisplay(sMASK,4) & """ size=""5"" maxlength=""10"" />" & vbcrlf
 response.write "  <input type=""" & lcl_hidden & """ name=""showContactPhone"" id=""showContactPhone"" value=""" & isDisplay(sMASK,5) & """ size=""5"" maxlength=""10"" />" & vbcrlf

 response.write "<div style=""margin-top:20px; margin-left:20px;"">" & vbcrlf
 response.write "<font class=""formtitle"">" & sTitle & "</font>" & vbcrlf

'BEGIN: Emergency Note --------------------------------------------------------
  if blnEmergencyNote then
     response.write "<div class=""warning"">" & sEmergencyText & "</div>" & vbcrlf
  end if
'END: Emergency Note ----------------------------------------------------------

 response.write "<div class=""group"">" & vbcrlf
 response.write "  <div class=""orgadminboxf"">" & vbcrlf
 response.write "<p>" & vbcrlf

'BEGIN: Intro Information -----------------------------------------------------
  if sIntroText <> "" then
			  response.write sIntroText & vbcrlf
  else
  			response.write " - <i> Introduction text is currently blank </i> -" & vbcrlf
  end if
'END: Intro Information -------------------------------------------------------

 response.write "</p>" & vbcrlf

'BEGIN: Contact Information ---------------------------------------------------
 if lcl_displayContactFields then
    response.write "<p><strong><u>Contact Information:</u></strong></p>" & vbcrlf
    response.write "<table cellspacing=""0"" border=""0"" cellpadding=""0"" class=""tablenewaction"">" & vbcrlf
                      DrawContactTable(sMask)
    response.write "</table>" & vbcrlf
    response.write "</p>" & vbcrlf
 end if
'END: Contact Information -----------------------------------------------------

'BEGIN: Add Problem/Issue Location --------------------------------------------
 if lcl_orghasfeature_issue_location then
    if blnIssueDisplay then
       response.write "<p>" & vbcrlf
       response.write "<strong><u>" & sIssueName & "</u></strong>" & vbcrlf
       response.write "</p>" & vbcrlf
       response.write "<p>" & sIssueDesc & "</p>" & vbcrlf

       'subDisplayIssueLocation sIssueMask, iStreetNumberInputType, iStreetAddressInputType, sIssueQues
       subDisplayIssueLocation sIssueMask, sIssueQues, sHideIssueLocAddInfo

       response.write "</p>" & vbcrlf
    end if
 end if
'END: Problem/Issue Location --------------------------------------------------

'BEGIN: Form Field Information ------------------------------------------------
 response.write "<p>" & vbcrlf

 subDisplayQuestions iFormID,sMask,blnIssueDisplay

 response.write "</p>" & vbcrlf
'END: Form Field Information --------------------------------------------------

 response.write "</form>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function DrawContactTable(sMASK)

  if IsDisplay(sMASK,1) then
    	if IsRequired(sMASK,1) <> "" then
       	response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtFirst_Name-text/req"" value=""First Name"">" & vbcrlf
    	end if
%>
  <tr>
      <td align="right">
        		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,1)%></span>
        		First Name:
        		</span>
     	</td>
      <td>
        		<span class="cot-text-emphasized" title="This field is required"> 
         	<input type="text" value="" name="cot_txtFirst_Name" id="txtFirst_Name" style="width:300;" maxlength="50">
         	</span>
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,2) then
    	if IsRequired(sMASK,2) <> "" then
       	response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtLast_Name-text/req"" value=""Last Name"">" & vbcrlf
     end if
%>
  <tr>
      <td align="right">
        		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,2)%></span>
        		Last Name:
        		</span>
      </td>
      <td>
        		<span class="cot-text-emphasized" title="This field is required">
        		<input type="text" value="" name="cot_txtLast_Name" id="txtLast_Name" style="width:300;" maxlength="50">
        		</span>
      </td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,3) then
     if IsRequired(sMASK,3) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtBusiness_Name-text/req"" value=""Business Name"">" & vbcrlf
     end if
%>
  <tr>
      <td align="right"><%=IsRequired(sMASK,3)%>
          Business Name:
      </td>
      <td>
        		<input type="text" value="" name="cot_txtBusiness_Name" id="txtBusiness_Name" style="width:300;" maxlength="255">
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,4) then
     if IsRequired(sMASK,4) <> "" then
        response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtEmail-text/req"" value=""Email Address"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,4)%>Email:
      </td>
      <td>
        		<input type="text" value="" name="cot_txtEmail" id="txtEmail" style="width:300;" maxlength="512">
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,5) then
     if IsRequired(sMASK,5) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtDaytime_Phone-text/req"" value=""Daytime Phone"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,5)%>Daytime Phone:
     	</td>
      <td>
        		<input type="text" value="" name="cot_txtDaytime_Phone" id="txtDaytime_Phone" style="width:300;" maxlength="50">
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,6) then
     if IsRequired(sMASK,6) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtFax/req"" value=""Fax"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,6)%>Fax:
     	</td>
      <td>
        		<input type="text" value="" name="cot_txtFax" id="txtFax" style="width:300;" maxlength="50">
      </td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,7) then
     if IsRequired(sMASK,7) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtStreet/req"" value=""Street"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,7)%>Street:
     	</td>
      <td>
        		<input type="text" value="" name="cot_txtStreet" id="txtStreet" style="width:300;" maxlength="255">
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,8) then
     if IsRequired(sMASK,8) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtCity/req"" value=""City"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,8)%>City:
     	</td>
      <td>
        		<input type="text" value="" name="cot_txtCity" id="txtCity" style="width:300;" maxlength="50">
      </td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,9) then
     if IsRequired(sMASK,9) <> "" then
      		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtState_vSlash_Province/req"" value=""State or Province"">"
     end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,9)%>State / Province:
     	</td>
      <td>
        		<input type="text" value="" name="cot_txtState_vSlash_Province" id="txtState_vSlash_Province" size="5" maxlength="50">
     	</td>
  </tr>
<%
  end if

  if IsDisplay(sMASK,10) then
      if IsRequired(sMASK,10) <> "" then
       		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtZIP_vSlash_Postal_Code/req"" value=""Zipcode"">"
      end if
%>
  <tr>
      <td align="right">
        		<%=IsRequired(sMASK,10)%>ZIP / Postal Code:
     	</td>
      <td>
          <input type="text" value="" name="cot_txtZIP_vSlash_Postal_Code" id="txtZIP_vSlash_Postal_Code" style="width:300;" maxlength="50">
     	</td>
  </tr>
<%
  end if

end function

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
sub subDisplayQuestions(iFormID,sMask,blnIssueDisplay)
	Dim sAnswerList

	sSQL = "SELECT * FROM egov_action_form_questions "
	sSQL = sSQL & " WHERE formid = " & iFormID
	sSQL = sSQL & " AND (isinternalonly <> 1 OR isinternalonly IS NULL) "
	'sSQL = sSQL & " AND orgid = " & iorgid
	sSQL = sSQL & " ORDER BY sequence"

	'response.write sSQL & "<br /><br />"

	set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	if not oQuestions.eof then

		  response.write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""tablenewaction"">"
	
    do while not oQuestions.eof
		iQuestionCount = iQuestionCount + 1
		sAnswerList = ""

		'Determine if required
		sIsrequired = oQuestions("isrequired")

		if sIsrequired = True then
			sIsrequired = " <font color=""red"">*</font> "
		else
			sIsrequired = ""
		end If
		
		'sAnswerList = Replace(oQuestions("answerlist"),Chr(34),"&quot;")
		lcl_answerlist = oQuestions("answerlist")

        if lcl_answerlist <> "" then
           lcl_answerlist = Replace(lcl_answerlist,Chr(34),"&quot;")
        end if

		'Tracking current form configuration for editing later
		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("fieldtype")   & """ name=""fieldtype"" />" & vbcrlf
		response.write "<input type=""" & lcl_hidden & """ value=""" & lcl_answerlist & """ name=""answerlist"" />" & vbcrlf
		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("isrequired")  & """ name=""isrequired"" />" & vbcrlf
		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("sequence")    & """ name=""sequence"" />" & vbcrlf
		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("pdfformname") & """ name=""pdfformname"" />" & vbcrlf
		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("pushfieldid") & """ name=""pushfieldid"" />" & vbcrlf

		select case oQuestions("fieldtype")

			case "2"
				'Build Radio Question
				if sIsrequired <> "" then
					response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."" />" & vbcrlf
				end if

				response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """ />" & vbcrlf
				response.write "<tr><td class=""question"">" & sIsrequired & oQuestions("prompt")& "</td></tr>" & vbcrlf

				if oQuestions("answerlist") <> "" then
					arrAnswers = split(oQuestions("answerlist"),chr(10))

					for alist = 0 to ubound(arrAnswers)
						response.write "<tr><td><input type=""radio"" name=""fmquestion" & iQuestionCount & """ value=""" & Replace(arrAnswers(alist),Chr(34),"&quot;") & """ class=""formradio"" />" & arrAnswers(alist) & "</td></tr>" & vbcrlf
					next
				end if

				response.write "<tr style=""display: none"">" & vbcrlf
				response.write "    <td><input type=""radio"" name=""fmquestion" & iQuestionCount & """ value=""default_novalue"" checked=""checked"" /></td>" & vbcrlf
				response.write "</tr>" & vbcrlf
				response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

       		case "4"
				'Build Select Question
				if sIsrequired <> "" then
					response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."" />" & vbcrlf
				end if

				response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """ />" & vbcrlf
				response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf

				if oQuestions("answerlist") <> "" then
					arrAnswers = split(oQuestions("answerlist"),chr(10))

					response.write "<tr>" & vbcrlf
					response.write "    <td><select class=""formselect"" name=""fmquestion" & iQuestionCount & """>" & vbcrlf

					for alist = 0 to ubound(arrAnswers)
						response.write "<option value=""" & formatSelectOptionValue(arrAnswers(alist)) & """>" & arrAnswers(alist) & "</option>" & vbcrlf
					next

					response.write "        </select>" & vbcrlf
				end if

				response.write "    </td>" & vbcrlf
				response.write "</tr>" & vbcrlf
				response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

       			case "6"
					'Build Checkbox Question
					if sIsrequired <> "" then
						response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."" />" & vbcrlf
					end if

					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """ />" & vbcrlf
					response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf

					if oQuestions("answerlist") <> "" then
						arrAnswers = split(oQuestions("answerlist"),chr(10))

						i = 0
						for alist = 0 to ubound(arrAnswers)
							i = i + 1
							response.write "<tr><td><input type=""checkbox"" name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """ value=""" & Replace(arrAnswers(alist),Chr(34),"&quot;") & """ class=""formcheckbox"" onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"">" & arrAnswers(alist) & "</td></tr>"
						next
						i = i + 1
					end if

					response.write "<tr style=""display: none"">" & vbcrlf
					response.write "    <td>" & vbcrlf
					response.write "        <input type=""checkbox"" name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """ value=""default_novalue"" CHECKED onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"">" & vbcrlf
					response.write "        <span id=""total_options_fmquestion" & iQuestionCount & """>" & i & "</span>" & vbcrlf
					response.write "    </td>" & vbcrlf
					response.write "</tr>" & vbcrlf
					response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

       			case "8"
					'Build Text Question
					if sIsrequired <> "" then
						response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" & LEFT(oQuestions("prompt"),75) & "..."" />" & vbcrlf
					end if

					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & oQuestions("prompt") & """ />" & vbcrlf
					response.write "<tr><td class=""question"">" & sIsrequired  &  oQuestions("prompt") & "</td></tr>"& vbcrlf
					if iFormID = "17890" and oQuestions("prompt") = "Start Date" then
						sFirstDate = FindNextRyeBusinessDay(DateAdd("d",7,Date()))
						sFirstDate = RIGHT("0" & Month(sFirstDate),2) & "/" & RIGHT("0" & Day(sFirstDate),2) & "/" & Year(sFirstDate)
						%><script>
						$(document).ready(function() {

    							// assuming the controls you want to attach the plugin to 
    							// have the "datepicker" class set
							$('#startdate').Zebra_DatePicker({
            							first_day_of_week: 0,
            							format: 'm/d/Y',
								show_clear_date: false,
            							show_select_today: false,
								disabled_dates: ['* * * 0,6'],
  								direction: ['<%=sFirstDate%>', false],
								onSelect: function () {
									var enddate = new Date($('#startdate').val());
									enddate.setDate(enddate.getDate() + 29);
									var yyyy = enddate.getFullYear().toString();
									var mm = (enddate.getMonth()+1).toString(); // getMonth() is zero-based
									var dd  = enddate.getDate().toString();
									$("#enddate").html(mm + "/" + dd + "/" + yyyy);
								}
								});
							
 							});
						</script><%
						response.write "<tr><td><input id=""startdate"" name=""fmquestion" & iQuestionCount & """ value=""" & sFirstDate & """ type=""text"" style=""width:90px;"" readonly maxlength=""" & lcl_text_field_length & """ /></td></tr>"& vbcrlf
						response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
						response.write "<tr><td class=""question"">End Date</td></tr>"& vbcrlf
						response.write "<tr><td><div id=""enddate"">" & DateAdd("d",29,sFirstDate) & "</div></td></tr>"& vbcrlf
					elseif iFormID = "17890" and oQuestions("prompt") = "Address" then
						%><script>
						$(document).ready(function() {
							$("#issuelocation").autocomplete({
							source: function( request, response ) {
								//alert( request.term );
								$("#residentaddressid").val("0");
								$("#addressinsystemmsg").html("We do not consider this to be a valid address.");
								jQuery("#addressinsystemmsg").toggleClass("addressfound", false);
								jQuery("#addressinsystemmsg").toggleClass("addressnotfound", true);
								jQuery.ajax({
									url: "http://apidev.egovlink.com/api/ActionForm/GetAddressList?callback=?",
									type: "GET",
									dataType: "jsonp",
									contentType: "application/json",
									data: {
										_OrgId : 153,
										_MatchString: request.term,
										_MaxRows: 15
									},
									success: function( data ) {
										response( jQuery.map( data, function( item ) {
											$('#ui-id-1').css('display', 'block');
											return {
												label: item.streetaddress,
												value: item.streetaddress, 
												residentadressid: item.residentadressid
											}
										}));
									}
								});
							},
							minLength: 2,
							select: function( event, ui ) {
								$("#residentaddressid").val( ui.item ? ui.item.residentadressid : "0" );
								$("#addressinsystemmsg").html("We consider this to be a valid address.");
								jQuery("#addressinsystemmsg").toggleClass("addressfound", true);
								jQuery("#addressinsystemmsg").toggleClass("addressnotfound", false);
							},
							close: function( event, ui ) {
								//alert($("#issuelocation").val());
								$("#residentaddressval").val($("#issuelocation").val());
							},
							change: function( event, ui ) {
								if ($("#residentaddressval").val() != $("#issuelocation").val())
								{
									$("#residentaddressid").val("0");
									$("#addressinsystemmsg").html("We do not consider this to be a valid address.");
									jQuery("#addressinsystemmsg").toggleClass("addressfound", false);
									jQuery("#addressinsystemmsg").toggleClass("addressnotfound", true);
								}
							}
						});
						});
						</script><%
						response.write "<tr><td>"
							response.write "<input id=""issuelocation"" name=""fmquestion" & iQuestionCount & """ value="""" type=""text"" style=""width:300px;"" maxlength=""" & lcl_text_field_length & """ />"
							response.write "<div id=""addressinsystemmsg"">Start typing an address in the box above and then select one from the popup list if there is a match.</div>"
							response.write "<input type=""" & lcl_hidden & """ name=""ef:residentaddressid-text/nonzero-req"" value=""Address"">" & vbcrlf
							response.write "<input type=""hidden"" id=""residentaddressid"" value="""" />"
							response.write "<input type=""hidden"" id=""residentaddressval"" value="""" />"
						response.write "</td></tr>"& vbcrlf
						
					else
						response.write "<tr><td><input name=""fmquestion" & iQuestionCount & """ value="""" type=""text"" style=""width:300px;"" maxlength=""" & lcl_text_field_length & """ /></td></tr>"& vbcrlf
					end if
					response.write "<tr><td>&nbsp;</td></tr>"& vbcrlf

        		case "10"
					'build TextArea Question
					if sIsrequired <> "" then
						response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" & LEFT(oQuestions("prompt"),75) & "..."">" & vbcrlf
					end if

					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & oQuestions("prompt") & """>"
					response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
					response.write "<tr>" & vbcrlf
					response.write "    <td>" & vbcrlf
					response.write "        <textarea name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & """ class=""formtextarea"" maxlength=""" & lcl_textarea_field_length & """ onchange=""checkMaxLength();""></textarea><br />" & vbcrlf
					response.write "    </td>" & vbcrlf
					response.write "</tr>"
					response.write "<tr><td>&nbsp;</td></tr>"

       			case else

   			end select

		    oQuestions.movenext
    loop


  		response.write "</table>" & vbcrlf

 end if



'Contact Options
 response.write "<div>" & vbcrlf
	response.write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""tablenewaction"">" & vbcrlf
	response.write "  <tr><td class=question>Citizen's preferred communication method:</td></tr>" & vbcrlf
	response.write "  <tr><td>" & vbcrlf
	response.write "          <select name=""selContactMethod"">" & vbcrlf
				                         Call subListContactMethods(iSelected)
	response.write "          </select>" & vbcrlf
 response.write "      </td></tr>" & vbcrlf
	response.write "</table>" & vbcrlf
 response.write "</div>" & vbcrlf

'Build the OnClick for the Submit Button
'Determine if the form is diplaying the issue/problem location section
 if lcl_orghasfeature_large_address_list then
 		 if lcl_orghasfeature_issue_location AND blnIssueDisplay then
   		  lcl_onclick = "checkAddress( 'FinalCheck', 'no' );"
	   else
       lcl_onclick = "document.frmRequestAction.submit();"
 		 end if
	else
    if lcl_orghasfeature_issue_location AND blnIssueDisplay then
 	     lcl_onclick = "if(document.frmRequestAction.skip_address.value=='0000'&&document.frmRequestAction.ques_issue2.value==''){"
      	lcl_onclick = lcl_onclick &   "alert('Required field missing: Address');"
       lcl_onclick = lcl_onclick &   "document.frmRequestAction.skip_address.focus();"
		     lcl_onclick = lcl_onclick & "}else{"                                 
  		   lcl_onclick = lcl_onclick &   "document.frmRequestAction.submit();"
    		 lcl_onclick = lcl_onclick & "}"
		  else
  		   lcl_onclick = "document.frmRequestAction.submit();"
    end if
	end if

'Check to see if the org has the feature to display the "Do Not Send" options.
'If "yes" then check to see if session("userid") is assigned to current form.
'If "yes" then offer options to not send notification email(s).
	if lcl_orghasfeature_actionline_donotsend_submissionemails AND isUserAssignedToRequest(iFormID) = "Y" then
	   response.write "<p>" & vbcrlf
    response.write "<input type=""checkbox"" name=""doNotSendSelfEmail"" id=""doNotSendSelfEmail"" value=""on"" />" & vbcrlf
 		 response.write "Do NOT send me email notification that this request was submitted.<br />" & vbcrlf
	   response.write "<input type=""checkbox"" name=""doNotSendAllEmail"" id=""doNotSendAllEmail"" value=""on"" onclick=""disableSelfSendChkbox()"" />" & vbcrlf
    response.write "Do NOT send anyone email notifications that this request was submitted." & vbcrlf
 		 response.write "</p>" & vbcrlf
	else
    response.write "<input type=""hidden"" name=""doNotSendSelfEmail"" id=""doNotSendSelfEmail"" value="""" />" & vbcrlf
 		 response.write "<input type=""hidden"" name=""doNotSendAllEmail"" id=""doNotSendAllEmail"" value="""" />" & vbcrlf
	end if

'Submit Button
	response.write "<p>" & vbcrlf
 response.write "<div style=""text-align:left;"">" & vbcrlf
 response.write "<input type=""button"" style=""width:350px;"" class=""actionbtn"" name=""btnSubmit"" value=""CREATE ACTION REQUEST"" onclick=""" & lcl_onclick & """ />" & vbcrlf
	response.write "</div>" & vbcrlf
 response.write "</p>" & vbcrlf

	set oQuestions = nothing
	'Set oForms     = Nothing

end sub

'------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
  		sReturnValue = " <font color=""red"">*</font> "
	Else
		  sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'------------------------------------------------------------------------------
Function IsDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
  		sReturnValue = True
	Else
		  sReturnValue = False
	End If

	IsDisplay = sReturnValue
End Function

'------------------------------------------------------------------------------
Sub subListContactMethods(iSelected)

	sSQL = "SELECT * FROM egov_contactmethods ORDER BY contactdescription"

	Set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oMethods.EOF Then
	
		Do While NOT oMethods.EOF 
			response.write "<option value=""" &  oMethods("rowid") & """>" & oMethods("contactdescription") & "</option>"
			oMethods.MoveNext
		Loop

	End If
	oMethods.close
	Set oMethods = Nothing 

End Sub

'------------------------------------------------------------------------------
'sub subDisplayIssueLocation( sIssueMask, iStreetNumberInputType, iStreetAddressInputType, sIssueQues )
sub subDisplayIssueLocation( sIssueMask, sIssueQues, sHideIssueLocAddInfo )
	if sIssueMask = "" or IsNull(sIssueMask) then
  		sIssueMask = "121111"
	end if

 if trim(sIssueQues) = "" OR IsNull(trim(sIssueQues)) then
		  sIssueQues = "Provide any additional information on problem location in the box below."
	end if
	
	response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tablenewaction"">" & vbcrlf

	if IsDisplay(sIssueMask,2) then
    response.write "  <tr valign=""top"">" & vbcrlf
  		response.write "      <td align=""right"" valign=""top"">" & IsRequired(sIssueMask,2) & "Address:&nbsp;</td>" & vbcrlf
		  response.write "      <td>" & vbcrlf

  		'fnDrawInputType iStreetAddressInputType, "1"
    'fnDrawInputType session("orgid")
		response.write "<!--TWF: " & timer - iPageLogStartSecs & "-->" & vbcrlf
    if lcl_orghasfeature_large_address_list then
 						DisplayLargeAddressList session("orgid"), "R"
    else
       DisplayAddress session("orgid"), "R", 1
    end if
		response.write "<!--TWF: " & timer - iPageLogStartSecs & "-->" & vbcrlf

    response.write "          <br /> - Or Other Not Listed - <br /> " & vbcrlf
    response.write "          <input name=""ques_issue2"" id=""ques_issue2"" type=""text"" size=""60"" maxlength=""75"" />" & vbcrlf
		  response.write "      </td>" & vbcrlf
  		response.write "  </tr>" & vbcrlf

   'Unit
    response.write "  <tr>" & vbcrlf
  		response.write "      <td align=""right"">Unit:&nbsp;</td>" & vbcrlf
		  response.write "      <td><input type=""text"" name=""streetunit"" size=""8"" maxlength=""10""></td>" & vbcrlf
  		response.write "  </tr>" & vbcrlf
 end if

	if IsDisplay(sIssueMask,6) and not sHideIssueLocAddInfo then
  		sZipRequired = IsRequired(sIssueMask,6)
	  	response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td>&nbsp;</td>" & vbcrlf
		  response.write "      <td>" & IsRequired(sIssueMask,6) & sIssueQues & "<br />" & vbcrlf
		  response.write "          <textarea name=""ques_issue6"" class=""formtextarea"" maxlength=""512""></textarea>" & vbcrlf
		  response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

		 'IF REQUIRED ADD JAVASCRIPT CHECK
		  if sZipRequired <> "" then
		    	response.write "<input type=""hidden"" name=""ef:ques_issue6-textarea/req"" value=""Issue\problem additional information field"" />" & vbcrlf
		  end if
	end if

	response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
'sub fnDrawInputType(p_orgid)
'    if lcl_orghasfeature_large_address_list then
' 						DisplayLargeAddressList p_orgid, "R"
'    else
'       DisplayAddress p_orgid, "R", 1
'    end if

'    response.write "<br /> - Or Other Not Listed - <br /> " & vbcrlf
'    response.write "<input name=""ques_issue2"" id=""ques_issue2"" type=""text"" size=""60"" maxlength=""75"" />" & vbcrlf
'end sub

'------------------------------------------------------------------------------
sub DisplayLargeAddressList( p_orgid, sResidenttype )
	Dim sSql, oAddressList, sCompareName

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
 'sSQL = sSQL & " AND residenttype = '" & sResidenttype & "' "
 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " ORDER BY sortstreetname "
	
	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 0, 1

	if NOT oAddressList.eof then
  		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" onchange=""clearMsg('residentstreetnumber')"" /> &nbsp; " & vbcrlf
		  response.write "<select name=""skip_address"" id=""skip_address"" onchange=""clearMsg('skip_address')"">" & vbcrlf
  		response.write "  <option value=""0000"">Choose street from dropdown...</option>" & vbcrlf

  		do while NOT oAddressList.eof
      'Build the full street address
       sCompareName = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

      'Determine if the address is selected
    			if sStreetName = sCompareName then
      				lcl_selected_address = " selected=""selected"""
      				bFound = True
       else
          lcl_selected_address = ""
    			end if

    			response.write "<option value=""" & sCompareName & """" & lcl_selected_address & ">" & sCompareName & "</option>" & vbcrlf

    			oAddressList.MoveNext
    loop

  		response.write "</select>&nbsp;" & vbcrlf
    response.write "<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
 end if

	oAddressList.Close
	set oAddressList = nothing

end sub

'------------------------------------------------------------------------------
function DisplayAddress( iorgid, sResidenttype, blninputtype )
	Dim sSql, oAddressList

	sSQL = "SELECT residentaddressid, sortstreetname, isnull(residentstreetnumber,'') as residentstreetnumber, "
 sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iorgid
 'sSQL = sSQL & " AND residenttype = '" & sResidenttype & "' "
 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " ORDER BY sortstreetname, residentstreetnumber "
	
	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

	response.write "<select name=""skip_address"" id=""skip_address"" onchange=""clearMsg('skip_address')"">" & vbcrlf
	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

	do while not oAddressList.eof
   'Build the full street address
    lcl_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

  		response.write "  <option value="""  & oAddressList("residentaddressid")  & """>" & lcl_street_name & "</option>" & vbcrlf

  		oAddressList.MoveNext
	loop

	response.write "</select>" & vbcrlf

	oAddressList.close
	set oAddressList = nothing

end function

'------------------------------------------------------------------------------
function DisplayAddressNumber( iorgid, sResidenttype, blninputtype  )
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & iorgid
' sSQL = sSQL & " AND residenttype='" & sResidenttype & "' "
 sSQL = sSQL & " excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetnumber is not null "
 sSQL = sSQL & " ORDER BY residentstreetnumber "

	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN") , 3, 1

	If clng(blninputtype) = 0 Then
		sReturnValue =  "<select name=""ques_issue1"">" & vbcrlf
	Else
		sReturnValue =  "<select name=""skip_addressnumber"" onChange=""document.frmRequestAction.ques_issue1.value=document.frmRequestAction.skip_addressnumber[document.frmRequestAction.skip_addressnumber.selectedIndex].value;"">" & vbcrlf
	End If 

'	sReturnValue = sReturnValue &  "<option value="" "">*not on list</option>"
	sReturnValue = sReturnValue &  "<option value="" "">Choose street from dropdown</option>" & vbcrlf
		
	Do While NOT oAddressList.EOF 
		sReturnValue = sReturnValue & "<option value=""" &  oAddressList("residentstreetnumber") & """>" & oAddressList("residentstreetnumber") & "</option>" & vbcrlf
		oAddressList.MoveNext
	Loop

	sReturnValue = sReturnValue & "</select>" & vbcrlf

	oAddressList.close
	Set oAddressList = Nothing 

	DisplayAddressNumber = sReturnValue 

end function

'------------------------------------------------------------------------------
Public Function GetDefaultValues()
	
	sSQL = "SELECT * FROM organizations where orgid='" & session("orgid") & "'"
	Set oLocationValues = Server.CreateObject("ADODB.Recordset")
	oLocationValues.Open sSQL, Application("DSN") , 3, 1

	If NOT oLocationValues.EOF Then
		sDefaultCity  = oLocationValues("defaultcity")
		sDefaultState = oLocationValues("defaultstate")
		sDefaultZip   = oLocationValues("defaultzip")
	End If

	Set oLocationValues = Nothing

End Function

'------------------------------------------------------------------------------
function isUserAssignedToRequest(iFormID)
  lcl_return = "N"

  if iFormID <> "" then
   		sSQL = "SELECT count(action_form_id) as total_assigned "
 				sSQL = sSQL & " FROM egov_action_request_forms "
 				sSQL = sSQL & " WHERE orgid = " & session("orgid")
 				sSQL = sSQL & " AND action_form_id = " & iFormID
     sSQL = sSQL & " AND (assigned_userid = "  & session("userid")
     sSQL = sSQL & "  OR  assigned_userid2 = " & session("userid")
     sSQL = sSQL & "  OR  assigned_userid3 = " & session("userid") & ")"

 				set oExists = Server.CreateObject("ADODB.Recordset")
 				oExists.Open sSQL, Application("DSN"), 1, 3
		
     if oExists("total_assigned") > 0 then
        lcl_return = "Y"
     end if

     oExists.close
     set oExists = nothing

  end if

  isUserAssignedToRequest = lcl_return

end Function

'------------------------------------------------------------------------------
function formatSelectOptionValue(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,chr(10),"")
     lcl_return = replace(lcl_return,chr(13),"")
     lcl_return = replace(lcl_return,"<br>","")
     lcl_return = replace(lcl_return,"<br />","")
     lcl_return = replace(lcl_return,"<BR>","")
     lcl_return = replace(lcl_return,"<BR />","")
     lcl_return = replace(lcl_return,"""","&quot;")
  end if

  formatSelectOptionValue = lcl_return

end function



%>
