<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: edit_form.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module is the New Action Line Requests.
'
' MODIFICATION HISTORY
' ?.?	 10/16/06	 Steve Loar - Security, Header and Nav changed
' ?.?  04/23/07  Kankan Li - Allowed Days to Resolve, plus some other minor changes
' 3.0  03/28/08  David Boyer - Added the Public-PDF option (form letter style)
' 3.1  11/12/08  David Boyer - Modified Public-PDF option (merge form field style)
' 3.2  07/15/10  David Boyer - Added "notifications" section.
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../" ' Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

 if not UserHasPermission( Session("UserId"), "alerts" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_requestmergeforms = orghasfeature("requestmergeforms")
 lcl_orghasfeature_accepted_days     = orghasfeature("accepted days")

'Check for user permissions
 lcl_userhaspermission_create_requests = userhaspermission(session("userid"),"create requests")

 lcl_success = ""

'If UPDATE then process items
 if Request.ServerVariables("REQUEST_METHOD") = "POST" then
   	UserID  = request("assignUserID")
  		UserID2 = request("assignUserID2")
  		UserID3 = request("assignUserID3")
  		deptId  = request("deptId")

   	allowedunresolveddays = request("allowedunresolveddays")
    public_actionline_pdf = request("public_actionline_pdf")
	
   	if UserID = "" then	 
      	errorEmailMsg = "*** No user specified. Update not completed. ***"
     		blnUpdate     = True
   	else  
       lcl_assigned_email        = "NULL"
       lcl_assigned_userid2      = "NULL"
       lcl_assigned_userid3      = "NULL"
       lcl_allowedunresolveddays = ""
       lcl_public_actionline_pdf = "NULL"
       lcl_deptid                = deptid

       if userid2 <> "" then
          lcl_assigned_userid2 = userid2
       end if

       if userid3 <> "" then
          lcl_assigned_userid3 = userid3
       end if

       if allowedunresolveddays <> "" then
          lcl_allowedunresolveddays = allowedunresolveddays
       end if

       if public_actionline_pdf <> "" then
          lcl_public_actionline_pdf = "'" & public_actionline_pdf & "'"
       end if

       sSQL = "UPDATE egov_action_request_forms SET "
       sSQL = sSQL & " assigned_email = "   & lcl_assigned_email   & ", "
       sSQL = sSQL & " assigned_userid = "  & userid               & ", "
       sSQL = sSQL & " assigned_userid2 = " & lcl_assigned_userid2 & ", "
       sSQL = sSQL & " assigned_userid3 = " & lcl_assigned_userid3 & ", "

       if lcl_allowedunresolveddays <> "" then
          sSQL = sSQL & " allowedunresolveddays = " & lcl_allowedunresolveddays & ", "
       end if

       sSQL = sSQL & " deptid = "                & lcl_deptid & ", "
       sSQL = sSQL & " public_actionline_pdf = " & lcl_public_actionline_pdf
       sSQL = sSQL & " WHERE action_form_id = "  & request("form_id")

  					set oUpdate = Server.CreateObject("ADODB.Recordset")
		  			oUpdate.Open sSQL, Application("DSN"), 1, 3

   				set oUpdate = nothing

   				blnUpdate   = True
       lcl_success = "SU"
   				session("updatedAssignedUserID") = UserID	

       if request("form_id") <> "" then
          lcl_action_formid = request("form_id")
       else
          lcl_action_formid = isnull(request("control"),0)
       end if

   			'Update Category
   		 	'sSQLca = "SELECT form_category_id FROM egov_forms_to_categories where action_form_id=" & request("FORM_ID")

   				'set uFormCatUpdate = Server.CreateObject("ADODB.Recordset")
   				'uFormCatUpdate.CursorLocation = 3
   				'uFormCatUpdate.Open sSQLca, Application("DSN"), 1, 3
   		 	'uFormCatUpdate("form_category_id") = request("catId")
   				'uFormCatUpdate.Update
   				'set uFormCatUpdate = nothing		
		
       sSQL = "UPDATE egov_forms_to_categories "
       sSQL = sSQL & " SET form_category_id = " & request("catid")
       sSQL = sSQL & " WHERE action_form_id = " & request("form_id")

   				set oFormCatUpdate = Server.CreateObject("ADODB.Recordset")
   				oFormCatUpdate.Open sSQL, Application("DSN"),3,1

   			'BEGIN: Update Escalations ----------------------------------------------
       e = 0
       if request("totalEscalations") <> "" then
          lcl_totalEscalations = request("totalEscalations")
       else
          lcl_totalEscalations = 0
       end if

       if lcl_totalEscalations > 0 then
      				for e = 1 to lcl_totalEscalations

             sSQL = ""

            'Validate escalation_id
             if dbready_number(request("escalation_id_"&e)) then
                lcl_escalationid = request("escalation_id_"&e)
             else
                lcl_escalationid = 0
             end if

            'Remove any escalations marked to be removed
             lcl_remove_esc = request("escalationRemove_"&e)

             if lcl_remove_esc = "Y" then
            				sSQL = "DELETE FROM egov_action_escalations "
                sSQL = sSQL & " WHERE orgid = " & session("orgid")
                sSQL = sSQL & " AND escalation_id = " & lcl_escalationid

             else
               'Validate fields.
                if dbready_number(request("escTime"&e)) then
                   lcl_escTime = request("escTime"&e)
                else
                   lcl_escTime = "NULL"
                end if

                if request("escCriteria"&e) <> "" then
                   lcl_escCriteria = "'" & dbsafe(request("escCriteria"&e)) & "'"
                else
                   lcl_escCriteria = "NULL"
                end if

                if request("escNotify"&e) <> "0" then
                   lcl_escNotify = "'" & dbsafe(request("escNotify"&e)) & "'"
                else
                   lcl_escNotify = "NULL"
                end if

               'Create escalation
                if lcl_escNotify <> "NULL" AND lcl_escCriteria <> "NULL" AND lcl_escTime <> "NULL" then
                   if lcl_escalationid = 0 then
            		 							sSQL = "INSERT INTO egov_action_escalations ("
                 		   sSQL = sSQL & "orgId,"
   		                 sSQL = sSQL & "action_form_id,"
        	         	   sSQL = sSQL & "escNotify,"
   		                 sSQL = sSQL & "escCriteria,"
           		         sSQL = sSQL & "escTime"
            		        sSQL = sSQL & ") VALUES ("
                 		   sSQL = sSQL &  session("orgid")  & ", "
   		                 sSQL = sSQL &  lcl_action_formid & ", "
        	      	      sSQL = sSQL &  lcl_escNotify     & ", "
      		              sSQL = sSQL &  lcl_escCriteria   & ", "
           		         sSQL = sSQL &  lcl_escTime
      		              sSQL = sSQL & ")"
                  'Update the escalation
                   else
                      sSQL = "UPDATE egov_action_escalations SET "
                      sSQL = sSQL & " escNotify = "   & lcl_escNotify   & ", "
                      sSQL = sSQL & " escCriteria = " & lcl_escCriteria & ", "   
                      sSQL = sSQL & " escTime = "     & lcl_escTime   
                      sSQL = sSQL & " WHERE orgid = " & session("orgid")
                      sSQL = sSQL & " AND escalation_id = " & lcl_escalationid
         			   				end if
                end if
             end if

             if sSQL <> "" then
      			 						set oEscalationMaint = Server.CreateObject("ADODB.Recordset")
      		   					oEscalationMaint.Open sSQL, Application("DSN"),3,1
                set oEscalationMaint = nothing
             end if

      				next
       end if
      'END: Update Escalations ------------------------------------------------

      'BEGIN: Email Notifications ---------------------------------------------
       lcl_total_notifications = request("totalNotificationRows")
       i = 0

       if lcl_total_notifications > 0 then
          for i = 1 to lcl_total_notifications
             sSQL = ""

            'Validate notificationid
             if dbready_number(request("notificationid_"&i)) then
                lcl_notificationid = request("notificationid_"&i)
             else
                lcl_notificationid = 0
             end if

            'Remove any notifications marked to be removed
             lcl_remove = request("notificationRemove_"&i)

             if lcl_remove = "Y" then
                sSQL = "DELETE FROM egov_action_notifications "
                sSQL = sSQL & " WHERE action_form_id = " & lcl_action_formid
                sSQL = sSQL & " AND orgid = " & session("orgid")
                sSQL = sSQL & " AND notificationid = " & lcl_notificationid
             else
               'Validate fields.
                if dbready_number(request("notificationSendTo_"&i)) then
                   lcl_sendto = request("notificationSendTo_"&i)
                else
                   lcl_sendto = "NULL"
                end if

                if request("notificationSendAction_"&i) <> "" then
                   lcl_email_action = "'" & dbsafe(request("notificationSendAction_"&i)) & "'"
                else
                   lcl_email_action = "NULL"
                end if

               'Create notification
                if lcl_sendto <> "NULL" AND lcl_email_action <> "NULL" then
                   if lcl_notificationid = 0 then
                      sSQL = "INSERT INTO egov_action_notifications ("
                      sSQL = sSQL & "orgid, "
                      sSQL = sSQL & "action_form_id, "
                      sSQL = sSQL & "sendto, "
                      sSQL = sSQL & "email_action, "
                      sSQL = sSQL & "createdby, "
                      sSQL = sSQL & "created_date"
                      sSQL = sSQL & ") VALUES ("
                      sSQL = sSQL & session("orgid")  & ", "
                      sSQL = sSQL & lcl_action_formid & ", "
                      sSQL = sSQL & lcl_sendto        & ", "
                      sSQL = sSQL & lcl_email_action  & ", "
                      sSQL = sSQL & session("userid") & ", "
                      sSQL = sSQL & "'" & Now()       & "'"
                      sSQL = sSQL & ")"
                  'Update the notification
                   else
                      sSQL = "UPDATE egov_action_notifications SET "
                      sSQL = sSQL & " sendto = "       & lcl_sendto & ", "
                      sSQL = sSQL & " email_action = " & lcl_email_action   
                      sSQL = sSQL & " WHERE orgid = " & session("orgid")
                      sSQL = sSQL & " AND notificationid = " & lcl_notificationid
                   end if
                end if
             end if

             if sSQL <> "" then
                set oNotificationMaint = Server.CreateObject("ADODB.Recordset")
                oNotificationMaint.Open sSQL, Application("DSN"), 0, 1
                'oNotificationMaint.close
                set oNotificationMaint = nothing
             end if
          next
       end if
      'END: Email Notifications -----------------------------------------------

   	end if
 end if

'Check for a form id
	iID = request("control")
	if iID = "" then
		  iID = "0"
 end if

'Check for a screen message
 lcl_onload = ""
 lcl_msg    = ""

  if blnUpdate then
		   iID = request("form_id")

  			if errorEmailMsg <> "" then 
		     	lcl_msg = errorEmailMsg
  			else
		     	lcl_msg = setupScreenMsg(lcl_success)
     end if

    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"

  end if

  Dim OrganizationProperty_WeekDaysOrWeekends
  OrganizationProperty_WeekDaysOrWeekends = getWeekDaysOrWeekendsLabel(session("orgid"))

  lcl_notfound_msg = ""

		sSQL = "SELECT action_form_name, "
  sSQL = sSQL & " assigned_userid, "
  sSQL = sSQL & " assigned_userid2, "
  sSQL = sSQL & " assigned_userid3, "
  sSQL = sSQL & " deptid, "
  sSQL = sSQL & " allowedunresolveddays, "
  sSQL = sSQL & " public_actionline_pdf "
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = sSQL & " WHERE action_form_id = " & iID

		set oForm = Server.CreateObject("ADODB.Recordset")
		oForm.Open sSQL, Application("DSN"), 3, 1

		if not oForm.eof then
  			sFormName             = oForm("action_form_name")
		  	sUserID               = oForm("assigned_UserID")
  			sUserID2              = oForm("assigned_UserID2")
		  	sUserID3              = oForm("assigned_UserID3")
  			sDeptID               = oForm("deptId") 
		  	allowedunresolveddays = oForm("allowedunresolveddays")
     public_actionline_pdf = oForm("public_actionline_pdf")

  			if sUserID = "" or IsNull(sUserID) then
		  		  sUserID = session("updatedAssignedUserID")
  			end if
		else
		  	lcl_notfound_msg = "FORM NOT FOUND"
		end if

		set oForm = nothing
		
		sSQLc = "SELECT form_category_id "
  sSQLc = sSQLc & " FROM egov_forms_to_categories "
  sSQLc = sSQLc & " WHERE action_form_id = " & iID

		set oFormCat = Server.CreateObject("ADODB.Recordset")
		oFormCat.Open sSQLc, Application("DSN"), 3, 1
		
		if not oFormCat.eof then
  			sCatID = oFormCat("form_category_id")
		else
		  	sCatId = 0
		end if

		if request.form("addEsc")="true" then
 				addEsc = 1
		else
	 			addEsc = 0
		end if

  if sDeptID = "" then
     sDeptID = 0
  end if

  if sUserID <> "" then
     checkUserEmail sUserID
  end if

  if sUserID2 <> "" then
     checkUserEmail sUserID2
  end if

  if sUserID3 <> "" then
     checkUserEmail sUserID3
  end if

  lcl_required_field = "<font color=""#ff0000"">*</font>"
%>
<html>
<head>
  <title><%=langBSActionLine%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>
  <script src="../scripts/formvalidation_msgdisplay.js"></script>
  <script language="javascript">
<!-- 
		team = new Array(	
		<%		
		SQLdepts1 = "SELECT groupid, "
  SQLdepts1 = SQLdepts1 & " orgid, "
  SQLdepts1 = SQLdepts1 & " groupname, "
  SQLdepts1 = SQLdepts1 & " groupdescription "
  SQLdepts1 = SQLdepts1 & " FROM groups "
  SQLdepts1 = SQLdepts1 & " WHERE grouptype = 2 "
  SQLdepts1 = SQLdepts1 & " AND orgid = " & session("orgid")
  SQLdepts1 = SQLdepts1 & " AND isInactive <> 1 "
  SQLdepts1 = SQLdepts1 & " ORDER BY groupname"

		set oDepts1 = Server.CreateObject("ADODB.Recordset")
		oDepts1.Open SQLdepts1, Application("DSN"), 1, 3
		
		beginning = true
		
		while not oDepts1.EOF
					eSQL1 = "SELECT U.userID, "
     eSQL1 = eSQL1 & "email, "
     eSQL1 = eSQL1 & "FirstName, "
     eSQL1 = eSQL1 & "LastName, "
     eSQL1 = eSQL1 & "GroupID "
     eSQL1 = eSQL1 & "FROM Users U "
     eSQL1 = eSQL1 &     " LEFT OUTER JOIN UsersGroups G ON U.userID = G.userID "
     eSQL1 = eSQL1 & " WHERE orgid = " & session("orgId")
     eSQL1 = eSQL1 & " AND GroupID = " & oDepts1("groupid")

					set oUsers1 = Server.CreateObject("ADODB.Recordset")
					oUsers1.Open eSQL1, Application("DSN"), 1, 3

					if beginning = false then
						  		response.write ","	& chr(13)
					end if
						  
					if oUsers1.EOF then
 							response.write chr(13) & "null"	& chr(13)
					else
						  
						  response.write "new Array("	& chr(13)
							
        beginning2 = true
							 response.write "new Array("" "",0),"	&vbcrlf
										
 							while not oUsers1.EOF
   								if beginning2 = false then
			 						  		response.write ","	& chr(13)
				 				  end if

						  			response.write "new Array(""" & oUsers1("FirstName") & " " & oUsers1("LastName") & """, " & oUsers1("userID") & ")"	
   								beginning2 = false

    							oUsers1.MoveNext
				 			wend

 							response.write ")"

					end if

     set oUsers1 = nothing
					beginning = false			

   		oDepts1.MoveNext
  wend

  set oDepts1 = nothing
   %>
		);
		
		function fillSelectFromArray99(selectCtrl, itemArray, goodPrompt, badPrompt, defaultItem) 
		{
			var i, j;
			var prompt;
			// empty existing items
			for (i = selectCtrl.options.length; i >= 0; i--) {
				selectCtrl.options[i] = null; 
			}
			prompt = (itemArray != null) ? goodPrompt : badPrompt;
			if (prompt == null) {
				j = 0;
			}
			else {
				selectCtrl.options[0] = new Option(prompt);
				j = 1;
			}
			if (itemArray != null) {
				// add new items
				for (i = 0; i < itemArray.length; i++) {
					selectCtrl.options[j] = new Option(itemArray[i][0]);
					if (itemArray[i][1] != null) {
						selectCtrl.options[j].value = itemArray[i][1]; 
					}
					j++;
				}
				// select first item (prompt) for sub list
				selectCtrl.options[0].selected = true;
			 }
		}
		
		function fillSelectFromArray(selectCtrl, itemArray, goodPrompt, badPrompt, defaultItem) 
		{
			var i, j;
			var prompt;
			// empty existing items
			for (i = selectCtrl.options.length; i >= 0; i--) 
			{
				selectCtrl.options[i] = null; 
			}
			prompt = (itemArray != null) ? goodPrompt : badPrompt;
			if (prompt == null) 
			{
				j = 0;
			}
			else 
			{
				selectCtrl.options[0] = new Option(prompt);
				j = 1;
			}
			if (itemArray != null) 
			{
				// add new items
				for (i = 0; i < itemArray.length; i++) 
				{
					selectCtrl.options[j] = new Option(itemArray[i][0]);
					if (itemArray[i][1] != null) 
					{
						selectCtrl.options[j].value = itemArray[i][1]; 
					}
					j++;
				}
				// select first item (prompt) for sub list
				selectCtrl.options[0].selected = true;
			 }
		}
		
		function fillSelectFromArray2(selectCtrl, itemArray, goodPrompt, badPrompt, defaultItem) 
		{
			var i, j;
			var prompt;
			// empty existing items
			for (i = selectCtrl.options.length; i >= 0; i--) 
			{
				selectCtrl.options[i] = null; 
			}
			prompt = (itemArray != null) ? goodPrompt : badPrompt;
			if (prompt == null) 
			{
				j = 0;
			}
			else 
			{
				selectCtrl.options[0] = new Option(prompt);
				j = 1;
			}
			if (itemArray != null) 
			{
				// add new items
				for (i = 0; i < itemArray.length; i++) 
				{
					selectCtrl.options[j] = new Option(itemArray[i][0]);
					if (itemArray[i][1] != null) {
						selectCtrl.options[j].value = itemArray[i][1]; 
					}
					j++;
				}
				// select first item (prompt) for sub list
				selectCtrl.options[0].selected = true;
			 }
		}
		
		function fillSelectFromArray3(selectCtrl, itemArray, goodPrompt, badPrompt, defaultItem) 
		{
			var i, j;
			var prompt;
			// empty existing items
			for (i = selectCtrl.options.length; i >= 0; i--) 
			{
				selectCtrl.options[i] = null; 
			}
			prompt = (itemArray != null) ? goodPrompt : badPrompt;
			if (prompt == null) 
			{
				j = 0;
			}
			else {
				selectCtrl.options[0] = new Option(prompt);
				j = 1;
			}
			if (itemArray != null) 
			{
				// add new items
				for (i = 0; i < itemArray.length; i++) 
				{
					selectCtrl.options[j] = new Option(itemArray[i][0]);
					if (itemArray[i][1] != null) {
						selectCtrl.options[j].value = itemArray[i][1]; 
					}
					j++;
				}
				// select first item (prompt) for sub list
				selectCtrl.options[0].selected = true;
			 }
		}

		function ValidateForm() {
   var lcl_return_false = "N";
   var lcl_filename = document.getElementById("public_actionline_pdf").value;
   var lcl_file_ext = lcl_filename.substr(lcl_filename.length-4);

   if(lcl_filename!="") {
      if(lcl_file_ext.toUpperCase()!=".PDF") {
         document.getElementById("public_actionline_pdf").focus();
         inlineMsg(document.getElementById('addPDF').id,'<strong>Invalid Value: </strong>The file is not a valid PDF file.',10,'addPDF');
         lcl_return_false = "Y"
      }
   }

			if(document.getElementById("assignUserID").value == 0) {
				  //alert("Please select a person to be notified.");
  				//document.frmUpdate.catId.focus();
      document.getElementById("assignUserID").focus();
      inlineMsg(document.getElementById('assignUserID').id,'<strong>Required Field Missing: </strong>Assigned To',10,'assignUserID');
      lcl_return_false = "Y"
			}

			if(document.getElementById("catId").value == 0) {
  				//alert("Please select a category.");
				  //document.frmUpdate.catId.focus();
      document.getElementById("catId").focus();
      inlineMsg(document.getElementById('catId').id,'<strong>Required Field Missing: </strong>Category',10,'catId');
      lcl_return_false = "Y"
			}

			if(document.getElementById("deptId").value == 0) {
  				//alert("Please select a department.");
				  //document.frmUpdate.deptId.focus();
      document.getElementById("deptId").focus();
      inlineMsg(document.getElementById('deptId').id,'<strong>Required Field Missing: </strong>Department',10,'deptId');
      lcl_return_false = "Y"
			}

   if(lcl_return_false=="Y") {
      return false;
   }else{
   			document.frmUpdate.submit();
   }
		}

		var isIE=document.all? 1:0

		function numCheck_NoPoint(e)
		{
			keyEntry = !isIE? e.which:event.keyCode;
			if( ((keyEntry >= '48') && (keyEntry <='57')) || (keyEntry==0) || (keyEntry==8) ) 
			{
				return true;
			}
			else
			{
				alert('Please enter numbers ONLY');
				return false;  
			}
		}

function addNotificationRow() {
  var mytbl            = document.getElementById('AddNotificationTBL');
  var totalrows_notify = Number(document.getElementById("totalNotificationRows").value);

  //Increase the total rows by one.  This is index for the new row.
  totalrows_notify = totalrows_notify+1;

  //Set up the new row.
  mytbl = document.getElementById('AddNotificationTBL').insertRow(totalrows_notify);

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg_notify   = "";
  var lcl_evenodd_notify = totalrows_notify/2;
      lcl_evenodd_notify = lcl_evenodd_notify.toString();

  if(lcl_evenodd_notify.indexOf('.') > 0) {
     lcl_rowbg_notify = "#eeeeee";
  }else{
     lcl_rowbg_notify = "#ffffff";
  }

  mytbl.style.background = lcl_rowbg_notify;

  //Build the cells for the new row.
  var a = mytbl.insertCell(0);  //Send Notification To
  var b = mytbl.insertCell(1);  //Send Action
  var c = mytbl.insertCell(2);  //Created Info
  var d = mytbl.insertCell(3);  //Remove Row (checkbox)

  //Build the cells in the new row.
  //Send Notification To
  var lcl_sendto = '<input type="hidden" name="notificationid_'+totalrows_notify+'" id="notificationid_'+totalrows_notify+'" value="0" size="10" maxlength="10" />';
  lcl_sendto += '<select name="notificationSendTo_'+totalrows_notify+'" id="notificationSendTo_'+totalrows_notify+'" onchange="clearMsg(\'notificationSendTo_'+totalrows_notify+'\')">';
  lcl_sendto += '<option value=""></option>';
  lcl_sendto +=  <% DrawAdminUsersNew_javascript "","Y" %>;
  lcl_sendto += '</select>';
  a.innerHTML = lcl_sendto;

  //Sent Action
  var lcl_send_action = '<select name="notificationSendAction_' +totalrows_notify+ '" id="notificationSendAction_' +totalrows_notify+ '" onchange="clearMsg(\'notificationSendAction_' +totalrows_notify+ '\')">';
  lcl_send_action += '<option value="request_updated">Updated</option>';
  lcl_send_action += '<option value="request_closed">set to Resolved/Dismissed status</option>';
  lcl_send_action += '</select>';
  b.innerHTML = lcl_send_action;

  //Created Info
  c.innerHTML='&nbsp;';

  //Remove Row (checkbox)
  d.align="center";
  d.innerHTML='<input type="checkbox" name="notificationRemove_'+totalrows_notify+'" value="Y" />';

  //update the total row count.
  document.getElementById("totalNotificationRows").value = totalrows_notify;
}

function addEscalationRow() {
  var mytbl_esc     = document.getElementById('AddEscalationTBL');
  var totalrows_esc = Number(document.getElementById("totalEscalations").value);

  //Increase the total rows by one.  This is index for the new row.
  totalrows_esc = totalrows_esc+1;

  //Set up the new row.
  mytbl_esc = document.getElementById('AddEscalationTBL').insertRow(totalrows_esc);

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg_esc   = "";
  var lcl_evenodd_esc = totalrows_esc/2;
      lcl_evenodd_esc = lcl_evenodd_esc.toString();

  if(lcl_evenodd_esc.indexOf('.') > 0) {
     lcl_rowbg_esc = "#eeeeee";
  }else{
     lcl_rowbg_esc = "#ffffff";
  }

  mytbl_esc.style.background = lcl_rowbg_esc;

  //Build the cells for the new row.
  var a = mytbl_esc.insertCell(0);  //Reminder/Escalation
  var b = mytbl_esc.insertCell(1);  //Remove Row (checkbox)

  //Build the cells in the new row.
  //Reminder/Escalation

  var lcl_escalation = '<span id="esc_every_'+totalrows_esc+'"></span><input type="hidden" name="escalation_id_'+totalrows_esc+'" id="escalation_id_'+totalrows_esc+'" value="0" size=""10"" maxlength=""10"" />';
  lcl_escalation += '<input type="text" name="escTime'+totalrows_esc+'" id="escTime'+totalrows_esc+'" size="2" value="" onkeypress="return numCheck_NoPoint(event)" /> <%=OrganizationProperty_WeekDaysOrWeekends%>';
  lcl_escalation += '&nbsp;';
  lcl_escalation += '<select name="escCriteria'+totalrows_esc+'" id="escCriteria'+totalrows_esc+'">';
  lcl_escalation += <% displayEscalationStatuses "", "Y" %>;
  lcl_escalation += '</select>';
  lcl_escalation += '&nbsp;';
  lcl_escalation += '<select name="escNotify'+totalrows_esc+'" id="escNotify'+totalrows_esc+'" onChange="checkEsc('+totalrows_esc+')">';
  lcl_escalation += <% displayAssignEmails "", "Y", true %>;
  lcl_escalation += '</select>';

  a.colSpan   = 3;
  a.innerHTML = lcl_escalation;

  //Remove Row (checkbox)
  b.align     = 'center';
  b.innerHTML = '<input type="checkbox" name="escalationRemove_'+totalrows_esc+'" id="escalationRemove_'+totalrows_esc+'" value="Y" />';

  //update the total row count.
  document.getElementById("totalEscalations").value = totalrows_esc;
}

function checkEsc(iRow)
{
	var selVal = document.getElementById("escNotify" + iRow).value;
	if (selVal == -1)
	{
		document.getElementById("esc_every_"+iRow).innerHTML = "Every&nbsp;";
	}
	else
	{
		document.getElementById("esc_every_"+iRow).innerHTML = "";
	}
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = '';
}

  function doPicker(iFieldID) {
    lcl_width  = 600;
    lcl_height = 400;
    lcl_left   = (screen.availWidth/2)-(lcl_width/2);
    lcl_top    = (screen.availHeight/2)-(lcl_height/2);

    eval('window.open("linkpicker/linkpicker.asp?fid=' + iFieldID + '", "_picker", "width=' + lcl_width + ',height=' + lcl_height + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top + '")');
  }
//-->
</script>

<style>
  body {
     background-color: #ffffff;
     margin:           0px 0px;
  }

  #screenMsg {
     font-size:   10pt;
     color:       #ff0000;
     font-weight: bold;
  }
</style>

</head>
<%
  'response.write "<body bgcolor=""#ffffff"" leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0"" onload=""" & lcl_onload & """>" & vbcrlf
  response.write "<body onload=""" & lcl_onload & """>" & vbcrlf

  ShowHeader sLevel
%>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""frmUpdate"" id=""frmUpdate"" action=""edit_form.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""addEsc"" id=""addEsc"" value=""false"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""form_id"" id=""form_id"" value=""" & iID & """ />" & vbcrlf

  if not lcl_orghasfeature_requestmergeforms then
     response.write "  <input type=""hidden"" name=""public_actionline_pdf"" id=""public_actionline_pdf"" value="""" size=""1"" maxlength="""" />" & vbcrlf
  end if

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Edit Action Line Request Form Alerts</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

  if lcl_notfound_msg <> "" then
     response.write lcl_notfound_msg
  end if

  response.write "          <p>" & vbcrlf
  response.write "          <div style=""font-size:10px; padding-bottom:5px;"">" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""history.back();"" />" & vbcrlf
  response.write "                      <input type=""button"" name=""updateButton"" id=""updateButton"" value=""Save Changes"" class=""button"" onclick=""clearScreenMsg();ValidateForm();"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td align=""right"">" & vbcrlf

  if lcl_userhaspermission_create_requests then
     response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Return to Form Creator"" class=""button"" onclick=""location.href='../admin/manage_form.asp?iformid=" & iID & "'"" /><br />" & vbcrlf
  end if

  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "          <table class=""tablelist"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
  response.write "          		<tr><th colspan=""2"" align=""left"">&nbsp;&nbsp;&nbsp;&nbsp;Form Name: " & sFormName & "</th></tr>" & vbcrlf
  response.write "          		<tr>" & vbcrlf
  response.write "                <td>Department</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                 			<select name=""deptId"" id=""deptId"" onchange=""clearMsg('deptId')"">" & vbcrlf
                                        displayDeptList sDeptID
  response.write "                 			</select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Category</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""catId"" id=""catId"" onchange=""clearMsg('catId')""> " & vbcrlf
                                        displayCategoryList sCatID
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td>Assigned To&nbsp;" & lcl_required_field & "</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""assignUserID"" id=""assignUserID"" onchange=""clearMsg('assignUserID')"">" & vbcrlf
                                        displayAssignEmails sUserID, "", false
  response.write "                  		</select>" & vbcrlf
  response.write "                    <br />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "          		    <td>Notification&nbsp;" & lcl_required_field & "</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""assignUserID2"" id=""assignUserID2"">" & vbcrlf
                                        displayAssignEmails sUserID2, "", false
  response.write "                    </select>" & vbcrlf
  response.write "                    <br />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "          		    <td>Notification&nbsp;" & lcl_required_field & "</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""assignUserID3"" id=""assignUserID3"">" & vbcrlf
                                        displayAssignEmails sUserID3, "", false
  response.write "                    </select>" & vbcrlf
  response.write "                    <br />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  if lcl_orghasfeature_accepted_days then
     response.write "  <tr>" & vbcrlf
     response.write "      <td style=""width:200px"">Accepted Days to Resolve</td>" & vbcrlf
     response.write "      <td><input name=""allowedunresolveddays"" id=""allowedunresolveddays"" value=""" & allowedunresolveddays & """ size=""3"" maxlength=""3"" onkeypress=""return numCheck_NoPoint(event)"" />&nbsp;" & OrganizationProperty_WeekDaysOrWeekends & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  if lcl_orghasfeature_requestmergeforms then
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td style=""width:200px"">Public-side PDF for<br />New Action Line Requests</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <input type=""text"" name=""public_actionline_pdf"" id=""public_actionline_pdf"" value=""" & public_actionline_pdf & """ size=""40"" maxlength=""1000"" onchange=""clearMsg('addPDF')"" />&nbsp;" & vbcrlf
     response.write "          <input type=""button"" name=""addPDF"" id=""addPDF"" value=""Select PDF"" class=""button"" onclick=""clearMsg('addPDF');doPicker('public_actionline_pdf');"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  response.write "  <tr>" & vbcrlf
  response.write "      <td align=""center"" colspan=""2"">" & vbcrlf
  response.write "          <font color=""#ff0000"">* = Users must have an email address to appear in assignment lists.</font>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'BEGIN: Reminders/Escalations ------------------------------------------------
  'response.write "  <tr><th colspan=""2"" align=""left"">&nbsp;&nbsp;&nbsp;&nbsp;Reminders/Escalations</th></tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""3"" align=""left"">" & vbcrlf
                            displayRemindersEscalations session("orgid"), iID
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Reminders/Escalations --------------------------------------------------

 'BEGIN: Notifications --------------------------------------------------------
  'response.write "  <tr><th colspan=""2"" align=""left"">&nbsp;&nbsp;&nbsp;&nbsp;Notifications</th></tr>" & vbcrlf

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""3"">" & vbcrlf
                            displayNotifications session("orgid"), iID
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Notifications ----------------------------------------------------------

  'if lcl_userhaspermission_create_requests then
  '   response.write "  <tr>" & vbcrlf
  '   response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
     'response.write "          &nbsp;<strong><a href=""../admin/manage_form.asp?iformid=" & iID & """>Return to E-Gov Form Creator</a></strong>" & vbcrlf
  '   response.write "          <input type=""button"" value=""Return to E-Gov Form Creator"" class=""button"" onclick=""location.href='../admin/manage_form.asp?iformid=" & iID & "'"" />" & vbcrlf
  '   response.write "      </td>" & vbcrlf
  '   response.write "  </tr>" & vbcrlf
  'end if

  response.write "</table>" & vbcrlf

  response.write "	  </div>" & vbcrlf

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "  	</div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDeptList(iDeptID)

 'List departments and set selected department
  SQLdepts = "SELECT groupid, "
  SQLdepts = SQLdepts & " orgid, "
  SQLdepts = SQLdepts & " groupname, "
  SQLdepts = SQLdepts & " groupdescription "
  SQLdepts = SQLdepts & " FROM groups "
  SQLdepts = SQLdepts & " WHERE  grouptype = 2 "
  SQLdepts = SQLdepts & " AND orgid = " & session("orgid")
  SQLdepts = SQLdepts & " AND isInactive <> 1 "
  SQLdepts = SQLdepts & " ORDER BY groupname "

  set oDepts = Server.CreateObject("ADODB.Recordset")
  oDepts.Open SQLdepts, Application("DSN"), 1, 3

  blnFoundSelected = False
  sDepartmentList  = ""
  iFirstDeptID     = ""

  if not oDepts.eof then
    	response.write "  <option value=""0"">Please Select...</option>" & vbcrlf

  	  while not oDepts.eof
		
      		if iFirstDeptID = "" then
        			iFirstDeptID = oDepts("groupid")
      		end if

      		if iDeptID = oDepts("groupid") then 
        			lcl_selected     = " selected=""selected"" " 
         		blnFoundSelected = True
      		else
        			lcl_selected     = ""
      		end if

      		response.write "  <option value=""" & oDepts("groupid") & """" & lcl_selected & ">" & oDepts("groupname") & "</option>" & vbcrlf

      		oDepts.movenext
     wend
  end if

  set oDepts = nothing

end sub

'------------------------------------------------------------------------------
sub displayCategoryList(iCatID)
  SQLcats = "SELECT * "
  SQLcats = SQLcats & " FROM egov_form_categories "
  SQLcats = SQLcats & " WHERE orgid = " & session("orgid")
  SQLcats = SQLcats & " ORDER BY form_category_name "

  set oCats = Server.CreateObject("ADODB.Recordset")
  oCats.Open SQLcats, Application("DSN"), 1, 3

 	if not oCats.eof then
   		response.write "<option value=""0"">Please Select...</option>" & vbcrlf

   		while not oCats.eof
     			if CLng(iCatID) = CLng(oCats("form_category_id")) then 
           lcl_selected = " selected=""selected"""
        else
           lcl_selected = ""
        end if

    				response.write "  <option value=""" & oCats("form_category_id") & """" & lcl_selected & ">" & oCats("form_category_name") & "</option>" & vbcrlf

     			oCats.movenext
		   wend
  end if

  oCats.close
  set oCats = nothing

end sub

'------------------------------------------------------------------------------
sub displayAssignEmails(iUserID, iIsJavascript, isEsc)

 'Is this being written to a javascript function?
  if iIsJavascript = "Y" then
     lcl_js_backslash = "\"
     lcl_js_singlequote  = "'"
  else
     lcl_js_backslash = ""
     lcl_js_singlequote  = ""
  end if

		eSQL = "SELECT userid, "
  eSQL = eSQL & " email, "
  eSQL = eSQL & " firstname, "
  eSQL = eSQL & " lastname, isdeleted "
		eSQL = eSQL & " FROM users "
		eSQL = eSQL & " WHERE orgid = " & session("orgid")
		eSQL = eSQL & " AND (IsRootAdmin is null or IsRootAdmin = 0) "
		eSQL = eSQL & " AND email <> '' "
  		eSQL = eSQL & " AND (isDeleted = 0 "
		if iUserID <> "" then eSQL = eSQL & " or userid = '" & iUserID & "'"
  		eSQL = eSQL & " ) "
		eSQL = eSQL & " ORDER BY isdeleted desc, lastname, firstname"

		set oUsers = Server.CreateObject("ADODB.Recordset")
		oUsers.Open eSQL, Application("DSN"), 1, 3
		
		if not oUsers.eof then

  			response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """0" & lcl_js_backslash & """>Please Select...</option>" & lcl_js_singlequote & vbcrlf
	if isEsc then
        	if iIsJavascript = "Y" then
           		response.write "+"
        	end if
 		if iUserID = -1 then
           		lcl_selected = " selected=" & lcl_js_backslash & """selected" & lcl_js_backslash & """"
        	else
           		lcl_selected = ""
       		end if
  			response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """-1" & lcl_js_backslash & """" & lcl_selected & ">Employee Assigned to Request</option>" & lcl_js_singlequote & vbcrlf
	end if

  			while not oUsers.eof

    				if iUserID = oUsers("userID") then
           lcl_selected = " selected=" & lcl_js_backslash & """selected" & lcl_js_backslash & """"
        else
           lcl_selected = ""
        end if

        if iIsJavascript = "Y" then
           response.write "+"
        end if

  						response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """" & oUsers("userID") & lcl_js_backslash & """" & lcl_selected & ">" 
						if oUsers("isdeleted") then response.write "["
						response.write replace(oUsers("FirstName"),"'",lcl_js_backslash & "'") & " " & replace(oUsers("LastName"),"'",lcl_js_backslash & "'") 
						if oUsers("isdeleted") then response.write "]"
						response.write "</option>" & lcl_js_singlequote & vbcrlf

    				oUsers.movenext
 	  	wend

     set oUsers = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub checkUserEmail(iUserID)

  sSQLu = "SELECT firstname, "
  sSQLu = sSQLu & " lastname "
  sSQLu = sSQLu & " FROM users "
  sSQLu = sSQLu & " WHERE userid = " & clng(iUserID)
  sSQLu = sSQLu & " AND email = '' "

  set oCheck = Server.CreateObject("ADODB.Recordset")
  oCheck.Open sSqlu, Application("DSN"), 1, 3

  if not oCheck.eof then
     response.write "<font size=""2"" color=""#f0000"">* "
     response.write oCheck("firstname") & " " & oCheck("lastname")
     response.write " is currently assigned to this form but does not have an e-mail.  "
     response.write "<a href=""../dirs/update_user.asp?userid=" & iUserID & """>Click here</a> to add an e-mail to "
     response.write oCheck("firstname") & " " & oCheck("lastname")
  end if

  set oCheck = nothing

end sub

'------------------------------------------------------------------------------
sub displayEscalationStatuses(iStatus, iIsJavascript)

  lcl_submitted  = ""
  lcl_unresolved = ""

 'Is this being written to a javascript function?
  if iIsJavascript = "Y" then
     lcl_js_backslash = "\"
     lcl_js_singlequote  = "'"
  else
     lcl_js_backslash = ""
     lcl_js_singlequote  = ""
  end if

 'Determine which status is selected
  if UCASE(iStatus) = "SUBMITTED" then
     lcl_submitted  = " selected=" & lcl_js_backslash & """selected" & lcl_js_backslash & """"
     lcl_unresolved = ""
  elseif UCASE(iStatus) = "UNRESOLVED" then
     lcl_submitted  = ""
     lcl_unresolved = " selected=" & lcl_js_backslash & """selected" & lcl_js_backslash & """"
  end if

  response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """" & lcl_js_backslash & """></option>" & lcl_js_singlequote & vbcrlf

  if iIsJavascript = "Y" then
     response.write "+"
  end if

  response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """SUBMITTED"  & lcl_js_backslash & """" & lcl_submitted  & ">SUBMITTED</option>" & lcl_js_singlequote & vbcrlf

  if iIsJavascript = "Y" then
     response.write "+"
  end if

  response.write "  " & lcl_js_singlequote & "<option value=" & lcl_js_backslash & """UNRESOLVED" & lcl_js_backslash & """" & lcl_unresolved & ">UNRESOLVED</option>" & lcl_js_singlequote & vbcrlf

end sub

'-- Copied from DrawAdminUsersNews (in includes/common.asp) -------------------
function DrawAdminUsersNew_javascript(sUserID,isEmailRequired)
	dim sSql, oUsers, selected

	sSQL = "SELECT userid, "
 sSQL = sSQL & " FirstName, "
 sSQL = sSQL & " LastName "
 sSQL = sSQL & " FROM Users "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND (IsRootAdmin IS NULL OR IsRootAdmin = 0) "

 if isEmailRequired = "Y" then
    sSQL = sSQL & " AND email IS NOT NULL "
    sSQL = sSQL & " AND email <> '' "
 end if

 sSQL = sSQL & " ORDER BY LastName, firstname "

	Set oAdminUsers = Server.CreateObject("ADODB.Recordset")
	oAdminUsers.Open sSQL, Application("DSN"), 1, 3

 i = 0
	while not oAdminUsers.eof
    i = i + 1
   	if suserid = oAdminUsers("userid") then
       selected = " selected=\""selected\"""
    else
       selected = ""
    end if

    if i > 1 then
       response.write "+"
    end if

  		response.write "'<option value=\""" & oAdminUsers("userid") & "\""" & selected & ">" & replace(oAdminUsers("FirstName"),"'","\'") & " " & replace(oAdminUsers("LastName"),"'","\'") & "</option>'" & vbcrlf
  		oAdminUsers.movenext
	wend

	oAdminUsers.close
	set oAdminUsers = nothing 

end function

'------------------------------------------------------------------------------
sub displayRemindersEscalations(iOrgID, iFormID)

  response.write "<fieldset id=""fieldset"" class=""fieldset"">" & vbcrlf
  response.write "  <legend>REMINDERS/ESCALATIONS&nbsp;</legend>" & vbcrlf
  response.write "  <div style=""margin-bottom:2px"">" & vbcrlf
  response.write "    <input type=""button"" name=""sAddEscalation"" id=""sAddEscalation"" value=""Add Reminder/Escalation"" class=""button"" onclick=""addEscalationRow();"" />" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <table id=""AddEscalationTBL"" border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""margin-top:10px"" width=""100%"">" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <th colspan=""3"">Reminder/Escalation</th>" & vbcrlf
  response.write "        <th align=""center"">Remove</th>" & vbcrlf
  response.write "    </tr>" & vbcrlf


		eSQLesc = "SELECT escalation_id, orgid, action_form_id, escTime, escCriteria, escNotify, NumEmailSent "
  eSQLesc = eSQLesc & " FROM egov_action_escalations "
  eSQLesc = eSQLesc & " WHERE orgid = " & iOrgID
  eSQLesc = eSQLesc & " AND action_form_id = " & iFormID
  eSQLesc = eSQLesc & " ORDER BY escTime asc "

		set oESC = Server.CreateObject("ADODB.Recordset")
		oESC.Open eSQLesc, Application("DSN"), 1, 3

		e = 0
  lcl_bgcolor = "#ffffff"

  if not oESC.eof then
   		while not oESC.eof

		      e = e + 1
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

        response.write "      <tr id=""addEscalationRow"&e&""" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "          <td colspan=""3"">" & vbcrlf
	response.write "	      <span id=""esc_every_" & e & """>" & vbcrlf
	if oESC("escNotify") = -1 then	response.write "Every&nbsp;"
	response.write "	      </span>" & vbcrlf
        response.write "              <input type=""hidden"" name=""escalation_id_"&e&""" id=""escalation_id_"&e&""" value=""" & oESC("escalation_id") & """ size=""10"" maxlength=""10"" />" & vbcrlf
        response.write "              <input type=""text"" name=""escTime"&e&""" id=""escTime"&e&""" size=""2"" value=""" & oESC("escTime") & """ onkeypress=""return numCheck_NoPoint(event)"" /> " & OrganizationProperty_WeekDaysOrWeekends & vbcrlf
 		   		response.write "              <select name=""escCriteria"&e&""" id=""escCriteria"&e&""">" & vbcrlf
                                        displayEscalationStatuses oESC("escCriteria"), ""
        response.write "              </select>" & vbcrlf
        response.write "              <select name=""escNotify"&e&""" id=""escNotify"&e&""" onChange=""checkEsc(" & e & ")"">" & vbcrlf
                                        displayAssignEmails oESC("escNotify"), "", true
        response.write "              </select>" & vbcrlf
        response.write "          </td>" & vbcrlf
        response.write "          <td align=""center"">" & vbcrlf
        response.write "              <input type=""checkbox"" name=""escalationRemove_"&e&""" id=""escalationRemove_"&e&""" value=""Y"" />" & vbcrlf
        response.write "          </td>" & vbcrlf
        response.write "      </tr>" & vbcrlf

        oESC.movenext
   		wend

     set oESC = nothing
  else
     'if addEsc = 0 then
     '   response.write "              <tr><td colspan=""3"" style=""text-align:center; color:#800000"">*** No reminders/escalations have been set ***</td></tr>" & vbcrlf
     'end if
  end if

'		if addEsc = 1 then
'   		e = e + 1

'     response.write "              <tr id=""addEscalationRow" & e & """ bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
'     response.write "                  <td colspan=""3"">" & vbcrlf
'     response.write "                      <input type=""text"" name=""escTime"&e&""" id=""escTime"&e&""" size=""2"" value="""" /> days" & vbcrlf
'     response.write "                      <select name=""escCriteria"&e&""" id=""escCriteria"&e&""">" & vbcrlf
'                                             displayEscalationStatuses "", ""
'     response.write "                      </select>" & vbcrlf
'     response.write "                      <select name=""escNotify"&e&""" id=""escNotify"&e&""">" & vbcrlf
'                                             displayAssignEmails "", "", true
'     response.write "                    		</select>" & vbcrlf
'     response.write "                  </td>" & vbcrlf
'     response.write "              </tr>" & vbcrlf
'  end if

  response.write "  </table>" & vbcrlf
  response.write "  <input type=""hidden"" name=""totalEscalations"" id=""totalEscalations"" value=""" & e & """ />" & vbcrlf
  response.write "</fieldset>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub displayNotifications(iOrgID, iFormID)
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>NOTIFICATIONS&nbsp;</legend>" & vbcrlf
  response.write "  <div style=""margin-bottom:2px"">" & vbcrlf
  response.write "    <input type=""button"" name=""sAddNotification"" id=""sAddNotification"" value=""Add Notification"" class=""button"" onclick=""addNotificationRow();"" />" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <table id=""AddNotificationTBL"" border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""margin-top:10px"" width=""100%"">" & vbcrlf
  response.write "    <tr align=""left"" id=""addNotificationRow_0"">" & vbcrlf
  response.write "        <th>Send Notification To</th>" & vbcrlf
  response.write "        <th>Send Notification<br />Whenever a request is...</th>" & vbcrlf
  response.write "        <th align=""center"">Created By</th>" & vbcrlf
  response.write "        <th align=""center"">Remove</th>" & vbcrlf
  response.write "    </tr>" & vbcrlf

  sSQLr = "SELECT n.notificationid, n.orgid, n.action_form_id, n.email_action, n.sendto, "
  sSQLr = sSQLr & " isnull(u.FirstName,'') AS SendToFirstName, isnull(u.LastName,'') AS SendToLastName, "
  sSQLr = sSQLr & " n.created_date, n.createdby, isnull(u2.FirstName,'') AS CreatedByFirstName, "
  sSQLr = sSQLr & " isnull(u2.LastName,'') AS CreatedByLastName, n.created_date "
  sSQLr = sSQLr & " FROM egov_action_notifications n "
  sSQLr = sSQLr &      " LEFT JOIN users u  ON n.sendto = u.userid "
  sSQLr = sSQLr &      " LEFT JOIN users u2 ON n.createdby = u2.userid "
  sSQLr = sSQLr & " WHERE n.orgid = " & iOrgID
  sSQLr = sSQLr & " AND n.action_form_id = " & iFormID
  sSQLr = sSQLr & " ORDER BY n.notificationid "

  set oNotifications = Server.CreateObject("ADODB.Recordset")
  oNotifications.Open sSQLr, Application("DSN"), 1, 3

  iRowCount   = 0
  lcl_bgcolor = "#ffffff"

  if not oNotifications.eof then
     while not oNotifications.eof
        iRowCount   = iRowCount + 1
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

        if oNotifications("email_action") = "request_closed" then
           lcl_selected_action_request_updated = ""
           lcl_selected_action_request_closed  = " selected=""selected"""
        else
           lcl_selected_action_request_updated = " selected=""selected"""
           lcl_selected_action_request_closed  = ""
        end if

        response.write "    <tr id=""addNotificationRow_" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <input type=""hidden"" name=""notificationid_" & iRowCount & """ id=""notificationid_" & iRowCount & """ value=""" & oNotifications("notificationid") & """ size=""10"" maxlength=""10"" />" & vbcrlf
        response.write "            <select name=""notificationSendTo_" & iRowCount & """ id=""notificationSendTo_" & iRowCount & """ onchange=""clearMsg('notificationSendTo_" & iRowCount & "')"">" & vbcrlf
                                      displayAssignEmails oNotifications("sendto"), "", false
        response.write "            </select>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <select name=""notificationSendAction_" & iRowCount & """ id=""notificationSendAction_" & iRowCount & """ onchange=""clearMsg('notificationSendAction_" & iRowCount & "')"">" & vbcrlf
        response.write "              <option value=""request_updated""" & lcl_selected_action_request_updated & ">Updated</option>" & vbcrlf
        response.write "              <option value=""request_closed"""  & lcl_selected_action_request_closed  & ">set to Resolved/Dismissed status</option>" & vbcrlf
        response.write "            </select>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td align=""center"">" & vbcrlf
        response.write              oNotifications("CreatedByFirstName") & " " & oNotifications("CreatedByLastName") & "<br />" & vbcrlf
        response.write              formatdatetime(oNotifications("created_date"),vbshortdate) & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td align=""center"">" & vbcrlf
        response.write "            <input type=""checkbox"" name=""notificationRemove_" & iRowCount & """ id=""notificationRemove_" & iRowCount & """ value=""Y"" />" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf

        oNotifications.movenext
     wend

  end if

  oNotifications.close
  set oNotifications = nothing

  response.write "  </table>" & vbcrlf
  response.write "  <input type=""hidden"" name=""totalNotificationRows"" id=""totalNotificationRows"" value=""" & iRowCount & """ />" & vbcrlf
  response.write "</fieldset>" & vbcrlf
end sub

'------------------------------------------------------------------------------
function getWeekDaysOrWeekendsLabel(iOrgID)

  lcl_return = "day(s)"

  if iOrgID <> "" then
     sSQL = "SELECT usesweekdays "
     sSQL = sSQL & " FROM Organizations "
     sSQL = sSQL & " WHERE orgid = " & iOrgID

     set oUsesWeekDays = Server.CreateObject("ADODB.Recordset")
     oUsesWeekDays.Open sSQL, Application("DSN"), 0, 1

     if not oUsesWeekDays.eof then
       	if oUsesWeekDays(0) then
         		lcl_return = "weekday(s)" 
        end if
     end if

     oUsesWeekDays.close
     set oUsesWeekDays = nothing

  end if

  getWeekDaysOrWeekendsLabel = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_msg = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_msg = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_msg = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
