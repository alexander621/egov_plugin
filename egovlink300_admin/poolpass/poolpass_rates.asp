<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="../membershipcards/membership_card_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: pool_pass_rates.asp
' AUTHOR: Steve Loar
' CREATED: 01/31/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/31/2006 Steve Loar - Code added
' 2.0  07/11/2006 Steve Loar - Membership related changes
' 2.1	 10/05/2006	Steve Loar - Header and nav changed
' 2.2  08/26/2008 David Boyer - Added Membership Renewals
' 2.3  01/27/2009 David Boyer - Added Card Layout dropdown list.
' 2.4  01/29/2009 David Boyer - Added "isPunchcard" and "Punchcard Limit" fields.
' 2.5  02/11/2009 David Boyer - Added "Non-Resident Limit"
' 2.6  03/02/2012 David Boyer - Added "terms"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iSeasonId, iSeasonYear, dStartDate, dEndDate, sResidentType, iMaxOrder, sDisplayText
 Dim iMembershipId, oMembership, iRowCount, x, sMembershipTypeName, sResidentTypeName

 sLevel = "../" ' Override of value from common.asp

 if not userhaspermission(session("userid"),"membership rates") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

 set oMembership = New classMembership

 if request("sMembershipType") <> "" then
    lcl_membership_type = request("sMembershipType")
 else
    'lcl_membership_type = "pool"
    lcl_membership_type = oMembership.GetFirstMembershipType()
 end if

'Set the membershipid the the type selected.  If this is the initial screen opening then default it to "pool".
 'oMembership.SetMembershipId( "pool" )
 oMembership.SetMembershipId(lcl_membership_type)

 if request("sResidentType") = "" then
   	sResidentType = "R"
 else 
   	sResidentType = request("sResidentType")
 end if

 iMaxOrder    = GetMaxDisplayOrder(sResidentType)
 lcl_scripts  = ""
 lcl_rowcount = 0
 lcl_rowspan  = "4"

'Set a style to control the widths of all of the tables
 lcl_style_tablewidth = " style=""width: 1000px"""

'Check for org features
 lcl_orghasfeature_membership_renewals              = orghasfeature("membership_renewals")
 lcl_orghasfeature_pool_attendance_view             = orghasfeature("pool_attendance_view")
 lcl_orghasfeature_card_layout_multiplelayouts      = orghasfeature("card_layout_multiplelayouts")
 lcl_orghasfeature_nonresidentcap                   = orghasfeature("nonresidentcap")
 lcl_orghasfeature_nonresident_preregistrationdates = orghasfeature("nonresident_preregistrationdates")

'Check for user permissions
 lcl_userhaspermission_membership_renewals              = userhaspermission(session("userid"),"membership_renewals")
 lcl_userhaspermission_pool_attendance_view             = userhaspermission(session("userid"),"pool_attendance_view")
 lcl_userhaspermission_nonresidentcap                   = orghasfeature("nonresidentcap")
 lcl_userhaspermission_nonresident_preregistrationdates = userhaspermission(session("userid"),"nonresident_preregistrationdates")

'Determine which screen message to display, if any.
 lcl_onload  = ""
 lcl_message = ""

 if request("success") = "SN" then
    lcl_message = "Successfully Created..."
 elseif request("success") = "SU" then
    lcl_message = "Successfully Updated..."
 elseif request("success") = "SD" then
    lcl_message = "Successfully Deleted..."
 end if

 if lcl_message <> "" then
    lcl_onload = "displayScreenMsg('" & lcl_message & "');"
 end if

'Set up the variables
 lcl_rateid                   = 0
 lcl_displayorder             = (iMaxOrder+1)
 lcl_description              = ""
 lcl_periodid                 = 0
 lcl_amount                   = 0
 lcl_maxsignups               = 0
 lcl_message                  = ""
 lcl_attendancetypeid         = ""
 lcl_publiccanpurchase        = True
 lcl_isRenewable              = False
 lcl_renewalstartdate         = ""
 lcl_isPunchcard              = False
 lcl_punchcardlimit           = 0
 lcl_renewaltimeafterexpire   = ""
 lcl_nonresidentlimit         = 0
 lcl_nonresident_prestartdate = ""
 lcl_nonresident_preenddate   = ""
 lcl_cardid                   = ""
 lcl_isEnabled                = 1
 lcl_checked_publicdisplay    = oMembership.ShowMembershipRatePublicDisplay(sResidentType)
%>
<html>
<head>
  <title>E-Gov Administration Console { Membership Rates }</title>

  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../poolpass.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

 <style type="text/css">
   #screenMsg {
      color:       #ff0000;
      font-size:   10pt;
      font-weight: bold;
   }

   #membershipTypeHeaderRow td {
      white-space: nowrap;
   }

   #membershipTerms {
      width:  400px;
      height: 200px;
   }

   .fieldset {
      border-radius: 5px;
   }

   .fieldset legend {
      font-size: 1em;
   }
 </style>

  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/isvaliddate.js"></script>
  <script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
  <!--
$(document).ready(function(){
  $('#membershipTerms').css('display','none');
  $('#saveTermsButton').prop('disabled',true);

  $('#editTermsButton').click(function() {
     $('#saveTermsButton').prop('disabled',false);
     $('#editTermsButton').prop('disabled',true);
     $('#membershipTerms').slideDown('slow');
  });

  $('#saveTermsButton').click(function() {
     $('#saveTermsButton').prop('disabled',true);
     $('#editTermsButton').prop('disabled',false);
     $('#membershipTerms').slideUp('slow');

     var lcl_membershipTerms = $('#membershipTerms').val();


     $.post('updateMembershipTerms.asp', {
        orgid:           '<%=session("orgid")%>',
        membershipid:    '<%=oMembership.MembershipId%>',
        membershipTerms: lcl_membershipTerms,
        action:          'EDIT_TERMS'
     }, function(result) {
        displayScreenMsg(result);
     });
  });

  $('#showTerms').click(function() {
     var lcl_showTerms = '';

     if(document.getElementById('showTerms').checked) {
        lcl_showTerms = 'on';
     }

     $.post('updateMembershipTerms.asp', {
        orgid:        '<%=session("orgid")%>',
        membershipid: '<%=oMembership.MembershipId%>',
        showTerms:    lcl_showTerms,
        action:       'SHOW_TERMS'
     }, function(result) {
        displayScreenMsg(result);
     });
  });

  $('#sMembershipType').change(function() {
     $('#membershiptypeform').submit();
  });

  $('#showMessage').click(function() {
     var lcl_showMessage = '';

     if(document.getElementById('showMessage').checked) {
        lcl_showMessage = 'on';
     }

     $.post('updateMembershipMessage.asp', {
        orgid:        '<%=session("orgid")%>',
        membershipid: '<%=oMembership.MembershipId%>',
        showMessage:  lcl_showMessage
     }, function(result) {
        displayScreenMsg(result);
     });
  });
});

function SaveRate(p_action) {
  if(p_action=="ADD") {
     lcl_form        = document.getElementById("rateform_add");
     lcl_total_rates = 0;
     lcl_i_start     = 0;
  }else{
    lcl_form        = document.getElementById("rateform_save");
    lcl_total_rates = document.getElementById("total_rates").value;
    lcl_i_start     = 1;
  }

  lcl_false_cnt   = 0;

  //---------------------------------------------------------------------------
  //Check the description for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         if(document.getElementById("description_"+i).value == "") {
        			 inlineMsg(document.getElementById("description_"+i).id,'<strong>Required Field Missing: </strong>Description',8,'description_'+i);
            lcl_false_cnt = lcl_false_cnt + 1;

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("description_"+i);
            }
         }else{
            clearMsg('description_'+i);
       		}
    }

    //If error messages exist then do not submit the form and return focus to the first field found in error.
      if(lcl_false_cnt > 0) {
         lcl_focus.focus();
         return false;
      }else{
         lcl_focus_cnt = 0;
      }

  //---------------------------------------------------------------------------
  //Check the "amount" and "format of the amount" for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         lcl_message    = "";
         var editAmount = /^\d+.?\d{0,2}$/;

     		//Check the dollar amount
   	     if(document.getElementById("amount_"+i).value == "")	{
            lcl_message = "<strong>Required Field Missing: </strong>Amount";
            lcl_false_cnt = lcl_false_cnt + 1;

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("amount_"+i);
            }
         }

      	//Validate the format of the amount
         if(document.getElementById("amount_"+i).value!="") {
           	var Ok = editAmount.test(document.getElementById("amount_"+i).value);

           	if(! Ok)	{
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The Amount must be in a valid money format";
               lcl_false_cnt = lcl_false_cnt + 1;

               if(lcl_false_cnt == 1) {
                  lcl_focus = document.getElementById("amount_"+i);
               }
            }
         }

         if(lcl_message!="") {
   	    	 	 inlineMsg(document.getElementById("amount_"+i).id,lcl_message,10,'amount_'+i);
         }else{
            clearMsg('amount_'+i)
      		 }
    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
    if(lcl_false_cnt > 0) {
       lcl_focus.focus();
       return false;
    }else{
       lcl_focus_cnt = 0;
    }

  //---------------------------------------------------------------------------
  //Check the "maxsignups" and "format of the maxsignups" for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         lcl_message        = "";
  	 	    var editMaxAllowed = /^\d*$/;

     		//Check the Max On Pass
   	     if(document.getElementById("maxsignups_"+i).value == "")	{
            lcl_message = "<strong>Required Field Missing: </strong>Max On Pass";
            lcl_false_cnt = lcl_false_cnt + 1;

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("maxsignups_"+i);
            }
         }

      	//Validate the format of the Max On Pass
         if(document.getElementById("maxsignups_"+i).value!="") {
           	var Ok = editMaxAllowed.test(document.getElementById("maxsignups_"+i).value);
           	if(! Ok)	{
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Max On Pass\" must be in a number format.";
               lcl_false_cnt = lcl_false_cnt + 1;

               if(lcl_false_cnt == 1) {
                  lcl_focus = document.getElementById("maxsignups_"+i);
               }
            }
         }

         if(lcl_message!="") {
    	    		 inlineMsg(document.getElementById("maxsignups_"+i).id,lcl_message,10,'maxsignups_'+i);
         }else{
            clearMsg('maxsignups_'+i);
       		}
    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
    if(lcl_false_cnt > 0) {
       lcl_focus.focus();
       return false;
    }else{
       lcl_focus_cnt = 0;
    }

  <% if lcl_orghasfeature_membership_renewals AND lcl_userhaspermission_membership_renewals then %>
  //---------------------------------------------------------------------------
  //Check the "Renewal Start Date" and "format of the Renewal Start Date" for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         lcl_message = "";

     		//If the isRewnewable checkbox is "checked" then check for a Renewal Start Date
   	     if(document.getElementById("isRenewable_"+i).checked && document.getElementById("renewalstartdate_"+i).value=="")	{

            lcl_message = "<strong>Required Field Missing: </strong>Renewal Start Date";
            lcl_false_cnt = lcl_false_cnt + 1;

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("renewalstartdate_"+i);
            }
         }

      	//Validate the format of the Renewal Start Date
         if(document.getElementById("isRenewable_"+i).checked && document.getElementById("renewalstartdate_"+i).value!="") {
           	var Ok = isValidDate(document.getElementById("renewalstartdate_"+i).value);
           	if(! Ok)	{
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Renewal Start Date\" must be in a date format.<br /><span style=\"color:#800000;\">(i.e. mm/dd/yyyy)</span>";
               lcl_false_cnt = lcl_false_cnt + 1;

               if(lcl_false_cnt == 1) {
                  lcl_focus = document.getElementById("renewalstartdate_"+i);
               }
            }
         }

         if(lcl_message!="") {
    	    		 inlineMsg(document.getElementById("renewalstartdate_"+i).id,lcl_message,10,'renewalstartdate_'+i);
         }else{
            clearMsg('renewalstartdate_'+i);
       		}
    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
    if(lcl_false_cnt > 0) {
       lcl_focus.focus();
       return false;
    }else{
       lcl_focus_cnt = 0;
    }

  //---------------------------------------------------------------------------
  //Check the "Days to Renew after Expiration Date" for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         lcl_message         = "";
  	 	    var editRenewalDays = /^\d*$/;

     		//Check the Days to Renew after Expiration Date
     		//If the isRewnewable checkbox is "checked" then check for a Days to Renew After Expiration Date
   	     if(document.getElementById("isRenewable_"+i).checked && document.getElementById("renewalTimeAfterExpire_"+i).value=="")	{

            lcl_message = "<strong>Required Field Missing: </strong>Days to Renew After Expiration Date";
            lcl_false_cnt = lcl_false_cnt + 1;

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("renewalTimeAfterExpire_"+i);
            }
         }

      	//Validate the format of the Days to Renew After Expiration Date
         if(document.getElementById("renewalTimeAfterExpire_"+i).value!="") {
           	var Ok = editRenewalDays.test(document.getElementById("renewalTimeAfterExpire_"+i).value);
           	if(! Ok)	{
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Days to Renew After Expiration Date\" must be in a number format.";
               lcl_false_cnt = lcl_false_cnt + 1;

               if(lcl_false_cnt == 1) {
                  lcl_focus = document.getElementById("renewalTimeAfterExpire_"+i);
               }
            }
         }

         if(lcl_message!="") {
    	    		 inlineMsg(document.getElementById("renewalTimeAfterExpire_"+i).id,lcl_message,10,'renewalTimeAfterExpire_'+i);
         }else{
            clearMsg('renewalTimeAfterExpire_'+i);
       		}
    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
    if(lcl_false_cnt > 0) {
       lcl_focus.focus();
       return false;
    }else{
       lcl_focus_cnt = 0;
    }

  //---------------------------------------------------------------------------
  //Check the "Total Punchcard Uses" for ALL of the rates.
  //---------------------------------------------------------------------------
    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
         lcl_message         = "";
  	 	    var punchcardLimit = /^\d*$/;

      	//Validate the format of the Total Punchcard Uses
         if(document.getElementById("punchcard_limit_"+i).value!="") {
           	var Ok = punchcardLimit.test(document.getElementById("punchcard_limit_"+i).value);
           	if(! Ok)	{
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Total Punchcard Uses\" must be in a number format.";
               lcl_false_cnt = lcl_false_cnt + 1;

               if(lcl_false_cnt == 1) {
                  lcl_focus = document.getElementById("punchcard_limit_"+i);
               }
            }
         }

         if(lcl_message!="") {
    	    		 inlineMsg(document.getElementById("punchcard_limit_"+i).id,lcl_message,10,'punchcard_limit_'+i);
         }else{
            clearMsg('punchcard_limit_'+i);
       		}
    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
    if(lcl_false_cnt > 0) {
       lcl_focus.focus();
       return false;
    }else{
       lcl_focus_cnt = 0;
    }

  //---------------------------------------------------------------------------
  //Check the "Non-Residents Cap" for ALL of the rates.
  //---------------------------------------------------------------------------
//    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
//         lcl_message        = "";
//  	 	    var nonResidentCap = /^\d*$/;

      	//Validate the format of the Non-Residents Cap
//         if(document.getElementById("nonresidentlimit_"+i).value!="") {
//           	var Ok = nonResidentCap.test(document.getElementById("nonresidentlimit_"+i).value);
//           	if(! Ok)	{
//               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Non-Resident Cap\" must be in a number format.";
//               lcl_false_cnt = lcl_false_cnt + 1;

//               if(lcl_false_cnt == 1) {
//                  lcl_focus = document.getElementById("nonresidentlimit_"+i);
//               }
//            }
//         }

//         if(lcl_message!="") {
//    	    		 inlineMsg(document.getElementById("nonresidentlimit_"+i).id,lcl_message,10,'nonresidentlimit_'+i);
//         }else{
//            clearMsg('nonresidentlimit_'+i);
//       		}
//    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
//    if(lcl_false_cnt > 0) {
//       lcl_focus.focus();
//       return false;
//    }else{
//       lcl_focus_cnt = 0;
//    }

  //---------------------------------------------------------------------------
  //Check the format of the "Non-Resident Pre-Registration Start Dates" for all of the rates.
  //---------------------------------------------------------------------------
//    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
//         lcl_message = "";

      	//Validate the format of the Pre-Registration Start Date
//         if(document.getElementById("nonresident_prestartdate_"+i).value!="") {
//           	var Ok = isValidDate(document.getElementById("nonresident_prestartdate_"+i).value);
//           	if(! Ok)	{
//               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Non-Resident Pre-Registration Start Date\" must be in a date format.<br /><span style=\"color:#800000;\">(i.e. mm/dd/yyyy)</span>";
//               lcl_false_cnt = lcl_false_cnt + 1;

//               if(lcl_false_cnt == 1) {
//                  lcl_focus = document.getElementById("nonresident_prestartdate_"+i);
//               }
//            }
//         }

//         if(lcl_message!="") {
//    	    		 inlineMsg(document.getElementById("nonresident_prestartdate_"+i).id,lcl_message,10,'nonresident_prestartdate_'+i);
//         }else{
//            clearMsg('nonresident_prestartdate_'+i);
//       		}
//    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
//    if(lcl_false_cnt > 0) {
//       lcl_focus.focus();
//       return false;
//    }else{
//       lcl_focus_cnt = 0;
//    }

  //---------------------------------------------------------------------------
  //Check the format of the "Non-Resident Pre-Registration End Dates" for all of the rates.
  //---------------------------------------------------------------------------
//    for (i=lcl_i_start; i<=lcl_total_rates; ++ i) {
//         lcl_message = "";

      	//Validate the format of the Pre-Registration End Date
//         if(document.getElementById("nonresident_preenddate_"+i).value!="") {
//           	var Ok = isValidDate(document.getElementById("nonresident_preenddate_"+i).value);
//           	if(! Ok)	{
//               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Non-Resident Pre-Registration End Date\" must be in a date format.<br /><span style=\"color:#800000;\">(i.e. mm/dd/yyyy)</span>";
//               lcl_false_cnt = lcl_false_cnt + 1;

//               if(lcl_false_cnt == 1) {
//                  lcl_focus = document.getElementById("nonresident_preenddate_"+i);
//               }
//            }
//         }

//         if(lcl_message!="") {
//    	    		 inlineMsg(document.getElementById("nonresident_preenddate_"+i).id,lcl_message,10,'nonresident_preenddate_'+i);
//         }else{
//            clearMsg('nonresident_preenddate_'+i);
//       		}
//    }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
//    if(lcl_false_cnt > 0) {
//       lcl_focus.focus();
//       return false;
//    }else{
//       lcl_focus_cnt = 0;
//    }
  <% end if %>

  		lcl_form.submit();

}

 function ChangeOrder(sResidentType,iRateid,iDisplayOrder,iDirection, iMembershipId) {
  	location.href='rate_move.asp?iDisplayOrder='+ iDisplayOrder + '&iRateid=' + iRateid + '&sResidentType=' + sResidentType + '&iDirection=' + iDirection + '&iMembershipId=' + iMembershipId + '&membershiptype=<%=lcl_membership_type%>';
	}

	function ConfirmDelete(sResidentType,iRateid,p_rowid) {
   var sDescription = document.getElementById("description_"+p_rowid).value;
 		var msg = 'Do you wish to delete ' + sDescription + '?';
	 	if (confirm(msg)) {
    			location.href='rate_delete.asp?sResidentType=' + sResidentType + '&iRateid=' + iRateid + '&sMembershipType=<%=lcl_membership_type%>';
 		}
	}

 function enableFields(p_field,p_rowid) {
   if(p_field=="IS_RENEWABLE") {
      if(document.getElementById("isRenewable_"+p_rowid).checked) {
         document.getElementById("renewalstartdate_"+p_rowid).disabled       = false;
         document.getElementById("datepicker_"+p_rowid).style.display        = "inline";
         document.getElementById("renewalTimeAfterExpire_"+p_rowid).disabled = false;
      }else{
         document.getElementById("renewalstartdate_"+p_rowid).disabled       = true;
         document.getElementById("datepicker_"+p_rowid).style.display        = "none";
         document.getElementById("renewalTimeAfterExpire_"+p_rowid).disabled = true;
      }
   }
 }

	function doCalendar(ToFrom) {
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
	}

 function showHidePunchcardLimit(iField,iRowID) {
   lcl_isPunchcard = document.getElementById(iField.id).checked;

   if(lcl_isPunchcard) {
      lcl_disabled = false;
   }else{
      lcl_disabled = true;
      document.getElementById("punchcard_limit_"+iRowID).value="";
   }

   document.getElementById("punchcard_limit_"+iRowID).disabled = lcl_disabled;

 }

function updateRatePublicDisplay() {
  lcl_iMembershipID   = document.getElementById("displayform_iMembershipId").value;
  lcl_sResidentType   = document.getElementById("displayform_sResidentType").value;
  lcl_sMembershipType = document.getElementById("displayform_sMembershipType").value;
  lcl_public_display  = 0;

  if(document.getElementById("displayform_public_display").checked==true) {
     lcl_public_display = 1;
  }

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&iMembershipID='   + encodeURIComponent(lcl_iMembershipID);
  sParameter     += '&sResidentType='   + encodeURIComponent(lcl_sResidentType);
  sParameter     += '&public_display='  + encodeURIComponent(lcl_public_display);
  sParameter     += '&sMembershipType=' + encodeURIComponent(lcl_sMembershipType);

  doAjax('rate_display_change.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function updateNonResidentInfo() {
  lcl_nonresidentlimit         = document.getElementById("nonresidentlimit").value;
  lcl_nonresident_prestartdate = document.getElementById("nonresident_prestartdate").value;
  lcl_nonresident_preenddate   = document.getElementById("nonresident_preenddate").value;

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&NonResidentLimit='        + encodeURIComponent(lcl_nonresidentlimit);
  sParameter     += '&NonResidentPreStartDate=' + encodeURIComponent(lcl_nonresident_prestartdate);
  sParameter     += '&NonResidentPreEndDate='   + encodeURIComponent(lcl_nonresident_preenddate);

  doAjax('updateNonResidentInfo.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}

  //-->
 </script>
</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
 'Set up the membership term fields
  getAdditionalMembershipInfo session("orgid"), _
                              oMembership.MembershipId, _
                              lcl_showMessage, _
                              lcl_showTerms, _
                              lcl_membershipTerms


  lcl_checked_showMessage = ""
  lcl_checked_showTerms   = ""

  if lcl_showMessage then
     lcl_checked_showMessage = " checked=""checked"""
  end if

  if lcl_showTerms then
     lcl_checked_showTerms = " checked=""checked"""
  end if

  response.write "<div id=""content"">" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0""" & lcl_style_tablewidth & ">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <h1>" & session("sOrgName") & " Membership Rates</h1>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"" id=""screenMsg"">&nbsp;</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<div class=""shadow""" & lcl_style_tablewidth & ">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin""" & lcl_style_tablewidth & ">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "		   	<th>Membership Type</th>" & vbcrlf
  response.write "      <th>Resident Type</th>" & vbcrlf
  response.write "      <th>Public Display</th>" & vbcrlf
  response.write "      <th>Display Message</th>" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  'response.write "      <th>Non-Resident</th>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "		<tr id=""membershipTypeHeaderRow"" align=""center"" valign=""top"">" & vbcrlf
  response.write "   			<td valign=""top"">" & vbcrlf
  response.write "          <form name=""membershiptypeform"" id=""membershiptypeform"" method=""post"" action=""poolpass_rates.asp"">" & vbcrlf
  response.write "            <select name=""sMembershipType"" id=""sMembershipType"">" & vbcrlf
                                showMembershipTypePicks(lcl_membership_type)
  response.write "            </select>" & vbcrlf
  response.write "          </form>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "   			<td valign=""top"">" & vbcrlf
  response.write "        		<form name=""nameform"" method=""post"" action=""poolpass_rates.asp"">" & vbcrlf
  response.write "            <input type=""hidden"" name=""sMembershipType"" id=""sMembershipType"" value=""" & lcl_membership_type & """ />" & vbcrlf
  response.write "			      	  <select name=""sResidentType"" onchange=""document.nameform.submit();"">" & vbcrlf
  response.write                ShowResidentPicks(sResidentType)
  response.write "      				  </select>" & vbcrlf
  response.write "    			   </form>" & vbcrlf
  response.write "    		</td>" & vbcrlf
  response.write "			   <td>" & vbcrlf
  response.write "				      <form name=""displayform"" method=""post"" action=""rate_display_change.asp"">" & vbcrlf
  response.write "       					<input type=""hidden"" name=""sResidentType"" id=""displayform_sResidentType"" value=""" & sResidentType & """ />" & vbcrlf
  response.write "       					<input type=""hidden"" name=""iMembershipId"" id=""displayform_iMembershipId"" value=""" & oMembership.MembershipId & """ />" & vbcrlf
  response.write "            <input type=""hidden"" name=""sMembershipType"" id=""displayform_sMembershipType"" value=""" & lcl_membership_type & """ />" & vbcrlf
  response.write "            <input type=""checkbox"" name=""displayform_public_display"" id=""displayform_public_display"" value=""1""" & lcl_checked_publicdisplay & " onclick=""clearScreenMsg();updateRatePublicDisplay();"" />Public Display" & vbcrlf
  response.write "      				</form>" & vbcrlf
  response.write "    		</td>" & vbcrlf
  response.write "   			<td align=""center"">" & vbcrlf
  response.write "          <input type=""checkbox"" name=""showMessage"" id=""showMessage"" value=""on""" & lcl_checked_showMessage & " /> Display Message to Public" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "   			<td align=""right"" valign=""top"" width=""42%"">" & vbcrlf
  response.write "          <input type=""checkbox"" name=""showTerms"" id=""showTerms"" value=""on""" & lcl_checked_showTerms & " /> Show Terms" & vbcrlf
  response.write "          <input type=""button"" name=""editTermsButton"" id=""editTermsButton"" value=""Edit Terms"" class=""button"" />" & vbcrlf
  response.write "          <input type=""button"" name=""saveTermsButton"" id=""saveTermsButton"" value=""Save Terms"" class=""button"" /><br />" & vbcrlf
  response.write "          <div>" & vbcrlf
  response.write "            <textarea name=""membershipTerms"" id=""membershipTerms"">" & lcl_membershipTerms & "</textarea>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write " 	</tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<br /><br /><br /><h2>New Rate for " & sMembershipTypeName & " &mdash; " & sResidentTypeName & "</h2>" & vbcrlf

 'BEGIN: Add Rate -------------------------------------------------------------
 'Display the column headings for the row
  displayRateRowHeaders "rateform_add", _
                        sResidentType, _
                        lcl_membership_type, _
                        lcl_style_tablewidth

 'Display the ADD RATE row
  displayRateRow lcl_rowcount, _
                 lcl_bgcolor, _
                 lcl_rowspan, _
                 sResidentType, _
                 iMaxOrder, _
                 lcl_rateid, _
                 lcl_displayorder, _
                 lcl_description, _
                 lcl_periodid, _
                 lcl_amount, _
                 lcl_maxsignups, _
                 lcl_message, _
                 lcl_attendancetypeid, _
                 lcl_publiccanpurchase, _
                 lcl_isRenewable, _
                 lcl_renewalstartdate, _
                 lcl_isPunchcard, _
                 lcl_punchcardlimit, _
                 lcl_renewaltimeafterexpire, _
                 lcl_cardid, _
                 lcl_isEnabled

 'Close the FORM, TABLE, and DIV tags as they are opened in "displayRateRowHeaders"
  response.write "</table>" & vbcrlf
  response.write "</div><br />" & vbcrlf
 'Display ADD button
  displayButtons "ADD","TOP"
  response.write "</form>" & vbcrlf
  response.write "<br /><br /><br /><br /><h2>Existing Rates for " & sMembershipTypeName & " &mdash; " & sResidentTypeName & "</h2>" & vbcrlf
 'END: Add Rates --------------------------------------------------------------

 'BEGIN: Update Rates ---------------------------------------------------------
 'Display the column headings for the row
  displayRateRowHeaders "rateform_save", _
                        sResidentType, _
                        lcl_membership_type, _
                        lcl_style_tablewidth

	'Get the rows of existing rates
  sSQL = "SELECT rateid, "
  sSQL = sSQL & " residenttype, "
  sSQL = sSQL & " description, "
  sSQL = sSQL & " amount, "
  sSQL = sSQL & " displayorder, "
  sSQL = sSQL & " maxsignups, "
  sSQL = sSQL & " message, "
  sSQL = sSQL & " periodid, "
  sSQL = sSQL & " publiccanpurchase, "
  sSQL = sSQL & " attendancetypeid, "
  sSQL = sSQL & " isRenewable, "
  sSQL = sSQL & " renewalstartdate, "
  sSQL = sSQL & " renewalTimeAfterExpire, "
  sSQL = sSQL & " cardid, "
  sSQL = sSQL & " isPunchcard, "
  sSQL = sSQL & " punchcard_limit, "
  sSQL = sSQL & " isEnabled "
  sSQL = sSQL & " FROM egov_poolpassrates "
  sSQL = sSQL & " WHERE orgid = "       & session("orgid")
  sSQL = sSQL & " AND residenttype = '" & sResidentType & "' "
  sSQL = sSQL & " AND membershipid = "  & oMembership.MembershipId
  sSQL = sSQL & " ORDER BY isEnabled DESC, displayorder "

  set oRate = Server.CreateObject("ADODB.Recordset")
  oRate.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
		
  lcl_bgcolor = "#eeeeee"

  if not oRate.eof then
     do while not oRate.eof
        lcl_rowcount = lcl_rowcount + 1

        if not oRate("isEnabled") then
           lcl_bgcolor = "#c0c0c0"
        end if

        displayRateRow lcl_rowcount, _
                       lcl_bgcolor, _
                       lcl_rowspan, _
                       sResidentType, _
                       iMaxOrder, _
                       oRate("rateid"), _
                       oRate("displayorder"), _
                       oRate("description"), _
                       oRate("periodid"), _
                       oRate("amount"), _
                       oRate("maxsignups"), _
                       oRate("message"), _
                       oRate("attendancetypeid"), _
                       oRate("publiccanpurchase"), _
                       oRate("isRenewable"), _
                       oRate("renewalstartdate"), _
                       oRate("isPunchcard"), _
                       oRate("punchcard_limit"), _
                       oRate("renewalTimeAfterExpire"), _
                       oRate("cardid"), _
                       oRate("isEnabled")

        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

        oRate.MoveNext
     loop
  else
     response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
     response.write "      <td colspan=""10"">" & vbcrlf
     response.write "          No Rates Exist" & vbcrlf
     response.write "          <script language=""javascript"">" & vbcrlf
     response.write "            document.getElementById(""button_save"").style.display=""none"";" & vbcrlf
     response.write "            document.getElementById(""table_save"").style.display=""none"";" & vbcrlf
     response.write "          </script>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
		end if

		oRate.close
		set oRate = nothing

 'Close the FORM, TABLE, and DIV tags as they are opened in "displayRateRowHeaders"
  response.write "  <tr><td colspan=""100""><input type=""hidden"" name=""total_rates"" id=""total_rates"" value=""" & lcl_rowcount & """ size=""5"" /></td></tr>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

  displayButtons "SAVE","BOTTOM"
 'END: Update Rates -----------------------------------------------------------

  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
 'Determine if there are any inline javascripts
  response.write "<script language=""javascript"">" & vbcrlf
  response.write "  setMaxLength();" & vbcrlf

  if lcl_scripts <> "" then
     response.write lcl_scripts
  end if

  response.write "</script>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

Set oMembership = Nothing 
'------------------------------------------------------------------------------
Function GetSeasonInfoByYear( iSeasonYear, ByRef dStartDate, ByRef dEndDate )
	Dim sSQL

	sSQL = "Select seasonid, startdate, enddate FROM egov_poolpassseason WHERE seasonyear = " & iSeasonYear & ""

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oSeason.eof Then 
	  	GetSeasonInfoByYear = oSeason("seasonid")
		  dStartDate          = oSeason("startdate")
		  dEndDate            = oSeason("enddate")
	End If
		
	oSeason.close
	Set oSeason = Nothing

End Function 

'------------------------------------------------------------------------------
Function GetSeasonInfoById( iSeasonId, ByRef dStartDate, ByRef dEndDate )
	Dim sSQL

	sSQL = "Select seasonyear, startdate, enddate FROM egov_poolpassseason WHERE seasonid = " & iSeasonId & ""

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN") , adOpenStatic, adLockReadOnly

	If Not oSeason.eof Then 
		  GetSeasonInfoById = oSeason("seasonyear")
	  	dStartDate        = oSeason("startdate")
	  	dEndDate          = oSeason("enddate")
	End If
		
	oSeason.close
	Set oSeason = Nothing

End Function 

'------------------------------------------------------------------------------
Sub getSeasonsSelect( iSeasonId )
	Dim sSQL

	sSQL = "Select seasonid, seasonyear FROM egov_poolpassseason where orgid = " & Session("OrgID") & " order by seasonyear"

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	Do While Not oSeason.eof 
		response.write vbcrlf & "<option value=" & Chr(34) & oSeason("seasonid") & Chr(34) 
		If clng(iSeasonId) = clng(oSeason("seasonid")) Then
			response.write " selected=" & Chr(34) & "selected" & Chr(34)
		End If 
		response.write ">" &  oSeason("seasonyear") & "</option>"
		oSeason.movenext
	Loop 
		
	oSeason.close
	Set oSeason = Nothing
End Sub 

'------------------------------------------------------------------------------
Function GetMaxDisplayOrder(sResidentType)
	Dim sSql, oMax

	sSQL = "SELECT MAX(displayorder) as MaxOrder "
 sSQL = sSQL & " FROM egov_poolpassrates "
 sSQL = sSQL & " WHERE orgid =" & Session("OrgID")
 sSQL = sSQL & " AND residenttype = '" & sResidentType & "'"

	Set oMax = Server.CreateObject("ADODB.Recordset")
	oMax.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	If IsNull(oMax("MaxOrder")) Then
		GetMaxDisplayOrder = 0
	Else
		GetMaxDisplayOrder = oMax("MaxOrder")
	End If 
	oMax.close
	Set oMax = Nothing

End Function 

'------------------------------------------------------------------------------
Function GetInitialMembershipId( iOrgID )
	Dim sSql, oMember

	sSQL = "Select MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & iOrgID 
	
	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	
	If IsNull(oMember("membershipid")) Then
		GetInitialMembershipId = 0
	Else
		GetInitialMembershipId = oMember("membershipid")
	End If 
	
	oMember.close
	Set oMember = Nothing
End Function 

'------------------------------------------------------------------------------
Function ShowResidentPicks(sResidentType)
	Dim sSQL, oTypes

	' Get the residenttypes
	sSQL = "Select resident_type, description FROM egov_poolpassresidenttypes WHERE orgid = " & Session("OrgID") & " order by displayorder"
	ShowResidentPicks = ""

	Set oTypes = Server.CreateObject("ADODB.Recordset")
	oTypes.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	
	Do While not oTypes.eof 
		ShowResidentPicks = ShowResidentPicks & vbcrlf & "<option value=""" & oTypes("resident_type") & """ "
		If sResidentType = oTypes("resident_type")  Then
			  ShowResidentPicks = ShowResidentPicks & " selected=""selected"" "
			sResidentTypeName = oTypes("description")
		End If 
		ShowResidentPicks = ShowResidentPicks & ">" & oTypes("description") & "</option>"
		oTypes.movenext
	Loop 

	oTypes.close
	Set oTypes = Nothing

End Function 

'------------------------------------------------------------------------------
Function ShowMembershipPicks(iMembershipId, iOrgId)
	Dim sSQL, oMembers

	' Get the memberships
	sSQL = "Select membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & iOrgId & " order by membershipdesc"
	ShowMembershipPicks = ""

	Set oMembers = Server.CreateObject("ADODB.Recordset")
	oMembers.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	
	Do While not oMembers.eof 
		ShowMembershipPicks = ShowMembershipPicks & vbcrlf & "<option value=""" & oMembers("membershipid") & """ "
		If clng(iMembershipId) = clng(oMembers("membershipid"))  Then
			ShowMembershipPicks = ShowMembershipPicks & " selected=""selected"" "
		End If 
		ShowMembershipPicks = ShowMembershipPicks & ">" & oMembers("membershipdesc") & "</option>"
		oMembers.movenext
	Loop 

	oMembers.close
	Set oMembers = Nothing

End Function 

'------------------------------------------------------------------------------
Function ShowPeriodPicks(iPeriodId, iOrgId)
	Dim sSQL, oPeriods

	' Get the Periods
	sSQL = "Select periodid, period_desc FROM egov_membership_periods WHERE orgid = " & iOrgId & " order by period_desc DESC"
	ShowPeriodPicks = ""

	Set oPeriods = Server.CreateObject("ADODB.Recordset")
	oPeriods.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	
	Do While not oPeriods.eof 
		ShowPeriodPicks = ShowPeriodPicks & vbcrlf & "<option value=""" & oPeriods("periodid") & """ "
		If clng(iPeriodId) = clng(oPeriods("periodid"))  Then
			ShowPeriodPicks = ShowPeriodPicks & " selected=""selected"" "
		End If 
		ShowPeriodPicks = ShowPeriodPicks & ">" & oPeriods("period_desc") & "</option>"
		oPeriods.movenext
	Loop 

	oPeriods.close
	Set oPeriods = Nothing

End Function 

'------------------------------------------------------------------------------
Function GetPublicDisplay(iMembershipId, sResidentType)
	Dim sSql
	GetPublicDisplay = ""

	sSQL = "Select public_display FROM egov_membership_rate_displays WHERE membershipid = " & iMembershipId & "and resident_type = '" & sResidentType & "'"
	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oDisplay.EOF Then 
		If oDisplay("public_display") Then
			GetPublicDisplay = "checked=""checked"" "
		End If 
	End If 

	oDisplay.close
	Set oDisplay = Nothing

End Function 

'------------------------------------------------------------------------------
Sub ShowPreselects( iRateId, p_rowid )
	Dim sSql, oFamily 

	sSQL = "SELECT relationship FROM egov_familymember_relationships WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oFamily = Server.CreateObject("ADODB.Recordset")
	oFamily.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	Do While Not oFamily.EOF  
  		response.write "<input type=""checkbox"" name=""relation_" & p_rowid & """ id=""relation_" & p_rowid & """ value=""" & oFamily("relationship") & """ " & GetPreselected( iRateId, oFamily("relationship") ) & " /> " & vbcrlf

  		If LCase(oFamily("relationship")) = "yourself" Then 
    			response.write " Purchaser"
   	Else
			    response.write oFamily("relationship") 
  		End If 

  		response.write "<br />" & vbcrlf
  		oFamily.MoveNext
	Loop 
		
	oFamily.close
	Set oFamily = Nothing
End Sub  

'------------------------------------------------------------------------------
Function GetPreselected( iRateId, sRelation )
	Dim oCmd
	GetPreselected = ""

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'	    .CommandText = "CheckPoolPassPreselected"
'	    .CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@iRateid", 3, 1, 4, iRateId)
'		.Parameters.Append oCmd.CreateParameter("@sRelation", 200, 1, 20, sRelation)
'		.Parameters.Append oCmd.CreateParameter("@Preselected", 11, 2, 1)
'	    .Execute
'	End With
		
'	GetPreselected = oCmd.Parameters("@Preselected").Value

'	Set oCmd = Nothing

	sSQL = "Select count(relation) as hits FROM egov_poolpasspreselected WHERE rateid = " & iRateId & " and relation = '" & sRelation & "'"

	Set oSelected = Server.CreateObject("ADODB.Recordset")
	oSelected.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If clng(oSelected("hits")) <> 0 Then 
		GetPreselected = " checked=""checked"" "
	End If
		
	oSelected.close
	Set oSelected = Nothing
End Function 

'-------------------------------------------------------------------------
sub showAttendanceTypes(p_value,p_rowid)
  response.write "<select name=""attendancetypeid_" & p_rowid & """ id=""attendancetypeid_" & p_rowid & """>" & vbcrlf

  sSQL = "SELECT attendancetypeid, attendancetype, isactive, isdefault "
  sSQL = sSQL & " FROM egov_pool_attendancetypes "
  sSQL = sSQL & " WHERE isactive = 1 "
  sSQL = sSQL & " ORDER BY attendancetypeid "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN") , 3, 1

  if not rs.eof then
     while not rs.eof
        if p_value = "" then
           if rs("isdefault") then
              lcl_selected = " selected"
           else
              lcl_selected = ""
           end if
        else
           if CLng(p_value) = CLng(rs("attendancetypeid")) then
              lcl_selected = " selected"
           else
              lcl_selected = ""
           end if
        end if

        response.write "  <option value=""" & rs("attendancetypeid") & """" & lcl_selected & ">" & rs("attendancetype") & "</option>" & vbcrlf

        rs.movenext
     wend
  end if

  set rs = nothing

  response.write "</select>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub showMembershipTypePicks(p_membership_type)

  sSQL = "SELECT membershipid, membership, membershipdesc "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY membershipdesc "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     while not rs.eof
        if UCASE(p_membership_type) = UCASE(rs("membership")) then
           lcl_selected = " selected=""selected"""
	   sMembershipTypeName = rs("membershipdesc")
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & rs("membership") & """" & lcl_selected & ">" & rs("membershipdesc") & "</option>" & vbcrlf
        rs.movenext
     wend
  end if

  set rs = nothing

end sub

'------------------------------------------------------------------------------
sub displayButtons(iType,iTopBottom)
  lcl_buttonType = "SAVE"
  lcl_topBottom  = "TOP"

  if iTopBottom <> "" then
     lcl_topBottom = UCASE(iTopBottom)
  end if

  if iTopBottom = "BOTTOM" then
     lcl_padding = "padding-top: 5px"
  else
     lcl_padding = "padding-bottom: 5px"
  end if

  if iType <> "" then
     lcl_buttonType = UCASE(iType)
  end if

  if iType = "ADD" then
     lcl_buttonid    = "button_add"
     lcl_buttonvalue = "Add Rate"
  else
     lcl_buttonid    = "button_save"
     lcl_buttonvalue = "Save Changes"
  end if

  response.write "<div style=""" & lcl_padding & """>" & vbcrlf
  response.write "  <input type=""button"" name=""sAction"" id=""" & lcl_buttonid & """ value=""" & lcl_buttonvalue & """ class=""button"" onclick=""SaveRate('" & lcl_buttonType & "');"" />" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayRateRowHeaders(iFormName, _
                          iResidentType, _
                          iMembershipType, _
                          iStyleTableWidth)

  if UCASE(iFormName) = "RATEFORM_ADD" then
     lcl_tableid = "table_add"
     lcl_button  = "ADD"
  else
     lcl_tableid = "table_save"
     lcl_button  = "SAVE"
  end if

 'Setup the add/update form
  response.write "<form name=""" & iFormName & """ id=""" & iFormName & """ method=""post"" action=""rate_save.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""sResidentType"" value=""" & iResidentType & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iMembershipId"" value=""" & oMembership.MembershipId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""sMembershipType"" id=""sMembershipType"" value=""" & iMembershipType & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ />" & vbcrlf

 ''Display ADD button
  'displayButtons lcl_button,"TOP"

 'Open Table and display Column Headers
  response.write "<div class=""shadow""" & lcl_style_tablewidth & ">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""" & lcl_tableid & """ class=""tableadmin""" & iStyleTableWidth & ">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <th>Description</th>" & vbcrlf
  response.write "      <th>Membership<br />Period</th>" & vbcrlf
  response.write "      <th>Amount</th>" & vbcrlf
  response.write "      <th>Max On<br />Pass</th>" & vbcrlf
  response.write "      <th>Message</th>" & vbcrlf
  response.write "      <th>Preselected</th>" & vbcrlf

  if lcl_userhaspermission_pool_attendance_view AND lcl_orghasfeature_pool_attendance_view then
     response.write "      <th>Attendance Type<br />(Reporting)</th>" & vbcrlf
  end if

  response.write "      <th>Allow<br />Public<br />Purchase</th>" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayRateRow(iRowCount, _
                   iBGColor, _
                   iRowSpan, _
                   iResidentType, _
                   iMaxOrder, _
                   iRateID, _
                   iDisplayOrder, _
                   iDescription, _
                   iPeriodID, _
                   iAmount, _
                   iMaxSignUps, _
                   iMessage, _
                   iAttendanceTypeID, _
                   iPublicCanPurchase, _
                   iIsRenewable, _
                   iRenewalStartDate, _
                   iIsPunchcard, _
                   iPunchcardLimit, _
                   iRenewalTimeAfterExpire, _
                   iCardID, _
                   iIsEnabled)

 'Setup/Format fields
  lcl_checked_publiccanpurchase = ""
  lcl_checked_disableRate       = ""
  lcl_renewalstartdate          = ""
  lcl_style_borderbottom        = ""

  if iPublicCanPurchase then
     lcl_checked_publiccanpurchase = " checked=""checked"""
  end if

  if not iIsEnabled then
     lcl_checked_disableRate = " checked=""checked"""
  end if

  if iRenewalStartDate <> "" then
     lcl_renewalstartdate = replace(iRenewalStartDate,"1/1/1900","")
  end if

  if iRowCount > 0 AND iRowCount < iMaxOrder then
     lcl_style_borderbottom = " style=""border-bottom: 1pt solid #336699"""
  end if

 'BEGIN: Rates First Line -----------------------------------------------------
  response.write "  <tr valign=""top"" bgcolor=""" & iBGColor & """>" & vbcrlf
  response.write "      <td nowrap=""nowrap"">" & vbcrlf
  response.write "          <input type=""hidden"" name=""rateid_"       & iRowCount & """ id=""rateid_"       & iRowCount & """ size=""5"" value=""" & iRateID & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""displayorder_" & iRowCount & """ id=""displayorder_" & iRowCount & """ value=""" & iDisplayOrder & """ />" & vbcrlf

  if not lcl_userhaspermission_pool_attendance_view then
     response.write "<input type=""hidden"" name=""attendancetypeid_" & iRowCount & """ value=""1"" size=""1"" maxlength=""1"">" & vbcrlf
  end if

 'Description
  response.write "          <input type=""text"" name=""description_"    & iRowCount & """ id=""description_"  & iRowCount & """ value=""" & iDescription & """ size=""20"" maxlength=""50"" onchange=""clearMsg('description_" & iRowCount & "');"" />" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Membership Period
  response.write "      <td>" & vbcrlf
  response.write "          <select name=""iPeriodId_" & iRowCount & """ id=""iPeriodId_" & iRowCount & """>" & vbcrlf
                              oMembership.ShowMembershipPeriodPicks iPeriodID
  response.write "          </select>" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Amount
  response.write "      <td nowrap=""nowrap"">" & vbcrlf
  response.write "          <input type=""text"" name=""amount_" & iRowCount & """ id=""amount_" & iRowCount & """ value=""" & FormatNumber(iAmount,2) & """ size=""7"" maxlength=""10"" onchange=""clearMsg('amount_" & iRowCount & "');"" />" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Max Signups (Max on Pass)
  response.write "      <td nowrap=""nowrap"">" & vbcrlf
  response.write "          <input type=""text"" name=""maxsignups_" & iRowCount & """ id=""maxsignups_" & iRowCount & """ value=""" & iMaxSignUps & """ size=""5"" maxlength=""3"" onchange=""clearMsg('maxsignups_" & iRowCount & "');"" />" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Message
  response.write "      <td rowspan=""" & iRowSpan & """ nowrap=""nowrap""" & lcl_style_borderbottom & ">" & vbcrlf
  response.write "          <textarea name=""message_" & iRowCount & """ id=""message_" & iRowCount & """ class=""message"" rows=""7"" cols=""30"" maxlength=""500"">" & iMessage & "</textarea>" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Pre-Selected
  response.write "      <td rowspan=""" & iRowSpan & """ nowrap=""nowrap""" & lcl_style_borderbottom & ">" & vbcrlf
                            ShowPreselects iRateID, iRowCount
  response.write "      </td>" & vbcrlf

 'Attendance Types (Reporting)
  if lcl_userhaspermission_pool_attendance_view AND lcl_orghasfeature_pool_attendance_view then
     'response.write "<td rowspan=""" & iRowSpan & """>" & showAttendanceTypes oRate("attendancetypeid"),iRowCount & "</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
                               showAttendanceTypes iAttendanceTypeID, iRowCount
     response.write "      </td>" & vbcrlf
  end if

 'Public Can Purchase
  response.write "      <td align=""center"">" & vbcrlf
  response.write "          <input type=""checkbox"" name=""publiccanpurchase_" & iRowCount & """ id=""publiccanpurchase_" & iRowCount & """" & lcl_checked_publiccanpurchase & " />" & vbcrlf
  response.write "      </td>" & vbcrlf

 'Delete Button
  if iRowCount > 0 then

    'Check to see if the rate exists on ANY purchases.
    'If "yes" then do NOT allow the deletion.
    'If "no" then DO show the "delete" button.
     sDisplayButtonRemove   = "&nbsp;"
     sDisplayButtonMoveUp   = ""
     sDisplaybuttonMoveDown = ""
     sRateExistsOnPurchase  = checkRateExistsOnPurchase(session("orgid"), _
                                                       iRateID)

     if not sRateExistsOnPurchase then
        sDisplayButtonRemove = "<input type=""button"" name=""sAction"" value=""Delete"" class=""button"" onclick=""ConfirmDelete('" & iResidentType & "'," & iRateID & "," & iRowCount & ");"" />"
     else
        sDisplayButtonRemove = "<div style=""text-align:center;"">"
        sDisplayButtonRemove = sDisplayButtonRemove & "<fieldset class=""fieldset"">"
        sDisplayButtonRemove = sDisplayButtonRemove & "<legend>Disable<br />Rate</legend>"
        sDisplayButtonRemove = sDisplayButtonRemove & "<input type=""checkbox"" name=""disableRate_" & iRowCount & """ id=""disableRate_" & iRowCount & """ value=""Y""" & lcl_checked_disableRate & " />"
        sDisplayButtonRemove = sDisplayButtonRemove & "</fieldset>"
        sDisplayButtonRemove = sDisplayButtonRemove & "</div>"
     end if

     if iRowCount <> 1 then
        sDisplayButtonMoveUp = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move Row UP');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder('" & iResidentType & "', " & iRateID & ", " & iDisplayOrder & ", -1, " & oMembership.MembershipId & ");""/><br />" & vbcrlf
     end if

     if iDisplayOrder <> iMaxOrder then
        sDisplayButtonMoveDown = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move Row DOWN');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder('" & iResidentType & "', " & iRateID & ", " & iDisplayOrder & ", 1, " & oMembership.MembershipId & ");"">" & vbcrlf
     end if

     response.write "      <td nowrap=""nowrap"" class=""action"">" & sDisplayButtonRemove & "</td>" & vbcrlf
     response.write "      <td nowrap=""nowrap"">" & sDisplaybuttonMoveUp & sDisplayButtonMoveDown & "</td>" & vbcrlf
  else
     response.write "      <td colspan=""2"">&nbsp;</td>" & vbcrlf
  end if

  response.write "  </tr>" & vbcrlf
 'END: Rates First Line -------------------------------------------------------

  response.write "  <tr bgcolor=""" & iBGColor & """>" & vbcrlf

 'BEGIN: Renewal --------------------------------------------------------------
  if lcl_orghasfeature_membership_renewals then
     if lcl_userhaspermission_membership_renewals then
    				if iIsRenewable then
           lcl_checked  = " checked=""checked"""
           lcl_showHide = "inline"
           lcl_scripts  = lcl_scripts & "document.getElementById(""renewalstartdate_" & iRowCount & """).disabled=false;" & vbcrlf
           lcl_scripts  = lcl_scripts & "document.getElementById(""renewalTimeAfterExpire_" & iRowCount & """).disabled=false;" & vbcrlf
        else
           lcl_checked  = ""
           lcl_showHide = "none"
           lcl_scripts  = lcl_scripts & "document.getElementById(""renewalTimeAfterExpire_" & iRowCount & """).disabled=true;" & vbcrlf
        end if

        lcl_scripts = lcl_scripts & "document.getElementById(""datepicker_" & iRowCount & """).style.display=""" & lcl_showHide & """;" & vbcrlf

        response.write "      <td>&nbsp;&nbsp;&nbsp;Is Rate Renewable? <input type=""checkbox"" name=""isRenewable_" & iRowCount & """ id=""isRenewable_" & iRowCount & """ value=""1"" onclick=""enableFields('IS_RENEWABLE'," & iRowCount & ");clearMsg('renewalstartdate_" & iRowCount & "');""" & lcl_checked & " /></td>" & vbcrlf
        response.write "      <td colspan=""3"">" & vbcrlf
        response.write "          Renewal Start Date: <input type=""text"" name=""renewalstartdate_" & iRowCount & """ id=""renewalstartdate_" & iRowCount & """ size=""10"" maxlength=""10"" value=""" & lcl_renewalstartdate & """ onchange=""clearMsg('renewalstartdate_" & iRowCount & "');"" DISABLED />" & vbcrlf
        response.write "          &nbsp;<img src=""../images/calendar.gif"" id=""datepicker_" & iRowCount & """ border=""0"" style=""cursor:hand;"" class=""hotspot"" onmouseover=""tooltip.show('Click to View Calendar');"" onmouseout=""tooltip.hide();"" onclick=""doCalendar('renewalstartdate_" & iRowCount & "')"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
     else
        response.write "      <td colspan=""4"">" & vbcrlf
        response.write "          <input type=""hidden"" name=""isRenewable_" & iRowCount & """ id=""isRenewable_" & iRowCount & """ value=""" & replace(replace(iIsRenewable,True,1),False,0) & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""renewalstartdate_" & iRowCount & """ id=""renewalstartdate_" & iRowCount & """ value=""" & iRenewalStartDate & """ size=""10"" maxlength=""10"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
     end if
  else
     response.write "      <td colspan=""4"">" & vbcrlf
     response.write "          <input type=""hidden"" name=""isRenewable_" & iRowCount & """ id=""isRenewable_" & iRowCount & """ value=""0"" />" & vbcrlf
     response.write "          <input type=""hidden"" name=""renewalstartdate_" & iRowCount & """ id=""renewalstartdate_" & iRowCount & """ value="""" size=""10"" maxlength=""10"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
  end if
 'END: Renewal ----------------------------------------------------------------

 'BEGIN: Punchcard ------------------------------------------------------------
  response.write "      <td valign=""top"" colspan=""4"" rowspan=""3""" & lcl_style_borderbottom & ">" & vbcrlf

  if lcl_orghasfeature_pool_attendance_view then
     if lcl_userhaspermission_pool_attendance_view then
        lcl_checked_punchcard   = ""
        lcl_isPunchcardDisabled = "true"

        if iIsPunchcard then
           lcl_checked_punchcard   = " checked=""checked"""
           lcl_isPunchcardDisabled = "false"
        end if

        response.write "          Is this a punchcard?&nbsp;" & vbcrlf
        response.write "          <input type=""checkbox"" name=""isPunchcard_" & iRowCount & """ id=""isPunchcard_" & iRowCount & """ value=""on"" onclick=""clearMsg('punchcard_limit_" & iRowCount & "');showHidePunchcardLimit(this,'" & iRowCount & "')""" & lcl_checked_punchcard & " />" & vbcrlf
        response.write "          <br />" & vbcrlf
        response.write "          Total Punchcard Uses:&nbsp;" & vbcrlf
        response.write "          <input type=""text"" name=""punchcard_limit_" & iRowCount & """ id=""punchcard_limit_" & iRowCount & """ size=""3"" maxlength=""10"" onchange=""clearMsg('punchcard_limit_" & iRowCount & "');"" value=""" & iPunchcardLimit & """ />" & vbcrlf

        lcl_scripts = lcl_scripts & "document.getElementById(""punchcard_limit_" & iRowCount & """).disabled=" & lcl_isPunchcardDisabled & ";" & vbcrlf

     else
        response.write "<input type=""hidden"" name=""isPunchcard_" & iRowCount & """ id=""isPunchcard_" & iRowCount & """ value=""" & iIsPunchcard & """ />" & vbcrlf
        response.write "<input type=""hidden"" name=""punchcard_limit_" & iRowCount & """ id=""punchcard_limit_" & iRowCount & """ value=""" & iPunchcardLimit & """ size=""10"" maxlength=""10"" />" & vbcrlf
     end if
  else
     response.write "<input type=""hidden"" name=""isPunchcard_" & iRowCount & """ id=""isPunchcard_" & iRowCount & """ value=""0"" />" & vbcrlf
     response.write "<input type=""hidden"" name=""punchcard_limit_" & iRowCount & """ id=""punchcard_limit_" & iRowCount & """ value=""0"" size=""10"" maxlength=""10"" />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
 'END: Punchcard --------------------------------------------------------------

  response.write "  </tr>" & vbcrlf
  response.write "  <tr bgcolor=""" & iBGColor & """>" & vbcrlf
  response.write "      <td align=""center"" colspan=""4"">" & vbcrlf

 'BEGIN: "Days to Renew...." --------------------------------------------------
  if lcl_orghasfeature_membership_renewals then
     if lcl_userhaspermission_membership_renewals then
        response.write "          Days to Renew after Expiration Date: <input type=""text"" name=""renewalTimeAfterExpire_" & iRowCount & """ id=""renewalTimeAfterExpire_" & iRowCount & """ size=""3"" maxlength=""5"" value=""" & iRenewalTimeAfterExpire & """ onchange=""if(this.value==''){this.value=0;};clearMsg('renewalTimeAfterExpire_" & iRowCount & "');"" DISABLED />" & vbcrlf
     else
        response.write "          <input type=""hidden"" name=""renewalTimeAfterExpire_" & iRowCount & """ id=""renewalTimeAfterExpire_" & iRowCount & """ size=""3"" maxlength=""5"" value=""" & iRenewalTimeAfterExpire & """ />" & vbcrlf
     end if
  else 
     response.write "          <input type=""hidden"" name=""renewalTimeAfterExpire_" & iRowCount & """ id=""renewalTimeAfterExpire_" & iRowCount & """ size=""3"" maxlength=""5"" value=""0"" />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
 'END: "Days to Renew...." ----------------------------------------------------

  response.write "  </tr>" & vbcrlf
  response.write "  <tr bgcolor=""" & iBGColor & """>" & vbcrlf

 'BEGIN: Multiple Card Layouts ------------------------------------------------
  response.write "      <td colspan=""4""" & lcl_style_borderbottom & ">" & vbcrlf

  if lcl_orghasfeature_card_layout_multiplelayouts then
     response.write "          &nbsp;&nbsp;&nbsp;Card Layout:&nbsp;" & vbcrlf
     response.write "          <select name=""cardid_" & iRowCount & """ id=""cardid_" & iRowCount & """>" & vbcrlf
                                       displayCardLayoutOptions iCardID,"Y"
     response.write "          </select>" & vbcrlf
  else
     response.write "          <input type=""hidden"" name=""cardid_" & iRowCount & """ id=""cardid_" & iRowCount & """ value=""" & iCardID & """ />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
 'END: Multiple Card Layouts --------------------------------------------------

  response.write "  </tr>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub getAdditionalMembershipInfo(ByVal iOrgID, _
                                ByVal iMembershipID, _
                                ByRef lcl_showMessage, _
                                ByRef lcl_showTerms, _
                                ByRef lcl_membershipTerms)

  dim sOrgID, sMembershipID

  sOrgID              = 0
  sMembershipID       = 0
  lcl_showMessage     = false
  lcl_showTerms       = false
  lcl_membershipTerms = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iMembershipID <> "" then
     sMembershipID = clng(iMembershipID)
  end if

  sSQL = "SELECT showmessage, "
  sSQL = sSQL & " showterms, "
  sSQL = sSQL & " membershipterms "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND membershipid = " & sMembershipID

  set oGetAdditionalMembershipInfo = Server.CreateObject("ADODB.Recordset")
  oGetAdditionalMembershipInfo.Open sSQL, Application("DSN"), 3, 1

  if not oGetAdditionalMembershipInfo.eof then
     lcl_showMessage     = oGetAdditionalMembershipInfo("showmessage")
     lcl_showTerms       = oGetAdditionalMembershipInfo("showterms")
     lcl_membershipTerms = oGetAdditionalMembershipInfo("membershipterms")
  end if

  oGetAdditionalMembershipInfo.close
  set oGetAdditionalMembershipInfo = nothing

end sub

'------------------------------------------------------------------------------
function checkRateExistsOnPurchase(iOrgID, _
                                   iRateID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT count(poolpassid) as totalPoolPasses "
  sSQL = sSQL & " FROM egov_poolpasspurchases "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND rateid = " & iRateID

  set oRateExistsOnPurchase = Server.CreateObject("ADODB.Recordset")
  oRateExistsOnPurchase.Open sSQL, Application("DSN"), 3, 1

  if not oRateExistsOnPurchase.eof then
     if oRateExistsOnPurchase("totalPoolPasses") > 0 then
        lcl_return = true
     end if
  end if

  oRateExistsOnPurchase.close
  set oRateExistsOnPurchase = nothing

  checkRateExistsOnPurchase = lcl_return

end function
%>
