<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_rates.asp
' AUTHOR: Steve Loar
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/31/06 Steve Loar - Code added
' 1.1	 10/05/06	Steve Loar - Security, Header and nav changed
' 1.2  09/08/08 David Boyer - Added Membership Renewals
' 1.3  03/09/09 David Boyer - Added Alternate Layout feature
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iUserId
 dim iMembershipId, oMembership, iPeriodId

 sLevel     = "../"  'Override of value from common.asp
 lcl_onload = ""

 if not userhaspermission(session("userid"),"purchase membership") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 set oMembership = New classMembership

'Set the membershipid to the one for pools, later this will be a dropdown
 if request("sMembershipType") <> "" then
    lcl_membership_type = request("sMembershipType")
 else
    'lcl_membership_type = "pool"
    lcl_membership_type = oMembership.GetFirstMembershipType()
 end if
 oMembership.SetMembershipId( lcl_membership_type )
 iMembershipId = oMembership.MembershipId

 iUserId       = getFirstUserID()
' sUserType     = ""
 iPeriodId     = clng(request("periodid"))
 iRateDesc     = request("rateDesc")

 if request("userid") <> "" then
   	iUserId = request("userid")
 end if

' sUserType     = GetUserResidentType(iUserid)
 'sResidentDesc = GetResidentTypeDesc(sUserType)

 if request("periodid") = "" then
   	iPeriodId = oMembership.GetFirstMembershipPeriodId( iMembershipId )
 end if

' if request("rateDesc") = "" then
'   	iRateDesc = oMembership.getFirstMembershipRateOption_AltLayout( sUserType, iMembershipId)
' end if

'Determine if the membership period selected is/isn't seasonal
 lcl_isSeasonal = checkIsSeasonal(iPeriodID)

 session("RedirectPage") = "../poolpass/poolpass_form.asp?userid=" & iUserId & "&iMembershipId=" & iMembershipId & "&periodid=" & iPeriodId
 session("RedirectLang") = "Return to Membership Purchase"

'Check for org features
 lcl_orghasfeature_membership_renewals            = orghasfeature("membership_renewals")
 lcl_orghasfeature_purchase_membership_alt_layout = orghasfeature("purchase_membership_alt_layout")

'Check for user permissions
 lcl_userhaspermission_membership_renewals = userhaspermission(session("userid"),"membership_renewals")

'Check to see if the org has the feature turned-on and the user has it assigned
 lcl_membershiprenewals_feature = "N"
 lcl_membership_alt_layout      = "N"

 if lcl_orghasfeature_membership_renewals AND lcl_userhaspermission_membership_renewals then
    lcl_membershiprenewals_feature = "Y"
 end if

 if lcl_orghasfeature_purchase_membership_alt_layout then
    lcl_membership_alt_layout = "Y"
 end if
%>
<html lang="en">
<head>
  <meta charset="UTF-8">
  
 	<title>E-Gov Administration Console {Membership Purchase}</title>

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="style_pool.css" />
  <link rel="stylesheet/less" type="text/less" href="stylesPoolPass.less" />

  <script src="../scripts/jquery-1.9.1.min.js"></script>
  <script src="../scripts/less-1.3.3.min.js"></script>

<script>
  <!--

  function membershipTypeChange()
  {
	if (document.getElementById('periodid'))
	{
		buildRightSideBottom($('#periodid').val());
	}
  }
$(document).ready(function() {
   //Set up the Name Search/User Info
   $('#selectName').css('display','none');
   $('#userInfoDiv').css('display','none');
   $('#membershipRenewals').css('display','none');
   $('#editUserButton').prop('disabled',true);
   $('#continueButton').prop('disabled',true);
   $('#searchname').focus();

   //Set up the Membership Rates/Types picks
   $('#rightsidepicks_top').html('');
   $('#rightsidepicks_bottom').html('');

   //alert('buildMembershipPurchaseSections.asp?purchaseSection=rightsidepicks&orgid=<%=session("orgid")%>&membershipid=<%=iMembershipID%>&periodid=<%=iPeriodID%>');

   //BEGIN: Build the right-side "top" picks ----------------------------------
//   $.post('buildMembershipPurchaseSections.asp', {
//      purchaseSection: 'rightsidepicks_top',
//      orgid:           '<%'session("orgid")%>',
//      membershipid:    '<%'iMembershipId%>',
      //rateDesc:        '<%'iRateDesc%>',
      //usertype:        '<%'sUserType%>',
//      periodid:        '<%'iPeriodID%>'
//   }, function(results_top) {
//      $('#rightsidepicks_top').html(results_top);

      //buildRightSideBottom('<%'iRateDesc%>', '<%'iPeriodID%>');
//      buildRightSideBottom('<%'iPeriodID%>');
//   });
   //END: Build the right-side "top" picks ------------------------------------

   //BEGIN: Search Button - Click ---------------------------------------------
   $('#searchButton').click(function() {
      //Clear/Close each section so we can start with a fresh search.
      //As we get the data we can expand the sections.
      $('#membershipRenewals').slideUp('slow',function() {
         $('#userInfoDiv').slideUp('slow',function() {
            $('#selectName').slideUp('slow', function() {
               var lcl_nameSearch = $('#searchname').val();

               //alert('buildMembershipPurchaseSections.asp?orgid=<%'session("orgid")%>&purchaseSection=namesearch_dropdown_options&namesearch=' + lcl_nameSearch);

               if(lcl_nameSearch == '')
               {
                  $('#editUserButton').prop('disabled',true);
               }
               else
               {
                  $.post('buildMembershipPurchaseSections.asp', {
                     purchaseSection: 'namesearch_dropdown_options',
                     orgid:           '<%=session("orgid")%>',
                     namesearch:      lcl_nameSearch
                  }, function(result) {
                     if(result == '') {
                        //$('#selectName').slideUp('slow');
                        $('#selectName').html('<div class=\'noResultsFoundMsg\'>No Results Found</div>');
                        $('#selectName').slideDown('slow');

                        hideRightSidePicks();
                     }
                     else
                     {
                        $('#selectName').html(result);

                        $('#selectName').slideDown('slow',function() {

                           if($('#userid').val() != '') {

                              var lcl_periodid = '';

                              $('input[name^="periodid"]').each(function(index) {
                                 //lcl_periodid = $('#periodid_' + index+1).val();
                                 if($(this).prop('checked')) {
                                    lcl_periodid = $(this).prop('id');
                                    lcl_periodid = lcl_periodid.replace('periodid_','');
                                    lcl_periodid = lcl_periodid.replace('periodid','');
                                 }
                              });

                              updateUserInfo($('#userid').val());
                           }
                        });
                     }
                  });
               }
            });
         });
      });
   });
   //END: Search Button - Click -----------------------------------------------

   $('#editUserButton').click(function() {
      var lcl_userid = $('#userid option:selected').val();

      location.href='../dirs/update_citizen.asp?userid=' + lcl_userid;
   });
});

//function buildRightSideBottom(iRateDesc, iPeriodID) {
function buildRightSideBottom(iPeriodID) {
   $('#rightsidepicks_bottom').slideUp('slow', function() {
      $('#rightsidepicks_bottom').html('');

      var lcl_userid = $('#userid option:selected').val();
//alert('buildMembershipPurchaseSections.asp?purchaseSection=rightsidepicks_bottom&orgid=<%=session("orgid")%>&membershipid=<%=iMembershipID%>&userid=' + lcl_userid + '&periodid=' + iPeriodID);
      //Build the right-side "bottom" picks
      if(lcl_userid != '') {
         var lcl_ratedesc = $('#rateDesc option:selected').val();

         $.post('buildMembershipPurchaseSections.asp', {
            purchaseSection: 'rightsidepicks_bottom',
            orgid:           '<%=session("orgid")%>',
            membershipid:    document.getElementById("iMembershipId").value,
            userid:          lcl_userid,
            rateDesc:        lcl_ratedesc,
            //usertype:        '<%'sUserType%>',
            periodid:        iPeriodID
         }, function(results_bottom) {
            $('#rightsidepicks_bottom').html(results_bottom);
            //$('#rightsidepicks_bottom').slideDown('slow');
            $('#rightsidepicks_bottom').css("display","inline-block");

           	$('#continueButton').prop('disabled',true);
		if (results_bottom.indexOf("radio") >= 0)
		{
            		$('#continueButton').prop('disabled',false);
		}
         });
      }
   });
}

function updateUserInfo(iFieldValue)
{
   var lcl_fieldvalue = '';

   $('#editUserButton').prop('disabled',true);

   $('#membershipRenewals').slideUp('slow',function() {
      $('#userInfoDiv').slideUp('slow',function() {
         $('#userInfoDiv').html();

         if(iFieldValue != '')
         {
            lcl_fieldvalue = iFieldValue;

            //First update the user information
            $.post('buildMembershipPurchaseSections.asp', {
               purchaseSection: 'show_user_info',
               orgid:           '<%=session("orgid")%>',
               userid:          lcl_fieldvalue
               //usertype:        '<%'sUserType%>'
            }, function(result) {
               if(result != '') {
                  $('#userInfoDiv').html(result);


                  updatePurchaseOptions();

                  //Before expanding the user/renewal sections, we need to get the renewal info.
                  //But only if the feature is enabled and any exists.
                  //var lcl_periodid = $("input[name=periodid]").val()
                  var lcl_membershiprenewals_feature = '<%=lcl_membershiprenewals_feature%>';

                  if(lcl_membershiprenewals_feature == 'Y') {
                     var lcl_periodid = '';

                     $('#userInfoDiv').slideDown(function() {
                        $('#editUserButton').prop('disabled',false);
                     });

                     updateRenewalInfo();
                  }
                  else
                  {
                     $('#userInfoDiv').slideDown(function() {
                        $('#editUserButton').prop('disabled',false);
                     });
                  }
               }
               else
               {
                  $('#selectName').slideUp('slow');
                  $('#selectName').html('<div class=\'noResultsFoundMsg\'>No Results Found</div>');
                  $('#selectName').slideDown('slow');

                  hideRightSidePicks();
               }
              
            });
         }
      });
   });
}

function updatePurchaseOptions() {
   var lcl_userid = $('#userid option:selected').val();
//alert('buildMembershipPurchaseSections.asp?purchaseSection=rightsidepicks_top&orgid=<%=session("orgid")%>&membershipid=<%=iMembershipID%>&userid=' + lcl_userid + '&periodid=' + lcl_periodid);

   if(lcl_userid != '') {
      $.post('buildMembershipPurchaseSections.asp', {
         purchaseSection: 'rightsidepicks_top',
         orgid:           '<%=session("orgid")%>',
         membershipid:    '<%=iMembershipId%>',
         userid:          lcl_userid,
         //rateDesc:        '<%'iRateDesc%>',
         //usertype:        '<%'sUserType%>',
         periodid:        '<%=iPeriodID%>'
      }, function(results_top) {
         $('#rightsidepicks_top').html(results_top);

         $('#rightsidepicks_top').slideDown('slow', function() {
            //buildRightSideBottom('<%'iRateDesc%>', '<%'iPeriodID%>');
            buildRightSideBottom('<%=iPeriodID%>');
         });
      });
   }
}

function updateRenewalInfo() {
   var lcl_membership_alt_layout = '<%=lcl_membership_alt_layout%>';

   $('#membershipRenewals').slideUp('slow',function() {
      $('#membershipRenewals').html('');

      var lcl_userid   = $('#userid').val();
      var lcl_periodid = '';

      if(lcl_membership_alt_layout == 'Y') {
         $('input[name^="periodid"]').each(function(index) {
            if($(this).prop('checked')) {
               lcl_periodid = $(this).prop('id');
               lcl_periodid = lcl_periodid.replace('periodid_','');
               lcl_periodid = lcl_periodid.replace('periodid','');
            }
         });
      }
      else
      {
         lcl_periodid = $('#periodid option:selected').val();
      }
//alert('buildMembershipPurchaseSections.asp?orgid=<%=session("orgid")%>&purchaseSection=show_renewal_info&membershipid=<%=iMembershipID%>&periodid=' + lcl_periodid + '&userid=' + lcl_userid);
      if(lcl_periodid != '') {
         $.post('buildMembershipPurchaseSections.asp', {
            purchaseSection: 'show_renewal_info',
            orgid:           '<%=session("orgid")%>',
            adminuserid:     '<%=session("userid")%>',
            userid:          lcl_userid,
            membershipid:    '<%=iMembershipId%>',
            periodid:        lcl_periodid
         }, function(result) {
            $('#membershipRenewals').html(result);
            $('#membershipRenewals').slideDown('slow');
         });
      }
   });
}

function hideRightSidePicks() {
  $('#rightsidepicks_bottom').slideUp('slow', function() {
     $('#rightsidepicks_top').slideUp('slow', function() {
        $('#rightsidepicks_top').html('');
        $('#rightsidepicks_bottom').html('');
     });
  });

  $('#continueButton').prop('disabled',true);
}

//	function UserPick() {
// 		document.getElementById('searchname').value        = '';
//	 	document.getElementById('results').value           = '';
//		 document.getElementById('searchresults').innerHTML = '';
// 		document.getElementById('searchStart').value       = -1;
<%
'  if lcl_orghasfeature_purchase_membership_alt_layout then
'     response.write "document.getElementById('rateDesc').value = '';" & vbcrlf
'  end if
%>
//	 	document.getElementById('purchaseMembership').submit();
//	}

	function ContinuePurchase() {
			var submitOK = true;
			if ($("input[type='radio'][name='rateid']").length)
			{
				if (!$("input[name='rateid']:checked").val()) {
   					alert('You haven\'t selected a rate.');
					submitOK = false;
				}
			}
			
			if (submitOK)
			{
   				document.getElementById('purchaseMembership').action = 'select_members.asp';
   				document.getElementById('purchaseMembership').submit();
			}
	}

//	function EditUser(iUserId) {
//		 location.href='../dirs/update_citizen.asp?userid=' + iUserId;
//	}

	function NewUser()	{
 		location.href='../dirs/register_citizen.asp';
	}

function renewPass(iPassId) {
//  if (confirm("Delete Pass #" + iPassId + "?")) {
   			location.href='select_members.asp?poolpassid=' + iPassId;
//		}
//  inlineMsg(document.getElementById("button_renew_"+iPassId).id,'<strong>Coming Soon: </strong>Renew Membership Option for PoolPassID: '+iPassId,8,'button_renew_'+iPassId);
}

function updateRateID(iRateID) {
  document.getElementById('rateid').value = iRateID;
}

function updateMembershipOptions() {
   //var lcl_ratedesc = $('#rateDesc').val();
   var lcl_periodid = $('#periodid').val();

   //buildRightSideBottom(lcl_ratedesc, lcl_periodid);
   buildRightSideBottom(lcl_periodid);
<%
  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "enableDisableContinueButton('" & iPeriodID & "');" & vbcrlf
  else
     response.write "enableDisableContinueButton('" & request("rateid") & "');" & vbcrlf
  end if
%>

}

//function submitPoolPassForm() {
//  document.getElementById('purchaseMembership').action = 'poolpass_form.asp';
//  document.getElementById('purchaseMembership').submit();
//}

function enableDisableContinueButton(iRowValue) {
<%
  if lcl_orghasfeature_purchase_membership_alt_layout then
     lcl_radio_id          = "periodid_"
     'lcl_onload = lcl_onload & "enableDisableContinueButton('" & iPeriodID & "');"
  else
     lcl_radio_id          = "rateid_"
     'lcl_onload = lcl_onload & "enableDisableContinueButton('" & request("rateid") & "');"
  end if

  response.write "  if(iRowValue=='') {" & vbcrlf
  response.write "     document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "  }else{" & vbcrlf
  response.write "     if(document.getElementById('" & lcl_radio_id & "'+iRowValue)) {"& vbcrlf
  response.write "        if(document.getElementById('" & lcl_radio_id & "'+iRowValue).checked==true) {" & vbcrlf
  response.write "           document.getElementById('continueButton').disabled=false;" & vbcrlf
  response.write "        }else{" & vbcrlf
  response.write "           document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "        }" & vbcrlf
  response.write "     }else{" & vbcrlf
  response.write "        document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "     }" & vbcrlf
  response.write "  }" & vbcrlf
%>
}
 //-->
 </script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<body onload="<%=lcl_onload%>">
<%
  response.write "<form name=""purchaseMembership"" id=""purchaseMembership"" method=""post"" action=""poolpass_form.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""results"" id=""results"" value="""" />" & vbcrlf
  'response.write "  <input type=""hidden"" name=""searchStart"" id=""searchStart"" value="""     & sSearchStart   & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""isSeasonal"" id=""isSeasonal"" value="""       & lcl_isSeasonal & """ />" & vbcrlf
  'response.write "  <input type=""hidden"" name=""usertype"" id=""usertype"" value="""           & sUserType      & """ />" & vbcrlf
  'response.write "  <input type=""hidden"" name=""iuserid"" id=""iuserid"" value="""             & iUserId        & """ />" & vbcrlf

  'response.write "  <input type=""hidden"" name=""iMembershipId"" id=""iMembershipId"" value=""" & iMembershipId  & """ />" & vbcrlf


  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "  <input type=""hidden"" name=""rateid"" id=""rateid"" value=""" & request("rateid") & """ />" & vbcrlf
  end if

  response.write "<div id=""content"">" & vbcrlf

  response.write "  <div id=""poolcentercontent"">" & vbcrlf
  response.write "    <table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "      <tr><th colspan=""2"" align=""left"">Purchase a " & session("sOrgName") & "&nbsp;Membership</th></tr>" & vbcrlf
  response.write "    	 <tr>" & vbcrlf
  response.write "          <td>" & vbcrlf
  response.write "  Membership Type: <select name=""iMembershipId"" id=""iMembershipId"" onChange=""membershipTypeChange();"">" & vbcrlf
  showMembershipTypePicks 0
  response.write "</select><br /><br />" & vbcrlf
  response.write "              Select a registered Citizen from the drop down list, select the pass they want and then press the " & vbcrlf
  response.write "              <strong>Continue with Purchase</strong> button.  If their name is not on the list then select " & vbcrlf
  response.write "              <strong>New User</strong> to add them to the list. If their information is incorrect then select " & vbcrlf
  response.write "              <strong>Edit User Profile</strong>." & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "	    	<tr>" & vbcrlf
  response.write "          <td nowrap=""nowrap"">" & vbcrlf
  response.write "              <fieldset id=""nameSearch"" class=""fieldset"">" & vbcrlf
  response.write "                <legend>Name Search</legend>" & vbcrlf
  response.write "                <input type=""text"" name=""searchname"" id=""searchname"" value="""" maxlength=""50"" onkeypress=""if(event.keyCode=='13'){$('#searchButton').click();return false;}"" />" & vbcrlf
  response.write "                <input type=""button"" name=""searchButton"" id=""searchButton"" class=""button"" value=""Search"" />" & vbcrlf
  response.write "       		    <br /><div id=""searchtip"">(last name, first name)</div>" & vbcrlf
  response.write "                <div id=""selectName""></div>" & vbcrlf
  response.write "   		     </fieldset>" & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    		<tr>" & vbcrlf
  response.write "          <td nowrap=""nowrap"">" & vbcrlf
  response.write "              <input type=""button"" name=""newUserButton"" id=""newUserButton"" class=""button"" onclick=""NewUser();"" value=""New User"" />&nbsp;&nbsp;" & vbcrlf
  response.write "              <input type=""button"" name=""editUserButton"" id=""editUserButton"" class=""button"" value=""Edit User Profile"" />" & vbcrlf
  'response.write "              <input type=""button"" name=""editUserButton"" id=""editUserButton"" class=""button"" onclick=""EditUser(" & iUserId & ");"" value=""Edit User Profile"" />" & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td>" & vbcrlf
  response.write "              <div id=""userInfoDiv""></div>" & vbcrlf
  response.write "              <div id=""membershipRenewals""></div>" & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf

 'Membership Renewals
'  response.write "    <div id=""membershipRenewals"">renewals here - left off here!!!</div>" & vbcrlf

  response.write "  </div>" & vbcrlf

 'Check to see which layout the org has turned-on
  'response.write "  <div id=""purchasePicks""></div>" & vbcrlf
  response.write "  <div id=""rightsidepicks"">" & vbcrlf
  response.write "    <div id=""rightsidepicks_top""></div>" & vbcrlf
  response.write "    <div id=""rightsidepicks_bottom""></div>" & vbcrlf

  'Build the "continue" button
  'Determine if "Continue with Purchase" button is displayed.
  'Is the citizen selected is a "non-resident"?
  'If "yes" then check to see if the org has the "nonresidentlimit" feature assigned to it.
  'If "yes" then total the number records, for non-residents, on egov_poolpasspurchases for the CURRENT "membership period" selected.
  'If the total records on egov_poolpasspurchases is LESS THAN the "Non-Resident Cap" set 
  response.write "    <div id=""rightsidepicks_continuebutton"">" & vbcrlf
  response.write "      <input type=""button"" name=""continue"" id=""continueButton"" class=""button"" value=""Continue with Purchase"" onclick=""ContinuePurchase();"" />" & vbcrlf
  response.write "    </div>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf

 'BEGIN: Renewal Memberships --------------------------------------------------


  'if lcl_membershiprenewals_feature = "Y" then

    'Show the renewal membership record if one exists for the user selected
     'showRenewalMembership iUserID, _
     '                      iMembershipID, _
     '                      iPeriodID
  'end if
 'END: Renewal Memberships ----------------------------------------------------

'  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

  set oMembership = nothing 

'------------------------------------------------------------------------------
sub showMembershipTerms(iOrgID, iMembershipID)

  dim sOrgID, sMembershipID, sTermsRequired

  sOrgID         = 0
  sMembershipID  = 0
  sTermsRequired = "N"

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iMembershipID <> "" then
     sMembershipID = clng(iMembershipID)
  end if

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf

  sSQL = "SELECT showterms, "
  sSQL = sSQL & " membershipterms "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND membershipid = " & sMembershipID
  sSQL = sSQL & " AND showterms = 1 "

 	set oMembershipTerms = Server.CreateObject("ADODB.Recordset")
	 oMembershipTerms.Open sSQL, Application("DSN"), 3, 1

  if not oMembershipTerms.eof then

     lcl_showterms = oMembershipTerms("showTerms")

     if lcl_showTerms then
        lcl_membershipterms = oMembershipTerms("membershipterms")
        sTermsRequired      = "Y"

        response.write "          <fieldset id=""membershipTerms"" class=""fieldset"">" & vbcrlf
        response.write "            <legend>Membership Terms&nbsp;</legend>" & vbcrlf
        response.write "            <p>" & lcl_membershipterms & "</p>" & vbcrlf
        response.write "            <p><input type=""checkbox"" name=""termsRead"" id=""termsRead"" value=""Y"" /> I have read terms</p>" & vbcrlf
        response.write "          </fieldset>" & vbcrlf
     end if
  end if

  oMembershipTerms.close
  set oMembershipTerms = nothing

  response.write "          <input type=""text"" name=""termsRequiredToContinue"" id=""termsRequiredToContinue"" value=""" & sTermsRequired & """ size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

end sub

'------------------------------------------------------------------------------
 function checkIsSeasonal(p_periodid)
   lcl_return = False

   if p_periodid <> "" then
      sSQL = "SELECT is_seasonal "
      sSQL = sSQL & " FROM egov_membership_periods "
      sSQL = sSQL & " WHERE orgid = " & session("orgid")
      sSQL = sSQL & " AND periodid = " & p_periodid

     	set oPeriod = Server.CreateObject("ADODB.Recordset")
    	 oPeriod.Open sSQL, Application("DSN"), 3, 1

      if not oPeriod.eof then
         if oPeriod("is_seasonal") = "" then
            lcl_return = False
         else
            lcl_return = oPeriod("is_seasonal")
         end if
      end if

      set oPeriod = nothing

   end if

   checkIsSeasonal = lcl_return

 end function

sub showMembershipTypePicks(p_membership_type)

  sSQL = "SELECT membershipid, membership, membershipdesc "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY membershipdesc "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     while not rs.eof

        response.write "  <option value=""" & rs("membershipid") & """>" & rs("membershipdesc") & "</option>" & vbcrlf
        rs.movenext
     wend
  end if

  set rs = nothing

end sub
%>
