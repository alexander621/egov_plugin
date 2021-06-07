<%@ Page Language="C#" AutoEventWireup="true" CodeFile="rd_register.aspx.cs" Inherits="rd_register" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID          = common.getOrgId();
    static string sOrgName        = common.getOrgName(sOrgID);
    static string lcl_addresstype = "";
    string sOrgVirtualSiteName    = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle             = "E-Gov Services " + sOrgName;

    static Boolean sOrgHasFeatureLargeAddressList = common.orgHasFeature(sOrgID, "large address list");
</script>
<%
    if (sOrgID.ToString() == "7")
    {
        sPageTitle = sOrgName;
    }
    
    //Set up variables for common user controls
    egov_navigation.egovsection  = "HIDE_SUBMENU";
    egov_navigation.rootcategory = "";
    egov_navigation.categoryid   = "";

    if (sOrgHasFeatureLargeAddressList)
    {
        lcl_addresstype = "LARGE";
    }
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="rd_global.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"css/style_" + sOrgID + ".css\" />"%>

  <script type="text/javascript" src="scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"http://www.egovlink.com/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="rd_scripts/egov_navigation.js"></script>
  
  <script type="text/javascript">
      $(document).ready(function() {
          if ('<%=sOrgHasFeatureLargeAddressList %>' == 'True') {
              $('#validAddressList').hide();
          }

          //enableDisableAddressFields('');

          $('#residentstreetnumber').focus(function() {
              $('#residentstreetnumber').removeClass('inputFieldHighlight');
          });

          $('#streetaddress').focus(function() {
              $('#streetaddress').removeClass('inputFieldHighlight');
          });
      });

      //BEGIN: Enable/Disable Address Fields ----------------------------------------------
      function enableDisableAddressFields(iMode) {
          var lcl_mode = '';

          if (iMode != '') {
              lcl_mode = iMode;
          }

          if (lcl_mode == 'disabled') {
              $('#residentstreetnumber').prop('disabled', 'disabled');
              $('#streetaddress').prop('disabled', 'disabled');
              $('#egov_users_useraddress').prop('disabled', 'disabled');
              $('#registerValidateAddressButton').prop('disabled', 'disabled');
          } else {
              $('#residentstreetnumber').prop('disabled', '');
              $('#streetaddress').prop('disabled', '');
              $('#egov_users_useraddress').prop('disabled', '');

              if ($('#egov_users_useraddress').val() != '') {
                  $('#registerValidateAddressButton').prop('disabled', 'disabled');
              } else {
                  if ($('#residentstreetnumber').val() != '') {
                      $('#registerValidateAddressButton').prop('disabled', '');
                  } else {
                      if ($('#streetaddress').val() == '0000') {
                          $('#registerValidateAddressButton').prop('disabled', 'disabled');
                      }
                  }
              }
          }
      }
      //END: Enable/Disable Address Fields ------------------------------------------------
      
      //BEGIN: Check Address --------------------------------------------------------------
      function checkAddress(iFunction, iValidate) {
          //if('<%=sOrgHasFeatureLargeAddressList.ToString() %>' == 'True') {
          var lcl_streetnumber = $('#residentstreetnumber').val();
          var lcl_streetname   = $('#streetaddress').val();
          var lcl_otheraddress = $('#egov_users_useraddress').val();
          var lcl_isFinalCheck = 'N';
          var lcl_success      = false;
          
          clearScreenMsg();

          if (iFunction == 'FinalCheck') {
              lcl_isFinalCheck = 'Y';
          }

          $('#isFinalCheck').val(lcl_isFinalCheck);

          //Validate the street number and name entered to determine if it is a valid address in the system for the org
          if (lcl_otheraddress == '') {
              var lcl_success = validateAddress();

              if (lcl_success) {
                  $.post('rd_checkaddress.aspx', {
                      addresstype: '<%=lcl_addresstype%>',
                      orgid: '<%=sOrgID %>',
                      stnumber: lcl_streetnumber,
                      stname: lcl_streetname,
                      returntype: 'CHECK'
                  }, function(result) {
                      displayValidAddressList(result);
                  });
              } else {
                  if (lcl_isFinalCheck == 'Y') {
                      //if(ValidateInput()) {
                      //    isemailentered();
                      //}
                      alert('ValidateInput + isemailentered 1');
                  }
              }
          } else {
              if (lcl_streetnumber != '' || lcl_streetname != '0000') {
                  lcl_success = validateAddress();

                  if (!lcl_success) {
                      FinalCheck('NOT FOUND', 1);
                  }
              } else {
                  if (lcl_isFinalCheck == 'Y') {
                      //if(ValidateInput()) {
                      //    isemailentered();
                      //}
                      alert('ValidateInput + isemailentered 2');
                  }
              }
          }
          //}
      }
      //END: Check Address ----------------------------------------------------------------
      
      //BEGIN: Validate Address -----------------------------------------------------------
      function validateAddress() {
          clearScreenMsg('residentstreetnumber');
          clearScreenMsg('streetaddress');
          clearScreenMsg('registerValidateAddressButton');
          
          //Remove any extra spaces
          $('#residentstreetnumber').val(jQuery.trim($('#residentstreetnumber').val()));
          
          //Check for non-numeric values
          if($('#residentstreetnumber').val() != '') {
              var rege = /^\d+$/;
              var Ok = rege.exec($('#residentstreetnumber').val());
              
              if(!Ok) {
                  $('#residentstreetnumber').focus();
                  inlineMsg(document.getElementById('residentstreetnumber').id,'<strong>Invalid Value: </strong> The Street Number must be numeric.',10,'residentstreetnumber');
                  return false;
              } else {
                  //Check that they picked a street name
                  if($('#streetaddress').val() == '0000') {
                      $('#streetaddress').focus();
                      inlineMsg(document.getElementById('streetaddress').id,'<strong>Required Field: </strong> Please select a street name from the list before validating the address.',10,'streetaddress');
                      return false;
                  } else {
                      return true;
                  }    
              }
          } else {
              if($('#streetaddress').val() == '0000') {
                  if('<%=sOrgHasFeatureLargeAddressList.ToString() %>' == 'True') {
                      inlineMsg(document.getElementById('registerValidateAddressButton').id, '<strong>Required Field: </strong> At least one address field must be entered before attempting to validate.', 10, 'registerValidateAddressButton');
                  } else {
                      inlineMsg(document.getElementById('streetaddress').id,'<strong>Required Field: </strong> An address must be entered before attempting to validate.',10,'streetaddress');
                  }
                  
                  return false;
              } else {
                  $('#residentstreetnumber').focus();
                  inlineMsg(document.getElementById('residentstreetnumber').id,'<strong>Required Field: </strong> Street Number',10,'residentstreetnumber');
                  return false;
              }
          }
      }
      //END: Validate Address -------------------------------------------------------------

      //BEGIN: Final Check ----------------------------------------------------------------
      function FinalCheck(sResults, iFalseCount) {
          var lcl_isFinalCheck = $('#isFinalCheck').val();

          if (sResults == 'FOUND CHECK') {
              //$('#validstreet').val('Y');
              $('#validAddressList').hide('slow');
              enableDisableAddressFields('');

              if (lcl_isFinalCheck == 'Y') {
                  //if (ValidateInput()) {
                  //    isemailentered();
                  //}
                  alert('ValidateInput - isemailentered');
              }
          } else if (sResults == 'SUBMIT') {
              if ($('#egov_users_useraddress').val() == '') {
                  var lcl_streetnumber = $('#residentstreetnumber').val();
                  var lcl_streetname   = $('#streetaddress').val();
              }

              if (iFalseCount > 0) {
                  return false;
              } else {
                  $('#register').submit();
                  return true;
              }
          } else {
              if ((sResults == 'FOUND SELECT') || (sResults == 'FOUND KEEP')) {
                  //if (sResults == 'FOUND SELECT') {
                  //    $('#validstreet').val('Y');
                  //} else {
                  //    $('#validstreet').val('N');
                  //}

                  $('#validAddressList').hide('slow');
                  enableDisableAddressFields('');

                  if (lcl_isFinalCheck == 'Y') {
                      //if (ValidateInput()) {
                      //    isemailentered();
                      //}
                      alert('ValidateInput - isemailentered');
                  }
              } else {
                  if ($('#egov_users_useraddress').val() != '') {
                      $('#validateAddressList').hide('slow');
                      enableDisableAddressFields('');

                      if (lcl_isFinalCheck == 'Y') {
                          //if (ValidateInput()) {
                          //    isemailentered();
                          //}
                          alert('ValidateInput - isemailentered');
                      }
                  } else {
                      $('#validAddressList').show('slow');
                      enableDisableAddressFields('disabled');
                  }
              }
          }
      }
      //END: Final Check ------------------------------------------------------------------

      //BEGIN: Display Valid Address List -------------------------------------------------
      function displayValidAddressList(iResult) {
          var lcl_streetnumber = $('#residentstreetnumber').val();
          var lcl_streetname   = $('#streetaddress').val();
          var lcl_isFinalCheck = $('#isFinalCheck').val();

          //Determine if the address is "valid" based on records in egov_residentaddresses for the org
          if (iResult == 'FOUND CHECK' || iResult == 'CANCEL') {
              if (iResult == 'FOUND CHECK') {
                  displayScreenMsg('Address is Valid');
                  //$('#validstreet').val('Y');
              }

              $('#validAddressList').hide('slow');

              enableDisableAddressFields('');

              if (iResult != 'CANCEL' && lcl_isFinalCheck == 'Y') {
                  //if (ValidateInput()) {
                  //    isemailentered();
                  //}
                  alert('ValidateInput - isemailentered');
              }
          } else {
              //displayScreenMsg('Invalid Address');
              //$('#validStreet').val('N');
              $('#oldstnumber').val(lcl_streetnumber);
              $('#stname').val(lcl_streetname);

              enableDisableAddressFields('disabled');

              $('#validAddressList').show('slow', function() {
                  var lcl_addressEntered  = $('#residentstreetnumber').val();
                      lcl_addressEntered += ' ';
                      lcl_addressEntered += $('#streetaddress').val();

                  $('#registerDisplayAddressEntered').html(lcl_addressEntered);

                  $.post('rd_checkaddress.aspx', {
                      addresstype: '<%=lcl_addresstype%>',
                      orgid: '<%=sOrgID %>',
                      stnumber: lcl_streetnumber,
                      stname: lcl_streetname,
                      returntype: 'DISPLAY_OPTIONS'
                  }, function(result) {
                      $('#addresspicklist').html(result);
                  });
              });
          }
      }
      //END: Display Valid Address List ---------------------------------------------------

      //BEGIN: Do Select ------------------------------------------------------------------
      function doSelect() {
          if ($('#stnumber').prop('selectedIndex') < 0) {
              inlineMsg(document.getElementById('stnumber').id, '<strong>Required Field Missing: </strong> Please select a valid address first.', 10, 'stnumber');
              return false;
          }

          clearScreenMsg();
          clearMsg('stnumber');
          $('#residentstreetnumber').val($('#stnumber').val());
          $('#egov_users_useraddress').val('');
          FinalCheck('FOUND SELECT', 0);
      }
      //END: Do Select --------------------------------------------------------------------

      //BEGIN: Cancel Pick ----------------------------------------------------------------
      function cancelPick() {
          clearScreenMsg();
          clearMsg('stnumber');
          displayValidAddressList('CANCEL');
      }
      //END: Cancel Pick ------------------------------------------------------------------

      //BEGIN: Do Keep --------------------------------------------------------------------
      function doKeep() {
          var lcl_streetnumber  = $('#oldstnumber').val();
          var lcl_streetname    = $('#stname').val();
          var lcl_streetaddress = '';

          if (lcl_streetname != '') {
              lcl_streetaddress = lcl_streetnumber;
          }

          if (lcl_streetname != '') {
              if (lcl_streetaddress != '') {
                  lcl_streetaddress += ' ';
                  lcl_streetaddress += lcl_streetname;
              } else {
                  lcl_streetaddress = lcl_streetname;
              }
          }

          FinalCheck('FOUND KEEP', 0);
          $('#egov_users_useraddress').val(lcl_streetaddress);
          $('#residentstreetnumber').val('');
          $('#streetaddress').val('');
          $('#streetaddress').prop('selectedIndex', 0);
      }
      //END: Do Keep ----------------------------------------------------------------------
      
      function displayScreenMsg(iMsg) {
          if (iMsg != '') {
              document.getElementById('registerErrorMsgDiv').innerHTML = '<div id="registerErrorMsg">' + iMsg + '</div>';
              window.setTimeout('clearScreenMsg()', (10 * 1000));
          }
      }

      function clearScreenMsg() {
          document.getElementById('registerErrorMsgDiv').innerHTML = '';
      }

      function cleanUpAddressFields(iField) {
          if (iField == 'egov_users_useraddress') {
              enableDisableAddressFields('');

              if ($('#egov_users_useraddress').val()) {
                  $('#residentstreetnumber').val('');
                  $('#streetaddress').val('0000');
              }
          } else {
              clearMsg(iField);
              clearMsg('registerValidateAddressButton');
              enableDisableAddressFields('');

              if (iField == 'streetaddress') {
                  if ($('#' + iField).val() != '0000') {
                      $('#egov_users_useraddress').val('');
                  }
              } else {
                  if ($('#' + iField).val() != '') {
                      $('#egov_users_useraddress').val('');
                  }
              }
          }
      }
  </script>
</head>
<body>
<div id="wrapper_body">
  <div id="wrapper_header">
    <Tbanner:banner ID="banner" runat="server" />
    <Tnavigation:navigation ID="egov_navigation" runat="server" egovsection="" rootcategory="" categoryid="" />
  </div>
  <div id="wrapper_content">
    <div id="content">
<%
  displayUserRegister(Convert.ToInt32(sOrgID));
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
