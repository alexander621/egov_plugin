<%@ Page Language="C#" AutoEventWireup="true" CodeFile="rd_user_login.aspx.cs" Inherits="rd_classes_rd_user_login" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID       = common.getOrgId();
    static string sOrgName     = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle          = "E-Gov Services " + sOrgName;
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
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="rd_global.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"css/style_" + sOrgID + ".css\" />"%>
    
  <script type="text/javascript" src="scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="rd_scripts/egov_navigation.js"></script>
  
  <script type="text/javascript">
      $(document).ready(function() {

          if ($('#email').prop('disabled') == false) {
              $('#email').focus();
          }

          $('#signInButton').click(function() {
              var lcl_false_count = 0;

              $('#signInButton').prop('disabled', true);
              $('#email').prop('disabled', true);
              $('#password').prop('disabled', true);

              $('#loginErrorMsg').html('<div id="loginMsgProcessing">Processing...</div>');

              /*
              if ($('#problemtextinput').val() != '') {
              $('#loginErrorMsg').html('<div id="loginMsg"><strong>Invalid Value:</strong> Please remove any input from the Internal Only field at the bottom of the form.</div>');

                  $('#signInButton').prop('disabled', false);
              $('#email').prop('disabled', false);
              $('#password').prop('disabled', false);
              //$('#problemtextinput').focus();
              //inlineMsg(document.getElementById('problemtextinput').id, '<strong>Invalid Value: </strong> Please remove any input from the Internal Only field at the bottom of the form.', 10, 'problemtextinput');
              lcl_false_count = lcl_false_count + 1;
              }
              */

              if (lcl_false_count > 0) {
                  return false;
              } else {
                  var lcl_orgid = $('#orgid').val();
                  var lcl_email = $('#email').val();
                  var lcl_pwd = $('#password').val();
                  var lcl_fst = $('#problemtextinput').val();

                  $.post('rd_user_login_action.aspx', {
                      orgid: lcl_orgid,
                      email: lcl_email,
                      password: lcl_pwd,
                      frmsubjecttext: lcl_fst
                  }, function(result) {
                      var lcl_result = result;

                      if (lcl_result.indexOf('FAILED') > -1) {
                          var lcl_failed_msg = 'The email and/or password entered are incorrect.';
                          $('#loginErrorMsg').html('<div id="loginMsg">' + lcl_failed_msg + '</div>');

                          $('#signInButton').prop('disabled', false);
                          $('#email').prop('disabled', false);
                          $('#password').prop('disabled', false);
                          $('#email').focus();
                      } else if (lcl_result.indexOf('NEWUSER') > -1) {
                          var lcl_email = $('#email').val();

                          var lcl_url = 'register.asp';
                          lcl_url += '?egov_users_useremail=' + lcl_email;

                          location.href = lcl_url;
                      } else if (lcl_result.indexOf('REDIRECT') > -1) {
                          var lcl_userid = "";
                          var lcl_resultStart = 0;

                          lcl_resultStart = lcl_result.indexOf('REDIRECT');
                          lcl_userid = lcl_result.substr(0, lcl_resultStart);

                          lcl_result = lcl_result.substr(lcl_resultStart);
                          lcl_result = lcl_result.replace('REDIRECT', '');

                          location.href = lcl_result;
                      }
                  });
              }
          });
      });
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
  displayUserLogin(Convert.ToInt32(sOrgID));
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
