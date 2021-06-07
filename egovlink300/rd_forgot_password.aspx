<%@ Page Language="C#" AutoEventWireup="true" CodeFile="rd_forgot_password.aspx.cs" Inherits="rd_forgot_password" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
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
<head runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="rd_global.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"css/style_" + sOrgID + ".css\" />"%>
    
  <script type="text/javascript" src="scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="rd_scripts/egov_navigation.js"></script>
  
<script type="text/javascript">
    $(document).ready(function() {
        $('#email').focus();

        $('#lookupButton').click(function() {
            var lcl_orgid = '<%=sOrgID%>';
            var lcl_email = $('#email').val();

            $('#forgotpwd_error').html('');

            $.post('rd_forgot_password_action.aspx', {
                orgid: lcl_orgid,
                email: lcl_email
            }, function(result) {
                if (result == 'SENT') {
                    $('#forgotpwd_error').html('<span>Password reset instructions have been sent to you.</span>');
                    $('#email').prop('disabled', 'true');
                    $('#lookupButton').prop('disabled', 'true');
                } else {
                    $('#forgotpwd_error').html('<span>The email address you entered does not exist.</span>');
                    $('#email').focus();
                }
            });
        });

        $('#loginButton').click(function() {
            location.href = 'rd_user_login.aspx';
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
  displayForgotPassword();
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
