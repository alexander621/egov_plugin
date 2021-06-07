<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_cart.aspx.cs" Inherits="rd_classes_class_cart" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Configuration" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    //HttpCookie sCookieUserID;

    static Int32 sUserID = 0;
    static Int32 sClassID = 0;
    static Int32 sCategoryID = 0;

    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    static string sSessionID         = "";
    static string sCategoryTitle = "";
    //static string sSessionIDName     = "";  //This is used to identify the column to save the session value to on "egov_aspnet_to_asp_usersessions"
    
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
    string lcl_isLoggedIn            = "";
    string lcl_checked_isLoggedInYes = "";
    string lcl_checked_isLoggedInNo  = "";
    
    static Int32 iRootCategoryID = classes.getFirstCategory(sOrgID);
    //Int32 sCategoryID            = iRootCategoryID;
</script>
<%
    if (sOrgID.ToString() == "7")
    {
        sPageTitle = sOrgName;
    }

    try
    {
        sUserID = Convert.ToInt32(Request.QueryString["userid"]);
    }
    catch
    {
        sUserID = 0;
    }

    try
    {
        sClassID = Convert.ToInt32(Request.QueryString["iClassID"]);
    }
    catch
    {
        sClassID = 0;
    }
    
    try
    {
        sCategoryID = Convert.ToInt32(Request.QueryString["categoryID"]);
    }
    catch
    {
        sCategoryID = iRootCategoryID;
    }
    sCategoryTitle = Request["categorytitle"];
    
    //Set up variables for common user controls
    egov_navigation.egovsection  = "CLASSES_NOSEARCH";
    egov_navigation.rootcategory = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid   = Convert.ToString(sCategoryID);

    //Setup User and Session Variables
    //sCookieUserID = Request.Cookies["useridx"];
    
    sSessionID = HttpContext.Current.Session.SessionID;

    Session["RedirectPage"]    = "rd_classes/class_categories.aspx";
    Session["RedirectLang"]    = "Return to Class Categories";
    Session["LoginDisplayMsg"] = "";
    Session["DisplayMsg"]      = "";
    Session["ManageURL"]       = "";
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
  
  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../rd_global.css" />
  <link type="text/css" rel="stylesheet" href="styles_class.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>
    
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/formatnumber.js"></script>
  <%="<script type=\"text/javascript\" src=\"" + common.getBaseURL("") + "/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>

<script type="text/javascript">
    $(document).ready(function() {
        $('#classSignUpButton').click(function() {
            eGovLink.Class.goToSignup();
        });

        $('#classListButton').click(function() {
            location.href = 'class_list.aspx';
        });

        $('#completeButton').click(function() {
            eGovLink.Class.checkPurchases();
        });
    });

    var eGovLink = eGovLink || {};

    eGovLink.Class = (function() {
        var toggleApplyCredit = function() {

            var toggleTo = "1";
            var isChecked = $('#checkapplycredit:checked').val() ? true : false;

            if (isChecked) {
                $('#applycredit').val('yes');
                $('#columnApplyCreditAmount').html(format_number($('#accountcredit').val(), 2));

                var totalDue = parseFloat($('#nTotal').val()) - parseFloat($('#accountcredit').val());

                $('#totalDueDisplay').html(format_number(totalDue, 2));
                $('#totaldue').val(totalDue);
            } else {
                $('#applycredit').val('no');
                $('#columnApplyCreditAmount').html(format_number(0.00, 2));
                $('#totaldue').val($('#nTotal').val());
                $('#totalDueDisplay').html(format_number($('#nTotal').val(), 2));
            }
        };

        var removeItem = function(iCartID, iTimeID, sBuyOrWait) {
            if (confirm('Remove this item from the cart?')) {
                var lcl_url = 'class_remove.aspx';
                lcl_url += '?cartid=' + iCartID;
                lcl_url += '&timeid=' + iTimeID;
                lcl_url += '&buyorwait=' + sBuyOrWait;

                window.location.href = lcl_url;
            }
        };

        var purchaseCart = function() {
            var lcl_url = '<%=ConfigurationManager.AppSettings["paymenturl"] %>';
            lcl_url += '/';
            lcl_url += '<%=sOrgVirtualSiteName %>';
            lcl_url += '/rd_classes/verisign_form.aspx';

            location.href = lcl_url;
        };

        var checkPurchases = function() {
            //See if the total charges are a number of if the apply credit is yes
            //and the balance due is 0 then redirect them
            if (parseFloat($('#nTotal').val()) > parseFloat(0.00) && $('#applycredit').val() == 'yes' && parseFloat($('#nTotal').val()) == parseFloat($('#accountcredit').val()) && parseFloat($('#totaldue').val()) == parseFloat(0.00)) {
                //Change the form's ACTION send the user to the process_payment.aspx page to simply
                //insert the required e-gov table records.
                var lcl_action = '';
                var lcl_action = '<%=ConfigurationManager.AppSettings["paymenturl"] %>';
                lcl_action += '/';
                lcl_action += '<%=sOrgVirtualSiteName %>';
                lcl_action += '/rd_classes/ProcessPayment.aspx';
                //lcl_action = 'ProcessPayment.aspx';
                lcl_action += '?itemnumber=<%=sSessionID %>';
                lcl_action += '&userid=<%=sUserID %>';
                lcl_action += '&applycredit=yes';

                document.cartForm.action = lcl_action;
                document.cartForm.submit();
            } else {
                //Check if any items in the cart are for purchase via AJAX
                //doAjax('check_cart_purchases.aspx', 'sessionid=<%=sSessionID %>', 'eGovLink.Class.PurchaseCheckReturn', 'get', '0');

                $.post('check_cart_purchases.aspx', {
                    sessionid: '<%=sSessionID%>'
                }, function(result) {
                    eGovLink.Class.purchaseCheckReturn(result);
                    //document.cartForm.submit();
                });
            }
        };

        var purchaseCheckReturn = function(sReturn) {
            if (sReturn == 'PURCHASES' || sReturn == 'WAITLISTONLY') {
                lcl_total = parseFloat($('#nTotal').val());

                //Check the total to determine how to process the cart
                if (lcl_total == parseFloat(0.00)) {
                    //Change the form's ACTION send the user to the process_payment.aspx page
                    //to simply insert the require e-gov table records.
                    var lcl_action = '<%=ConfigurationManager.AppSettings["paymenturl"] %>';
                    lcl_action += '/';
                    lcl_action += '<%=sOrgVirtualSiteName %>';
                    lcl_action += '/rd_classes/ProcessPayment.aspx';
                    lcl_action += '?itemnumber=<%=sSessionID %>';
                    lcl_action += '&userid=<%=sUserID %>';
                    lcl_action += '&applycredit=no';

                    document.cartForm.action = lcl_action;
                }

                //Send them on to Payment Page (class_paymentform.aspx)
                document.cartForm.submit();
            } else {
                //if (sReturn == 'WAITLISTONLY') {
                    //send them to waitlist only processing

                    //alert('finish');
                    //location.href = 'process_to_waitlist.aspx';
                //}
                //else {
                    //Something is wrong, so re-post this page.  Maybe the session timed off
                    location.href = 'class_cart.aspx';
                //}
            }
        };

        var goToSignup = function() {
            var lcl_url = 'class_signup.aspx';
            lcl_url += '?classid=<%=sClassID.ToString() %>';
            lcl_url += '&categoryid=<%=sCategoryID.ToString() %>';
            lcl_url += '&categorytitle=' + encodeURI('<%=sCategoryTitle %>');

            location.href = lcl_url;
        }

        /*
        var removeMerchandiseItem = function(iCartID) {
        if (confirm('Remove this merchandise purchase from the cart?')) {
        var lcl_url = '../merchandise/merchandiseremove.aspx';
        lcl_url += '?cartid=' + iCartID;

                location.href = lcl_url;
        }
        };
        
        var editTeam = function(iCartID, iClassID) {
        var lcl_url = 'regattateamsignup.asp';
        lcl_url += '?classid=' + iClassID;
        lcl_url += '&cartid=' + iCartID;

            location.href = lcl_url;
        }

        var editMembers = function(iCartID, iClassID) {
        var lcl_url = 'regattamembersignup.asp';
        lcl_url += '?classid=' + iClassID;
        lcl_url += '&cartid=' + iCartID;

            location.href = lcl_url;
        }

        var editMerchandise = function(iCartID) {
        var lcl_url = '../merchandise/merchandiseofferings.asp';
        lcl_url += '?cartid=' + iCartID;

            location.href = lcl_url;
        }
        */

        //This makes the functions publically accessible
        return {
            toggleApplyCredit: toggleApplyCredit,
            removeItem: removeItem,
            purchaseCart: purchaseCart,
            checkPurchases: checkPurchases,
            purchaseCheckReturn: purchaseCheckReturn,
            goToSignup: goToSignup
            //removeMerchandiseItem: removeMerchandiseItem,
            //editTeam: editTeam,
            //editMembers: editMembers,
            //editMerchandise: editMerchandise
        };
    } ());
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
    displayCart(Convert.ToInt32(sOrgID),
                sSessionID);
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
