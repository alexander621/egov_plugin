<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_signup.aspx.cs" Inherits="rd_classes_class_signup" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<%@ Register TagPrefix="classes_memberWarning" TagName="classesMemberWarning" Src="../rd_includes/egov_classes_memberwarning.ascx" %>

<!DOCTYPE html>
<script runat="server">
    HttpCookie sCookieUserID;
    static string sOrgID         = common.getOrgId();
    static string sOrgName       = common.getOrgName(sOrgID);
    static string sSessionID     = "";
    static string sSessionIDName = "";  //This is used to identify the column to save the session value to on "egov_aspnet_to_asp_usersessions"
    string sOrgVirtualSiteName   = common.getOrgInfo(sOrgID, "orgVirtualSiteName");    
    string sPageTitle            = "E-Gov Services " + sOrgName;
    //string sCategoryTitle        = "";
    string sUserID               = "";

    static Int32 sUserSessionID  = 0;
    static Int32 iRootCategoryID = classes.getFirstCategory(sOrgID);
    Int32 sCategoryID            = iRootCategoryID;
    Int32 sClassID               = 0;

    static Boolean sSetupCoachFields     = false;
    static Boolean sIsUserMissingKeyData = false;
    static Boolean sIsUserMissingAddress = true;
    static Boolean sOrgHasFeature_registrationReqAddress    = common.orgHasFeature(sOrgID, "registration req address");
    static Boolean sOrgHasFeature_customRegistrationCraigCO = common.orgHasFeature(sOrgID, "custom_registration_craigco");
    static Boolean sOrgHasFeature_emergencyInfoRequired     = common.orgHasFeature(sOrgID, "emergency info required");
</script>
<%
    //Validate parameters being passed in.
    if (Request["categoryid"] != null)
    {
        try
        {
            sCategoryID = Convert.ToInt32(Request["categoryid"]);
        }
        catch
        {
            Response.Redirect("class_categories.aspx");
        }
    }
    else
    {
        Response.Redirect("class_categories.aspx");
    }
    
    if (Request["classid"] != null)
    {
        try
        {
            sClassID = Convert.ToInt32(Request["classid"]);
        }
        catch
        {
            Response.Redirect("class_categories.aspx");
        }
    }
    else
    {
        Response.Redirect("class_categories.aspx");
    }

    //if (Request["categorytitle"] != null)
    //{
    //    sCategoryTitle = Request["categorytitle"].ToString();
    //}

    //Setup User and Session Variables
    //sCookieUserID = Request.Cookies["useridx"];
    sCookieUserID = Request.Cookies["userid"];
    sSessionID    = HttpContext.Current.Session.SessionID;

    //Session["RedirectPage"] = "rd_classes/class_signup.aspx?classid=" + sClassID.ToString() + "&categoryid=" + sCategoryID.ToString() + "&categorytitle=" + sCategoryTitle;
    Session["RedirectPage"] = "rd_classes/class_signup.aspx?classid=" + sClassID.ToString() + "&categoryid=" + sCategoryID.ToString();
    Session["RedirectLang"]    = "Return to Class Signup";
    Session["LoginDisplayMsg"] = "";
    Session["DisplayMsg"]      = "";
    Session["ManageURL"]       = "";

    //If they do not have a userid set then take them to the login page automatically
    if (sCookieUserID == null || sCookieUserID.Value == "" || sCookieUserID.Value == "-1")
    {
        Session["LoginDisplayMsg"] = "Please sign in first and then we will send you right along.";
        sSessionIDName = "RedirectPage";

        //sUserSessionID = common.setASPNETtoASP_sessionVariables(Convert.ToInt32(sOrgID),
        //                                                        sSessionID,
        //                                                        Convert.ToString(sCookieUserID.Value),
        //                                                        sSessionIDName,
        //                                                        Convert.ToString(Session["RedirectPage"]));
        //Response.Redirect("../user_aspnet_to_asp.asp");
        //Response.Redirect("../dtb_user_login.asp?usid=" + sUserSessionID.ToString());
        Response.Redirect("../rd_user_login.aspx");
    }
    else
    {
        sUserID = sCookieUserID.Value;
    }

    //See if the user is missing any "key" data OR if the user address is required
    //The 2nd validation is for Bullhead City only.
    sIsUserMissingKeyData = common.UserIsMissingKeyData(Convert.ToInt32(sUserID));
    sIsUserMissingAddress = classes.UserIsMissingAddress(Convert.ToInt32(sUserID));
    
    if ((sIsUserMissingKeyData) || (sOrgHasFeature_registrationReqAddress && sIsUserMissingAddress))
    {
        Session["DisplayMsg"] = "Your account is missing some critical information we need to know before you can continue.";
        Response.Redirect("../manage_account.asp");
    }
    
    if (sOrgID == "7")
    {
        sPageTitle = sOrgName;
    }

    //Set up variables for common user controls
    egov_navigation.egovsection  = "CLASSES_NOSEARCH";
    egov_navigation.rootcategory = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid   = Convert.ToString(sCategoryID);

    //Set up variables for feature specific user controls
    egov_classes_memberwarning.orgid = sOrgID;
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../rd_global.css" />
  <link type="text/css" rel="stylesheet" href="styles_class.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>

  <%="<script type=\"text/javascript\" src=\"" + common.getBaseURL("") + "/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>
  <script type="text/javascript" src="https://maps.googleapis.com/maps/api/js?sensor=false"></script>
  
  <script type="text/javascript">
      $(document).ready(function() {
          setupCoachFields();
          
          $('#familymemberid').change(function() {
              var lcl_familymemberid = $('#familymemberid option:selected').val();

              $('#classEmergencyInfoDiv').slideUp('slow', function() {
                 $('#classEmergencyInfoDiv').html('');

                 $.post('class_confirminfo.aspx', {
                     orgid:          '<%=sOrgID%>',
                     familymemberid: lcl_familymemberid
                 }, function(result) {
                    $('#classEmergencyInfoDiv').html(result);
                    $('#emergencycontact').val($('#emergencycontactmaint').val());
                    $('#emergencyphone').val($('#emergencyphone_areacode').val() + $('#emergencyphone_exchange').val() + $('#emergencyphone_line').val());
                    $('#classEmergencyInfoDiv').slideDown('slow');
                 });
              });
          });

          $('#addToCartButton').click(function() {
              validateForm();
          });
      });
      
      function goToList() {
          location.href = 'class_list.aspx?categoryid=<%=sCategoryID.ToString()%>';
      }
      
      function goToDetails() {
          var lcl_url  = 'class_details.aspx';
              lcl_url += '?classid=<%=sClassID%>';
              lcl_url += '&categoryid=<%=sCategoryID%>';

          location.href = lcl_url;
      }

      function updateFamily(iUserID) {
          location.href = '../family_list.asp?userid=' + iUserID;
      }

      function setupCoachFields() {
          //Check to see if a value has been selected in the "I would like to" volunteer coach field.
          //If one has been selected then enable the other volunteer coach fields.
          //If one has not then disable them.
          //lcl_type = document.getElementById("rostercoachtype").value;
          lcl_type = $('#rostercoachtype').val();

          if (lcl_type != "") {
              //document.getElementById("volunteerCoachInfo").style.visibility="visible";
              $('#volunteerCoachInfo').show('slow');
          } else {
              //document.getElementById("volunteerCoachInfo").style.visibility="hidden";
              $('#volunteerCoachInfo').hide('slow');
          }
      }

      function getSelectedRadioValue(buttonGroup) {
          //return the value of the selected radio button or '' if no button is selected.
          var i = getSelectedRadio(buttonGroup);

          return $('#' + buttonGroup + i).val();
      }

      function getSelectedRadio(buttonGroup) {
          //return the array number of the selected radio button or -1 if no button is selected.
          var lcl_totalcount = $('#activityTimeTotalCount').val();

          if(lcl_totalcount > 0) {
              for (var i = 1; i <= lcl_totalcount; i++) {
                  if(document.getElementById(buttonGroup + i))
                  {
                     if (document.getElementById(buttonGroup + i).checked) {
                         return i;
                     }
                  }
              }
          } else {
              //if we get to this point, no radio button is selected          
              return 1;
          }
      }
      
      function changebuyorwait(iTimeID) {
         $('#buyorwait').val($('#buyorwait' + iTimeID).val());
         clearMsg('timeid1');
      }
      
      //****************************************************************************
      // These javascript functions are coded in shorthand and MUST be
      // copied EXACTLY as they are if used anywhere else.
      //****************************************************************************
<% if(sOrgHasFeature_emergencyInfoRequired) { %>
      var isNN = (navigator.appName.indexOf("Netscape") != -1);
      
		function autoTab( input,len, e ) 
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
			    //alert(input.id + getIndex(input)+addNdx);
			    //input.form[(getIndex(input)+addNdx) % input.form.length].focus();
			    var next = $(e.target).next();
			    next.focus();
		    }

		
		    function containsElement( arr, ele ) 
		    {
			   var found = false, index = 0;

			   while(!found && index < arr.length)
			   	  if(arr[index] == ele)
					found = true;
				else
					index++;
		  	   return found;
		    }

		    function getIndex( input ) 
		    {
			   var index = -1, i = 0, found = false;
			   var iTotalInputs = input.form.length;

			   while (i < input.form.length && index == -1)
				if (input.form[i] == input)index = i;
				else i++;
					return index;
			   }
			   return true;
		}
<% } %>
		//****************************************************************************
		//
		//****************************************************************************

		    

      function validateForm() {
          var lcl_return_false = 'N';
          var lcl_focus        = '';

          //If the ticket field exists, check that something is entered.
          var bExistsQuantity = eval($('#quantity'));
          var bExistsTerms    = eval($('#terms'));
          
          if (bExistsQuantity) {
          
              if ($('#quantity').val() == '') {
                  lcl_focus = 'quantity';
                  inlineMsg(document.getElementById("quantity").id, '<strong>Required Field Missing: </strong>Ticket quantity', 10, 'quantity');
                  lcl_return_false = 'Y';
              } else {
                  var rege = /^\d+$/;
                  var Ok = rege.test($('#quantity').val());

                  if (!Ok) {
                      lcl_focus = 'quantity';
                      inlineMsg(document.getElementById("quantity").id, '<strong>Invalid Value: </strong>The ticket quantity must be a number.', 10, 'quantity');
                      lcl_return_false = 'Y';
                  } else {
                      //Check that the quantity is not more than what is available if they are buying.
                      if ($('#buyorwait').val() == 'B') {
                          var iTimeID;
                          var iAvail;
                          var iQty;
                          
                          //alert('LEFT OFF HERE!!! \nNeed to figure out how to check to see if a time radio option is selected and which one it is.');
                          iTimeID = getSelectedRadioValue('timeid');
                          
                          //Get the availability for the selected time.
                          iAvail = Number(eval($('#avail' + iTimeID).val()));
                          iQty   = Number(eval($('#quantity').val()));

                          //Check that the ticket quantity is not greater than what is available.
                          if (iQty > iAvail) {
                              lcl_focus = 'quantity';
                              inlineMsg(document.getElementById("quantity").id, '<strong>Invalid Value: </strong>The ticket quantity cannot be greater than the availability.', 10, 'quantity');
                              lcl_return_false = 'Y';
                          }
                      }
                  }
              }

              //Check to see if the waiver is checked.
              if (bExistsTerms) {
                  if(document.getElementById('terms'))
                  {
                     if (document.getElementById('terms').checked == false) {
                         lcl_focus = 'terms';
                         inlineMsg(document.getElementById("terms").id, '<strong>Required Field Missing: </strong>You must agree to the waiver and release terms by checking the box next to \'I agree\'.', 10, 'terms');
                         lcl_return_false = 'Y';
                     }
                  }
              }
          }
          
   <% if(sOrgHasFeature_emergencyInfoRequired) { %>
          var lcl_emergencyphone  = $('#emergencyphone_areacode').val();
              lcl_emergencyphone += $('#emergencyphone_exchange').val();
              lcl_emergencyphone += $('#emergencyphone_line').val();
   
          //Validate the emergency phone
          if(lcl_emergencyphone == '') {
              inlineMsg(document.getElementById('emergencyphone_line').id, '<strong>Required Field Missing: </strong>Emergency Phone', 10, 'emergencyphone_line');
              lcl_focus = 'emergencyphone_line';
              lcl_return_false = 'Y';
          } else {
              var lcl_emergency_areacode = $('#emergencyphone_areacode').val();
              var lcl_emergency_exchange = $('#emergencyphone_exchange').val();
              var lcl_emergency_line     = $('#emergencyphone_line').val();
                              
              if(lcl_emergencyphone.length < 10) {
                  inlineMsg(document.getElementById("emergencyphone_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Emergency Phone.',10,'emergencyphone_line');
                  lcl_focus = 'emergencyphone_areacode';
                  lcl_return_false = 'Y';
              } else {
                  var emergencyPhone = new Number(lcl_emergency_areacode + lcl_emergency_exchange + lcl_emergency_line);
                                  
                  if(emergencyPhone.toString() == 'NaN') {
                      inlineMsg(document.getElementById("emergencyphone_line").id,'<strong>Invalid Value: </strong>Emergency Phone must be numeric',10,'emergencyphone_line');
                      lcl_focus = 'emergencyphone_areacode';
                      lcl_return_false = 'Y';
                  } else {
                      //Set the actual value to the "generated field" value typed in by the user
                      $('#emergencyphone').val(lcl_emergency_areacode + lcl_emergency_exchange + lcl_emergency_line);
                  }
              }
          }
          if($('#emergencycontactmaint').val() == '') {
              inlineMsg(document.getElementById('emergencycontactmaint').id, '<strong>Required Field Missing: </strong>Emergency Contact', 10, 'emergencycontactmaint');
              lcl_focus = 'emergencycontactmaint';
              lcl_return_false = 'Y';
          } else {
              //Set the actual value to the "generated field" value typed in by the user
              $('#emergencycontact').val($('#emergencycontactmaint').val());
          }
   <%
      }
      
      if (sOrgHasFeature_customRegistrationCraigCO)
      {
   %>
          //Validate Team Registration Fields
          if ($('#displayrosterpublic').val() == 'True') {
          
              if ($('#teamreg_coach_enabled').val() == 'BOTH') {
              
                  //Check to see if a "coach type" has been selected.
                  //If so then Full Name and at least one of the phone numbers and/or email are required.
                  if ($('#rostercoachtype').val() != '') {

                      //Build the daytime phone
                      var lcl_dayphone  = $('#skip_volcoachday_areacode').val();
                          lcl_dayphone += $('#skip_volcoachday_exchange').val();
                          lcl_dayphone += $('#skip_volcoachday_line').val();

                      //Build the cell phone
                      var lcl_cellphone  = $('#skip_volcoachcell_areacode').val();
                          lcl_cellphone += $('#skip_volcoachcell_exchange').val();
                          lcl_cellphone += $('#skip_volcoachcell_line').val();
                      
                      //At least one method of contact is required
                      if(lcl_dayphone == '' && lcl_cellphone == '' && $('#rostervolunteercoachemail').val() == '') {
                          lcl_focus = 'skip_volcoachday_areacode';
                          inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Required Field Missing: </strong>One method of contact must be entered.',10,'skip_volcoachday_line');
                          lcl_return_false = 'Y';
                      } else {
                          //Validate the email
                          if($('#rostervolunteercoachemail').val() != '') {
                              //var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
							  var rege = /.+@.+\..+/i;
                              var Ok = rege.test($('#rostervolunteercoachemail').val());
                              
                              if(! Ok) {
                                  lcl_focus = 'rostervolunteercoachemail';
                                  inlineMsg(document.getElementById("rostervolunteercoachemail").id,'<strong>Invalid Value: </strong>The volunteer coach email must be in a valid format.',10,'rostervolunteercoachemail');
                                  lcl_return_false = 'Y';
                              }
                          }
                          
                          //Validate the cell phone
                          if(lcl_cellphone != '') {
                              var lcl_cell_areacode = $('#skip_volcoachcell_areacode').val();
                              var lcl_cell_exchange = $('#skip_volcoachcell_exchange').val();
                              var lcl_cell_line     = $('#skip_volcoachcell_line').val();
                              
                              if(lcl_cellphone.length < 10) {
                                  lcl_focus = 'skip_volcoachcell_areacode';
                                  inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Cell Phone.',10,'skip_volcoachcell_line');
                                  lcl_return_false = 'Y';
                              } else {
                                  var cellPhone = new Number(lcl_cell_areacode + lcl_cell_exchange + lcl_cell_line);
                                  
                                  if(cellPhone.toString() == 'NaN') {
                                      lcl_focus = 'skip_volcoachcell_areacode';
                                      inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>Cell Phone must be numeric',10,'skip_volcoachcell_line');
                                      lcl_return_false = 'Y';
                                  }
                              }
                          }
                          
                          //Validate the day phone
                          if(lcl_dayphone != '') {
                              lcl_day_areacode = $('#skip_volcoachday_areacode').val();
                              lcl_day_exchange = $('#skip_volcoachday_exchange').val();
                              lcl_day_line     = $('#skip_volcoachday_line').val();
                              
                              if(lcl_dayphone.length < 10) {
                                  lcl_focus = 'skip_volcoachday_areacode';
                                  inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Day Phone.',10,'skip_volcoachday_line');
                                  lcl_return_false = 'Y';
                              } else {
                                  var dayPhone = new Number(lcl_day_areacode + lcl_day_exchange + lcl_day_line);
                                  
                                  if(dayPhone.toString() == 'NaN') {
                                      lcl_focus = 'skip_volcoachday_areacode';
                                      inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>Day Phone must be numeric.',10,'skip_volcoachday_line');
                                      lcl_return_false = 'Y';
                                  }
                              }
                          }
                      }
                      
                      //Validate the Full Name
                      if($('#rostervolunteercoachname').val() == '') {
                          lcl_focus = 'rostervolunteercoachname';
                          inlineMsg(document.getElementById("rostervolunteercoachname").id,'<strong>Required Field Missing: </strong>Volunteer Coach - Full Name.',10,'rostervolunteercoachname');
                          lcl_return_false = 'Y';
                      }
                  }
              }
              
              //Validate the Grade
              if($('#teamreg_grade_enabled').val() == 'BOTH') {
                  if($('#rostergrade').val() == '') {
                      lcl_focus = 'rostergrade';
                      inlineMsg(document.getElementById("rostergrade").id,'<strong>Required Field Missing: </strong>Grade.',10,'rostergrade');
                      lcl_return_false = 'Y';
                  } else {
                      var rosterGrade = new Number(document.getElementById("rostergrade").value);

                      if(rosterGrade.toString() == "NaN") {
                          lcl_focus = 'rostergrade';
                          inlineMsg(document.getElementById("rostergrade").id,'<strong>Invalid Value: </strong>Grade must be numeric.',10,'rostergrade');
                          lcl_return_false = 'Y';
                      }
                  }
              }
          }
   <% } %>
          //var bExistsTimeID = eval($('#timeid'));
          var lcl_totalcount = $('#activityTimeTotalCount').val();
          
          if(lcl_totalcount > 0) {
              //Multiple choices
              var lcl_radioChoiceSelected = 0;

              for(counter = 1; counter <= lcl_totalcount; counter++) {
                  if(document.getElementById('timeid' + counter))
                  {
                     if(document.getElementById('timeid' + counter).checked) {
                         lcl_radioChoiceSelected = lcl_radioChoiceSelected + 1;
                     }
                  }
              }
              
              if(lcl_radioChoiceSelected < 1) {
                  //If there were no selections made display an alert box
                  lcl_focus        = 'timeid1';
                  lcl_return_false = 'Y';
                  inlineMsg(document.getElementById('timeid1').id,'<strong>Required Field Missing: </strong>Please select an activity to continue with this purchase.',10,'timeid1');
              }
          } else {
              //One choice
              lcl_focus        = 'timeid1';
              lcl_return_false = 'Y';
              inlineMsg(document.getElementById('timeid1').id,'<strong>Required Field Missing: </strong>Please select an activity to continue with this purchase.',10,'timeid1');
          }

          if (lcl_return_false == 'Y') {
              if (lcl_focus != '') {
                  $('#' + lcl_focus).focus();
              }
          } else {
              $('#PurchaseForm').submit();
              //location.href = 'class_cart.aspx';
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
      <classes_memberWarning:classesMemberWarning id="egov_classes_memberwarning" runat="server" orgid="" />
      <% displayClassSignUp(Convert.ToInt32(sOrgID),
                            Convert.ToInt32(sUserID),
                            iRootCategoryID,
                            sCategoryID,
                            //sCategoryTitle,
                            sClassID); %>
    </div>
  </div>
  
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
