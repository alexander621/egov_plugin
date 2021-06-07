<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_details.aspx.cs" Inherits="rd_classes_class_details" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<%@ Register TagPrefix="classes_memberWarning" TagName="classesMemberWarning" Src="../rd_includes/egov_classes_memberwarning.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID       = common.getOrgId();
    static string sOrgName     = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle          = "E-Gov Services " + sOrgName;

    static Int32 iRootCategoryID = classes.getFirstCategory(sOrgID);
    Int32 sCategoryID            = iRootCategoryID;
    Int32 sClassID               = 0;

    Boolean sIsClassRegattaEvent = false;
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

    if (sOrgID == "7")
    {
        sPageTitle = sOrgName;
    }

    //Set up variables for common user controls
    egov_navigation.egovsection = "CLASSES_NOSEARCH";
    egov_navigation.rootcategory = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid = Convert.ToString(sCategoryID);

    //Set up variables for feature specific user controls
    egov_classes_memberwarning.orgid = sOrgID;
    
    //Set up page variables
    sIsClassRegattaEvent = classes.classIsRegattaEvent(sClassID);
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../rd_global.css" />
  <link type="text/css" rel="stylesheet" href="styles_class.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>

  <%="<script type=\"text/javascript\" src=\"" + common.getBaseURL("") + "/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>
  <script src="https://maps.googleapis.com/maps/api/js?sensor=false&key=AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"></script>


<script type="text/javascript">
    $(document).ready(function() {
        initialize();

        $('#location_getDirectionsButton').click(function() {
            openGoogleMap();
        });
    });
  
  function initialize() {
      var sFullAddress = $('#location_fulladdress').val();

      if (sFullAddress != '') {
          var geo = new google.maps.Geocoder;

          geo.geocode({ 'address': sFullAddress }, function(results, status) {
              var lcl_latlng;
              var lcl_latitude;
              var lcl_longitude;
              var lcl_latlng_index;

              if (status == google.maps.GeocoderStatus.OK) {
                  lcl_latlng = results[0].geometry.location;

                  if (status == google.maps.GeocoderStatus.OK) {
                      lcl_latlng = results[0].geometry.location;
                      lcl_latlng = lcl_latlng.toString();
                      lcl_latlng = lcl_latlng.replace('(', '');
                      lcl_latlng = lcl_latlng.replace(')', '');

                      lcl_latlng_index = lcl_latlng.indexOf(',');
                      lcl_latitude = lcl_latlng.substr(0, lcl_latlng_index);
                      lcl_longitude = lcl_latlng.substr(lcl_latlng_index + 2);

                      $('#location_latitude').val(lcl_latitude);
                      $('#location_longitude').val(lcl_longitude);

                      var latlng = new google.maps.LatLng(lcl_latitude, lcl_longitude);

                      var myOptions = {
                          zoom: 13,
                          center: latlng,
                          mapTypeId: google.maps.MapTypeId.ROADMAP
                      };

                      var map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);
                      var marker;

                      setTimeout(function() {
                          marker = new google.maps.Marker({
                              map: map,
                              draggable: false,
                              animation: google.maps.Animation.DROP,
                              position: latlng,
                              title: $('#location_name').val()
                          });
                      }, 2000);
                  }

                  //} else {
                  //    alert("Geocode was not successful for the following reason: " + status);
              }
          });
      }
  }

  function openGoogleMap() {
      var lcl_fulladdress   = '';
      var lcl_googlemap_url = '';

      lcl_fulladdress = $('#location_fulladdress').val();

      if (lcl_fulladdress.length > 0) {
          lcl_fulladdress = lcl_fulladdress.replace(/ /gi, '+');
      }

      lcl_googlemap_url = 'http://maps.google.com/maps?q=' + lcl_fulladdress;

      window.open(lcl_googlemap_url, '_blank');
 
      event.preventDefault();
  }

  function goToRegattaMaint(iPageType, iClassID) {
      var lcl_pagetype = '';  //i.e. TEAM or MEMBER
      var lcl_page_url = 'regattamembersignup.asp';
      var lcl_classid  = '';

      if (iPageType != '' && iPageType != undefined) {
          lcl_pagetype = iPageType.toString().toUpperCase();
      }

      if (iClassID != '' && iClassID != undefined) {
          lcl_classid = iClassID;
      }
      
      if (lcl_pagetype == "TEAM") {
          lcl_page_url = 'regattateamsignup.asp';
      }

      location.href = lcl_page_url + '?classid=' + lcl_classid;
  }

  //function goToSignUp(iClassID, iCategoryID, iCategoryTitle) {
  function goToSignUp(iClassID, iCategoryID) {
      var lcl_URL  = "class_signup.aspx";
          lcl_URL += "?classid=" + iClassID;
          lcl_URL += "&categoryid=" + iCategoryID;
          //lcl_URL += "&categorytitle=" + encodeURI(iCategoryTitle.replace("'","\'"));
          
          location.href = lcl_URL;
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
<%
    if (sIsClassRegattaEvent)
    {
        displayRegattaItem(Convert.ToInt32(sOrgID),
                           sClassID,
                           sCategoryID);
    } else {
        displayClassInfo(Convert.ToInt32(sOrgID),
                         sClassID,
                         sCategoryID);
    }
%>
    </div>
  </div>
  
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
