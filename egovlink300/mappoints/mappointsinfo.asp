<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappointsinfo.asp
' AUTHOR:   David Boyer
' CREATED:  03/05/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the Mayor's Blog
'
' MODIFICATION HISTORY
' 1.0  04/06/10	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
' if isFeatureOffline("mappoints") = "Y" then
'    response.redirect "outage_feature_offline.asp"
' end if

 Dim oMapPoints

 set oMapPoints = New classOrganization

 lcl_mappointid = ""
 lcl_feature    = ""
 lcl_return_url = ""

 if request("m") <> "" then
    if not containsApostrophe(request("m")) then
       lcl_mappointid = trim(request("m"))
    end if
 end if

 if lcl_mappointid <> "" then
    if isnumeric(lcl_mappointid) then
       lcl_mappointid = CLng(lcl_mappointid)
    else
       response.redirect "mappoints.asp"
    end if
 else
    response.redirect "mappoints.asp"
 end if

 if request("f") <> "" then
    if not containsApostrophe(request("f")) then
       lcl_feature = request("f")

       if lcl_return_url <> "" then
          lcl_return_url = lcl_return_url & "&f=" & lcl_feature
       else
          lcl_return_url = "?f=" & lcl_feature
       end if
    end if
 end if

 if request("mpt") <> "" then
    if not containsApostrophe(request("mpt")) then
       lcl_mpt = request("mpt")

       if lcl_return_url <> "" then
          lcl_return_url = lcl_return_url & "&m=" & lcl_mpt
       else
          lcl_return_url = "?m=" & lcl_mpt
       end if
    end if
 end if

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Set up the BODY "onload" and "onunload"
 'lcl_onload   = "load()"
 'lcl_onunload = "GUnload()"

'Set up the page variables
 lcl_mappoint_typeid     = 0
 lcl_description         = ""
 'lcl_statusid            = 0
 lcl_mappoint_latitude   = ""
 lcl_mappoint_longitude  = ""
 lcl_mappoint_displaymap = 1
 
'Retrieve the mappoint data
 sSQL = "SELECT "
 sSQL = sSQL & " mp.mappoint_typeid, "
 sSQL = sSQL & " mpt.description, "
 'sSQL = sSQL & " mp.statusid, "
 sSQL = sSQL & " mp.latitude, "
 sSQL = sSQL & " mp.longitude, "
 sSQL = sSQL & " mpt.displayMap "
 sSQL = sSQL & " FROM egov_mappoints mp, egov_mappoints_types mpt "
 sSQL = sSQL & " WHERE mp.mappoint_typeid = mpt.mappoint_typeid "
 sSQL = sSQL & " AND mp.orgid = " & iorgid
 sSQL = sSQL & " AND mp.mappointid = " & lcl_mappointid
 sSQL = sSQL & " AND mp.isActive = 1 "
 'sSQL = sSQL & " AND mptf.displayInResults = 1 "

 	set oMapPointInfo = Server.CreateObject("ADODB.Recordset")
 	oMapPointInfo.Open sSQL, Application("DSN"), 3, 1

  if not oMapPointInfo.eof then
     lcl_mappoint_typeid     = oMapPointInfo("mappoint_typeid")
     lcl_description         = oMapPointInfo("description")
     'lcl_statusid            = oMapPointInfo("statusid")
     lcl_mappoint_latitude   = oMapPointInfo("latitude")
     lcl_mappoint_longitude  = oMapPointInfo("longitude")
     lcl_mappoint_displaymap = oMapPointInfo("displayMap")
  end if

  oMapPointInfo.close
  set oMapPointInfo = nothing

  lcl_onload   = ""
  lcl_onunload = ""

  if lcl_mappoint_displaymap _
  AND (lcl_mappoint_latitude <> "" AND not isnull(lcl_mappoint_latitude)) _
  AND (lcl_mappoint_longitude <> "" AND not isnull(lcl_mappoint_longitude)) then
       lcl_onload   = "initialize(" & lcl_mappoint_latitude & "," & lcl_mappoint_longitude & ");"
       lcl_onunload = "GUnload();"
  end if
%>
<html>
<head>
 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

 	<script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/easyform.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>

<% 
sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"
if lcl_mappoint_displaymap then %>
 	<script src="https://maps.google.com/maps?file=api&amp;v=2&amp;key=<%=sGoogleMapAPIKey%>" type="text/javascript"></script>

<script type="text/javascript">
//    var lcl_streetview;

//    function initialize() {
//      var lcl_location = new GLatLng(42.345573,-71.098326);

//      panoramaOptions = { latlng:lcl_location };
//      lcl_streetview  = new GStreetviewPanorama(document.getElementById("streetview"), panoramaOptions);
//      GEvent.addListener(lcl_streetview, "error", handleNoFlash);
//    }
    
//    function handleNoFlash(errorCode) {
//      if (errorCode == FLASH_UNAVAILABLE) {
//          alert("Error: Flash doesn't appear to be supported by your browser");
//          return;
//      }
//    }

    var myPano                 = '';
    var lcl_streetviewlocation = '';
    
    function initialize(iLatitude,iLongitude) {
      lcl_streetviewlocation = new GLatLng(iLatitude,iLongitude);

      panoramaOptions = { latlng:lcl_streetviewlocation };
      myPano = new GStreetviewPanorama(document.getElementById("pano"), panoramaOptions);
      GEvent.addListener(myPano, "error", handleNoFlash);
    }
    
    function handleNoFlash(errorCode) {
      if (errorCode == "FLASH_UNAVAILABLE") {
        alert("Error: Flash doesn't appear to be supported by your browser");
        return;
      }
    }  
</script>
<% end if %>

<script language="javascript">
  function openWin(p_url, p_width, p_height) {
    w = 900;
    h = 600;

    if((p_width!="")&&(p_width!=undefined)) {
        w = p_width;
    }

    if((p_height!="")&&(p_height!=undefined)) {
        h = p_height;
    }

    l = (screen.AvailWidth/2)-(w/2);
    t = (screen.AvailHeight/2)-(h/2);
    lcl_url = p_url;

    eval('window.open("' + lcl_url + '", "_picker", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
  }

  function openGoogleMap() {
    var lcl_google_address = "";
    var lcl_googlemap_url  = "";

    lcl_google_address = document.getElementById("googleAddress").value;

    if(lcl_google_address != "") {
       lcl_googlemap_url += "http://maps.google.com/maps?hl=en&expIds=17259,17291,27615,27846,28155&sugexp=ldymls&xhr=t";
       lcl_googlemap_url += "&q=" + lcl_google_address;
       lcl_googlemap_url += "&cp=17&um=1&ie=UTF-8&sa=N&tab=wl";
    }

    openWin(lcl_googlemap_url,"","");
  }
</script>
</head>
<!--#include file="../include_top.asp"-->

<% RegisteredUserDisplay("../") %>

<div id="content">
  <div id="centercontent">
<p>
<table border="0" cellspacing="0" cellpadding="2">
  <tr>
      <td colspan="2">
          <p><button type="button" name="returnButton" id="returnButton" class="button" style="cursor:pointer;" onclick="location.href='mappoints.asp<%=lcl_return_url%>'">Return to <%=lcl_description%></button></p><br />
      </td>
  </tr>
  <tr valign="top">
      <td>
          <fieldset>
            <legend style="color:#800000"><%=lcl_description%>&nbsp;</legend>
          <%
            response.write "<div align=""right"">" & vbcrlf
                            displayAddThisButton iorgid
            response.write "</div>" & vbcrlf

            showMapPointInfo iorgid, lcl_mappointid
          %>
          </fieldset>
      </td>
      <td>
          <% if lcl_mappoint_displaymap then %>
          <div align="right"><a href="http://maps.google.com/support/bin/static.py?page=guide.cs&guide=21670&topic=21671&answer=144350" color="#0000ff" target="_blank">How to navigate in Google Maps</a></div>
          <div name="pano" id="pano" style="width: 500px; height: 300px"></div>
          <div style="text-align:center; color:#800000;">NOTE: Street-Level View is approximate.  You may need to rotate the image</div>
          <% end if %>
      </td>
  </tr>
</table>
</p>
<p>&nbsp;</p>

  </div>
</div>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub showMapPointInfo(p_orgid, p_mappointid)

  sSQL = "SELECT mpv.mp_valueid, "
  sSQL = sSQL & " mptf.fieldname, "
  sSQL = sSQL & " mpv.fieldvalue, "
  sSQL = sSQL & " mpv.fieldtype, "
  sSQL = sSQL & " mptf.resultsOrder "
  sSQL = sSQL & " FROM egov_mappoints_values mpv, egov_mappoints_types_fields mptf "
  sSQL = sSQL & " WHERE mpv.mp_fieldid = mptf.mp_fieldid"
  sSQL = sSQL & " AND mpv.orgid = " & p_orgid
  sSQL = sSQL & " AND mpv.mappointid = " & p_mappointid
  sSQL = sSQL & " AND mptf.displayInInfoPage = 1 "
  sSQL = sSQL & " AND mpv.fieldvalue <> '' "
  sSQL = sSQL & " AND mpv.fieldvalue is not null "
  sSQL = sSQL & " ORDER BY mptf.resultsOrder "

 	set oMPTFieldInfo = Server.CreateObject("ADODB.Recordset")
 	oMPTFieldInfo.Open sSQL, Application("DSN"), 3, 1

  if not oMPTFieldInfo.eof then

     response.write "<p>" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:400px;"">" & vbcrlf

     do while not oMPTFieldInfo.eof

       'Setup/Format the fieldvalue for display
        lcl_fieldvalue     = ""
        lcl_google_address = ""

        if oMPTFieldInfo("fieldvalue") <> "" then
           lcl_fieldvalue = replace(replace(oMPTFieldInfo("fieldvalue"),chr(10),""),chr(13),"<br />")
        end if

        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td nowrap=""nowrap""><strong>" & oMPTFieldInfo("fieldname") & "</strong></td>" & vbcrlf
        response.write "      <td>" & lcl_fieldvalue & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        if oMPTFieldInfo("fieldtype") = "ADDRESS" AND lcl_fieldvalue <> "" then
           lcl_google_address = replace(lcl_fieldvalue," ","+")

           response.write "  <tr>" & vbcrlf
           response.write "      <td>&nbsp;</td>" & vbcrlf
           response.write "      <td>" & vbcrlf
           response.write "          <input type=""hidden"" name=""googleAddress"" id=""googleAddress"" value=""" & lcl_google_address& """ size=""20"" />" & vbcrlf
           response.write "          <input type=""button"" name=""openGoogleMap"" id=""openGoogleMap"" class=""button"" value=""Open in Google Maps"" onclick=""openGoogleMap()"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "  </tr>" & vbcrlf
        end if

        oMPTFieldInfo.movenext
     loop

     response.write "</table>" & vbcrkf
     response.write "</p>" & vbcrlf

  end if

  oMPTFieldInfo.close
  set oMPTFieldInfo = nothing

end sub

'------------------------------------------------------------------------------
sub showPoints(p_SQL, p_orgid, iMapPointTypeID)

	 iPointCount             = 1
  iRowCount               = 0
  lcl_showpoints          = ""
  lcl_displayheaders      = ""
  lcl_previous_mappointid = 0
  lcl_latitude            = 0.00
  lcl_longitude           = 0.00

 	set oPoints = Server.CreateObject("ADODB.Recordset")
 	oPoints.Open p_SQL, Application("DSN"), 3, 1

  if not oPoints.eof then
    	do while not oPoints.eof

        iRowCount   = iRowCount + 1

        if lcl_previous_mappointid <> oPoints("mappointid") then
           if iRowCount > 1 then
              lcl_showpoints = lcl_showpoints & "</table>"
              lcl_showpoints = lcl_showpoints & "</div>"

           			response.write "point = new GLatLng(" & lcl_latitude & "," & lcl_longitude & ");" & vbcrlf
           			response.write "map.addOverlay(createMarker(point, " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "'));" & vbcrlf

              iPointCount    = iPointCount + 1
              iRowCount      = 1
              lcl_latitude   = 0.00
              lcl_longitude  = 0.00
           end if

           lcl_showpoints = ""
           lcl_showpoints = lcl_showpoints & "<div class=""info"">"
           lcl_showpoints = lcl_showpoints & "<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""text-align:left;"">"
           lcl_showpoints = lcl_showpoints &   "<tr>"
           lcl_showpoints = lcl_showpoints &       "<td colspan=""2"" align=""center""><strong># " & iPointCount & "</strong></td>"
           lcl_showpoints = lcl_showpoints &   "</tr>"
        end if

        if oPoints("fieldvalue") <> "" then
           lcl_showpoints = lcl_showpoints &   "<tr>"
           lcl_showpoints = lcl_showpoints &       "<td><strong>" & oPoints("fieldname")  & ": </strong></td>"
           lcl_showpoints = lcl_showpoints &       "<td>"         & oPoints("fieldvalue") & "</td>"
           lcl_showpoints = lcl_showpoints &   "</tr>"
        end if

        lcl_previous_mappointid = oPoints("mappointid")
        lcl_label               = buildStreetAddress(oPoints("streetnumber"), oPoints("streetprefix"), oPoints("streetaddress"), oPoints("streetsuffix"), oPoints("streetdirection"))
        lcl_latitude            = oPoints("latitude")
        lcl_longitude           = oPoints("longitude")
        lcl_mpcolor             = oPoints("mappointcolor")

        oPoints.movenext
     loop

     if iRowCount > 1 then
        lcl_showpoints = lcl_showpoints & "</table>"
        lcl_showpoints = lcl_showpoints & "</div>"

        response.write "point = new GLatLng(" & lcl_latitude & "," & lcl_longitude & ");" & vbcrlf
        response.write "map.addOverlay(createMarker(point, " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "'));" & vbcrlf
     end if

  end if

  oPoints.close
  set oPoints = nothing

end sub
%>
