<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: MAP_TEST.ASP
' AUTHOR: STEVE LOAR
' CREATED: 11/28/2006
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  THIS MODULE DISPLAYS PROBLEM LOCATIONS.
'
' MODIFICATION HISTORY
' 1.0   11/28/06	STEVE LOAR - INITIAL VERSION CREATED
' 2.0   01/31/07	JOHN STULLENBERGER - ADDED CODE REMOVE THE NEED FOR SEPARATE GOOGLE LICENSES.
' ???   09/06/07    DAVID BOYER - ADDED CODE FOR SUB-STATUS
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
'if isFeatureOffline("action line") = "Y" then
'   response.redirect "../admin/outage_feature_offline.asp"
'end if

'INITIALIZE VARIABLES
 Dim sLat, sLng, iOrgid, sMapQuery, sBaseURL

 sLevel    = "../" ' Override of value from common.asp
 sLat      = ""
 sLng      = ""
 iOrgid    = request("orgid")
 sMapQuery = request("map_query")
 sOrderBy  = request("orderBy")
 sBaseURL  = request("current_url")

'Setup ORDER BY 
 if sOrderBy = "streetname" then
    lcl_order_by = "UPPER(sortstreetname), CAST(streetnumber AS int) "
 elseif sOrderBy = "submittedby" then
    lcl_order_by = "UPPER(userlname), UPPER(userfname) "
 else
    lcl_order_by = sOrderBy
 end if

'CHECK TO MAKE SURE WE HAVE DATA BEFORE CONTINUING
 if iOrgid = "" OR sMapQuery = "" then
    response.write "!Missing Data Cannot Continue!"
    response.end
 end if

'GET CITY'S MAP CENTER POINT
 GetCityPoint sLat, sLng 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<title>E-Gov Services <%=sOrgName%></title>
<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
<link rel="stylesheet" type="text/css" href="../global.css" />

<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%=Application("MAP_KEY")%>" type="text/javascript"></script>
<script type="text/javascript">
//<![CDATA[

var gmarkers     = new Array(); 
var badaddresses = new Array();

//var points = new Array();  // for panning to a point
var i  = 0;
var ba = 0;
var map;

function load() {
  if(GBrowserIsCompatible()) {
     map = new GMap2(document.getElementById("map"));
     map.addControl(new GLargeMapControl());
     map.addControl(new GMapTypeControl());
     map.setCenter(new GLatLng(<%=sLat%>, <%=sLng%>), 13);

     //gmarkers = new Array(); 
     //badaddresses = new Array();
     //var i = 0;
     //var ba = 0;
     var point;
     <% ShowPoints %>
  }
}

//Creates a marker at the given point with the given number label
function createMarker(point, number, ID, sMsg) {

  //Create our marker icon
  var icon              = new GIcon();

  //icon.image = "http://labs.google.com/ridefinder/images/mm_20_red.png";
  //icon.shadow = "http://labs.google.com/ridefinder/images/mm_20_shadow.png";
  icon.image            = "../images/small_orange.gif";
  icon.iconSize         = new GSize(9, 9);

  //icon.shadowSize = new GSize(22, 20);
  icon.iconAnchor       = new GPoint(6, 9);
  icon.infoWindowAnchor = new GPoint(5, 1);

  //var marker = new GMarker(point, icon);

  var marker            = new GMarker(point, icon);
  //INCREMENT MARKER ARRAY TO REFERENCE OFF MAP LINKS
  gmarkers[i]=marker;

  //points[i] = point;
  i++;
  GEvent.addListener(marker, "click", function() {
     marker.openInfoWindowHtml(sMsg + "<a href='http://<%=sBaseURL%>action_respond.asp?control=" + ID + "'>Click for Details</a>");
  });

  return marker;
}

function checkplot(iIndex) {

  // IF FALSE POINT WAS NOT PLOTTED
  //if (isbadaddress(iIndex) != true) {
  //	alert('Error: This address could not be plotted!');
  //}
  //else
  //{
  // POINT WAS PLOTTED. MOVE TO MAP TO SHOW MARKER
  GEvent.trigger(gmarkers[iIndex], 'click');
  location.href='#';
  //}
}

//]]>
</script>
</head>
<body onload="load()" onunload="GUnload()">
<div id="content">
<!--CLOSE WINDOW-->
<form>
  <input type="button" onClick="window.close();" style="font-weight:bold;" value="Close Google Map Window">
</form>
<!--CLOSE WINDOW-->

<div id="centercontent">
  <div align="center" id="map" style="width: 795px; height: 595px"></div>
<p>
<%ListRequests()%>
</p>
  </div>
</div>

</body>
</html>
<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Sub GetCityPoint( ByRef sLat, ByRef sLng )
'------------------------------------------------------------------------------------------------------------
Sub GetCityPoint( ByRef sLat, ByRef sLng )
   'Get the point to center the map
    Dim sSql, oPoint

    sSql = "Select latitude, longitude from organizations where orgid = " & iOrgid

    Set oPoint = Server.CreateObject("ADODB.Recordset")
    oPoint.Open sSQL, Application("DSN"), 3, 1

    If Not oPoint.EOF Then 
       sLat = oPoint("latitude")
       sLng = oPoint("longitude")
    End If 

    oPoint.close
    Set oPoint = Nothing 
End Sub 

'------------------------------------------------------------------------------------------------------------
' Sub ShowPoints()
'------------------------------------------------------------------------------------------------------------
Sub ShowPoints()
	Dim sSql, oPoint, iPointCount, sMsg

	iPointCount = 0

 'sSql = "Select actionrequestresponseid, latitude, longitude from egov_action_response_issue_location where actionrequestresponseid > 3523" 
 sSQL = sMapQuery
 sSQL = sSQL & " ORDER BY " & lcl_order_by

 if sOrderBy = "submit_date" then
    sSQL = sSQL & " desc"
 end if

'	sSql = sMapQuery & " Order by UPPER(streetaddress), CAST(streetnumber AS int)"
	'response.write sSql
	'response.end

	Set oPoint = Server.CreateObject("ADODB.Recordset")
	oPoint.Open sSQL, Application("DSN"), 3, 1

	Do While Not oPoint.EOF
	   If CDbl(oPoint("latitude")) <> CDbl(0) Then 
	      iPointCount = iPointCount + 1
	      iTrackingNumber = oPoint("action_autoid") & replace(FormatDateTime(oPoint("submit_date"),4),":","")
	      sMsg = "<div class='info'>"
	      'sMsg = sMsg & "<b>Marker: </b>" & iPointCount & "<br />"
	      sMsg = sMsg & "<b>Tracking Number: </b>" & iTrackingNumber                               & "<br />"
	      sMsg = sMsg & "<b>Ticket Name: </b>"     & replace(oPoint("action_formTitle"),vbcrlf,"") & "<br />"
       'sMsg = sMsg & "<b>Location: </b>"        & oPoint("streetnumber") & " " & oPoint("streetname") & "<br />"
	      sMsg = sMsg & "<b>Location: </b>"        & oPoint("streetname")                          & "<br />"
	      sMsg = sMsg & "</div>"
	      response.write  vbcrlf & "point = new GLatLng(" & oPoint("latitude") & "," & oPoint("longitude") & ");"
	      'response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",'" & oPoint("action_autoid") & "','" & sMsg & "'));" 
	      response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",""" & oPoint("action_autoid") & """,""" & sMsg & """));" 
	   End If 
	   oPoint.MoveNext
    Loop 

	oPoint.close
	Set oPoint = Nothing 
End Sub 

'--------------------------------------------------------------------------------------------------
' FUNCTION ListRequests()
'--------------------------------------------------------------------------------------------------
Function ListRequests()
  Dim sSql, oTickets, bgcolor, iMarkerCount

'  sSql = sMapQuery & " Order by streetname, streetnumber "
'  sSql = sMapQuery & " ORDER BY UPPER(streetaddress), CAST(streetnumber AS int) "

  sSQL = sMapQuery
  sSQL = sSQL & " ORDER BY " & lcl_order_by

  if sOrderBy = "submit_date" then
     sSQL = sSQL & " desc"
  end if

  Set oTickets = Server.CreateObject("ADODB.Recordset")
  oTickets.Open sSQL, Application("DSN"), 3, 1
  iMarkerCount = 0

  If NOT oTickets.EOF Then
    'START TABLE
     response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" class=""tablelist"">"
     response.write vbcrlf & "  <tr>"
     'response.write "<th>Marker</th>"
     response.write          "      <th>Tracking Number</th>"
   	 response.write          "      <th>Ticket Name</th>"
   	 response.write          "      <th>Location</th>"
     response.write          "  </tr>"

     bgcolor = "#eeeeee"
     Do While NOT oTickets.EOF 
        If bgcolor="#eeeeee" Then
           bgcolor="#ffffff" 
        Else
           bgcolor="#eeeeee"
        End If

        If CDbl(oTickets("latitude")) <> CDbl(0) Then 
           iMarkerCount = iMarkerCount + 1
           response.write vbcrlf & vbtab & "<tr bgcolor=""" & bgcolor & """ onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"" onClick=""checkplot("& (iMarkerCount - 1) &");"">"
          'response.write "     <td><a href=# onClick=""checkplot("& (iMarkerCount - 1) &");"" >" & iMarkerCount & "</a></td>"
          'response.write "     <td align=""center"">" & iMarkerCount & "</td>"
           response.write "    <td align=""center"">" & oTickets("action_autoid") & replace(FormatDateTime(oTickets("submit_date"),4),":","") & "</td>"
		         response.write "    <td align=""center"">" & oTickets("action_formTitle") & "</td>"
          'response.write "     <td align=""center"">" & oTickets("streetnumber") & " " & oTickets("streetname") & "</td></tr>"
      		   response.write "    <td align=""center"">" & oTickets("streetname")       & "</td>"
           response.write "</tr>"
        End If
        oTickets.MoveNext
     Loop

     response.write vbcrlf & "</table>"

     If iMarkerCount = 0 Then 
        response.write "<p>There are no mappable locations.</p>"
     End If 

  Else 
     response.write "<p>There are no mappable locations.</p>"
  End If

End Function
%>