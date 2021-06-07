<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_line_map.asp
' AUTHOR: Steve Loar
' CREATED: 12/12/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays problem locations on a map.
'
' MODIFICATION HISTORY
' 1.0  12/12/06	 Steve Loar - Initial Version Created
' 1.1	 03/26/07	 Steve Loar - Check of residentaddressid added for spiders and bots
' 1.2  01/22/08  David Boyer - Added "isFeatureOffline" to screen
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

	Dim iOrgId, sLat, sLng, sPointLat, sPointLng, sAddress
	sLat = ""
	sLng = ""
	sPointLat = ""
	sPointLng = ""
	sAddress = ""
	iOrgId = 5

	' chenged to check numeric here to eliminate the rain of errors generated from the cLng() conversion that was here - SJL 5/7/2013
	If IsNull(request("residentaddressid")) Or request("residentaddressid") = "" Or IsNumeric(request("residentaddressid")) = False Then
		response.redirect "action.asp"
	End If 

	' if any bad address ids get through above, they will most likely fail the cLng() call here - SJL 5/7/2013
	iOrgId = GetLatLng( CLng(request("residentaddressid")), sPointLat, sPointLng, sAddress )

	GetCityPoint iOrgId, sLat, sLng 

%>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8"/>

    <title>E-Gov Services <%=sOrgName%></title>
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

    <script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAAStOU3ZcVxvfIkpDt0vutZRRhH_ubqthHQhc8uJXKQ2FoSctO4BR2n0L6MBF66QP3CpygU6yG0BVT9Q"
      type="text/javascript"></script>

    <script type="text/javascript">
    //<![CDATA[

	var gmarkers = new Array(); 
	var badaddresses = new Array();
	//var points = new Array();  // for panning to a point
	var i = 0;
	var ba = 0;
	var map;

    function load() 
	{
      if (GBrowserIsCompatible()) 
	  {
        map = new GMap2(document.getElementById("map"));
		map.addControl(new GLargeMapControl());
		map.addControl(new GMapTypeControl());
        map.setCenter(new GLatLng(<%=sPointLat%>, <%=sPointLng%>), 13);
		
		var point;
		<% ShowPoint sPointLat, sPointLng, sAddress %>
		
      }
    }

	// Creates a marker at the given point with the given number label
	function createMarker(point, number, ID, sMsg) 
	{
		// Create our marker icon
		var icon = new GIcon();
		//icon.image = "http://labs.google.com/ridefinder/images/mm_20_red.png";
		//icon.shadow = "http://labs.google.com/ridefinder/images/mm_20_shadow.png";
		icon.image = "images/small_orange.gif";
		icon.iconSize = new GSize(9, 9);
		//icon.shadowSize = new GSize(22, 20);
		icon.iconAnchor = new GPoint(6, 9);
		icon.infoWindowAnchor = new GPoint(5, 1);
		//var marker = new GMarker(point, icon);

		var marker = new GMarker(point);
		//INCREMENT MARKER ARRAY TO REFERENCE OFF MAP LINKS
		gmarkers[i]=marker;
		//points[i] = point;
		i++;
		GEvent.addListener(marker, "click", function() {
			marker.openInfoWindowHtml(sMsg);
		});
		return marker;
	}

	function checkplot(iIndex)
	{
		
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
		<div id="centercontent">

		<p>
			<% =sAddress %>
		</p>

	    <div align="center" id="map" style="width: 795px; height: 595px"></div>

		</div>
	</div>

</body>
</html>

<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Sub GetCityPoint( ByVal iOrgId, ByRef sLat, ByRef sLng )
'------------------------------------------------------------------------------------------------------------
Sub GetCityPoint( ByVal iOrgId, ByRef sLat, ByRef sLng )
	' Get the point to center the map
	Dim sSql, oPoint

	sSql = "Select latitude, longitude from organizations where orgid = " & iOrgId

	Set oPoint = Server.CreateObject("ADODB.Recordset")
	oPoint.Open sSQL, Application("DSN"), 0, 1

	If Not oPoint.EOF Then 
		sLat = oPoint("latitude")
		sLng = oPoint("longitude")
	End If 

	oPoint.close
	Set oPoint = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowPoint()
'------------------------------------------------------------------------------------------------------------
Sub ShowPoint( sPointLat, sPointLng, sMsg )
	Dim oPoint


'	sMsg = "<div class=""info"">"
'	sMsg = sMsg & "<b>Tracking Number: </b>" & iTrackingNumber & "<br />"
'	sMsg = sMsg & "<b>Ticket Name: </b>" & replace(oPoint("action_formTitle"),vbcrlf,"") & "<br />"
'	sMsg = sMsg & "<b>Location: </b>" & oPoint("streetnumber") & " " & oPoint("streetname") & "<br />"
'	sMsg = sMsg & "</div>"

	response.write  vbcrlf & "point = new GLatLng(" & sPointLat & "," & sPointLng & ");"
	response.write vbcrlf & "map.addOverlay(createMarker(point, 1,'Problem Location','" & sMsg & "'));" 

End Sub 


'--------------------------------------------------------------------------------------------------
' FUNCTION ListRequests()
'--------------------------------------------------------------------------------------------------
Function ListRequests()
	Dim sSql, oTickets, bgcolor, iMarkerCount

	sSql = session("MAP_QUERY") & " Order by streetname, streetnumber"
	Set oTickets = Server.CreateObject("ADODB.Recordset")
	oTickets.Open sSql, Application("DSN"), 3, 1
	iMarkerCount = 0

	If NOT oTickets.EOF Then
		' START TABLE
		response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" class=""tablelist"">"
		response.write vbcrlf & "<tr>"
		'response.write "<th>Marker</th>"
		response.write "<th>Tracking Number</th><th>Ticket Name</th><th>Location</th></tr>"
		
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
				'response.write "<td><a href=# onClick=""checkplot("& (iMarkerCount - 1) &");"" >" & iMarkerCount & "</a></td>"
				'response.write "<td align=""center"">" & iMarkerCount & "</td>"
				response.write "<td align=""center"">" & oTickets("action_autoid") & replace(FormatDateTime(oTickets("submit_date"),4),":","") & "</td><td align=""center"">" & oTickets("action_formTitle") & "</td><td align=""center"">" & oTickets("streetnumber") & " " & oTickets("streetname") & "</td></tr>"
			End If
			oTickets.MoveNext
		Loop

		response.write vbcrlf & "</table>"

	Else 
		response.write "<p>There are no mappable locations.</p>"
	End If

End Function


'--------------------------------------------------------------------------------------------------
' Function GetLatLng( ByVal iResidentaddressid, ByRef sPointLat, ByRef sPointLng, ByRef sAddress )
'--------------------------------------------------------------------------------------------------
Function GetLatLng( ByVal iResidentaddressid, ByRef sPointLat, ByRef sPointLng, ByRef sAddress )
	Dim sSql, oAddressList, iOrgId

	iOrgId = 5
	sSql = "SELECT orgid, latitude, longitude, isnull(residentstreetnumber,'') as residentstreetnumber, "
	sSql = sSql & " residentstreetname, residentcity, residentstate " 
	sSql = sSql & " FROM egov_residentaddresses where residentaddressid = " & iResidentaddressid 

	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSql, Application("DSN"), 0, 1

	If Not oAddressList.EOF Then 
		sPointLat = oAddressList("latitude")
		sPointLng = oAddressList("longitude")
		sAddress = oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname") & " " & oAddressList("residentcity") & ", " & oAddressList("residentstate")
		iOrgId = oAddressList("orgid")
	End If 

	oAddressList.close
	Set oAddressList = Nothing 

	GetLatLng = iOrgId

End Function 


%>