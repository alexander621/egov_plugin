<!-- #include file="../includes/common.asp" //-->
<!DOCTYPE html>
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: MAP_TEST.ASP
' AUTHOR: STEVE LOAR
' CREATED: 11/28/2006
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  THIS MODULE DISPLAYS PROBLEM LOCATIONS.
'
' MODIFICATION HISTORY
' 1.0  11/28/2006	STEVE LOAR - INITIAL VERSION CREATED
' 2.0  01/31/2007	JOHN STULLENBERGER - ADDED CODE REMOVE THE NEED FOR SEPARATE GOOGLE LICENSES.
' 3.0  09/06/2007 David Boyer - ADDED CODE FOR SUB-STATUS
' 4.0  02/14/2014 David Boyer - Modified Google Map to latest version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim sLat, sLng, sOrgID, sMapQuery, sBaseURL
sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"

 sLevel    = "../" ' Override of value from common.asp
 sLat      = ""
 sLng      = ""
 sOrgID    = request("orgid")
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
 if sOrgID = "" OR sMapQuery = "" then
    response.write "!Missing Data Cannot Continue!"
    response.end
 end if

'GET CITY'S MAP CENTER POINT
 GetCityPoint sLat, sLng 

'Get all info to build Google Map and data points
 lcl_mappoints  = getMapPointsInfo("MAPPOINTS")
 lcl_bubbleinfo = getMapPointsInfo("BUBBLEINFO")
%>
<html>
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title>E-Gov Services <%=sOrgName%></title>

  <link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" href="../global.css" />

<style>
  #map_canvas
  {
    margin: 5px auto 10px auto;
    width:  600px;
    height: 400px;
    border: 1pt solid #000000;
    border-radius: 6px;
  }

  #buttonCloseWindow
  {
     margin: 5px;
  }

  #requestsList td
  {
     text-align: center;
  }

/*--------------------------------------------------------------------------------
BEGIN: Set up for screens with max of 800px
----------------------------------------------------------------------------------*/
@media screen and (max-width: 800px)
{
   #content
   {
      padding: 0;
      margin: 0;
   }

   #centercontent
   {
      width: 100%;
      margin: 0;
   }

   #content table
   {
      left: 0px;
      border: none;
      border-top: 1px solid #336699;
      border-bottom: 1px solid #336699;
   }
}

/*--------------------------------------------------------------------------------
BEGIN: Set up for screens with max of 620px
----------------------------------------------------------------------------------*/
@media screen and (max-width: 620px)
{
   #map_canvas
   {
      width: 94%;
      height: 300px;
   }
}
</style>

<script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false&key=<%= sGoogleMapAPIKey %>"></script>
<script src="../scripts/jquery-1.9.1.min.js"></script>

<script>
var mappoints  = [
<%=lcl_mappoints%>
];

var bubbleInfo = [
<%=replace(lcl_bubbleinfo,vbcrlf," ")%>
];

var markers    = [];
var infowindow = [];
var map;

$(document).ready(function() {
  $('#buttonCloseWindow').click(function() {
    window.close();
  });

  initialize();
});

function initialize()
{
    var lcl_latitude  = '<%=sLat%>';
    var lcl_longitude = '<%=sLng%>';

    var latlng = new google.maps.LatLng(lcl_latitude, lcl_longitude);

    var myOptions = {
        zoom: 13,
        center: latlng,
        mapTypeId: google.maps.MapTypeId.ROADMAP
    };

    map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);

    for (var i=0; i < mappoints.length; i++)
    {
       addMarker(i);
    }
}

  function addMarker(iRowCount) {
    var lcl_pointcolor         = '';
    var lcl_markernum          = iRowCount + 1;
    var lcl_marker             = '';
    var lcl_marker_url         = 'http://gmaps-samples.googlecode.com/svn/trunk/markers/';
    var lcl_marker_numberlimit = 99;

    if(lcl_pointcolor == '') {
       lcl_pointcolor = 'green';
    }

    if(lcl_markernum > lcl_marker_numberlimit) {
       lcl_marker = 'blank';
    } else {
       lcl_marker = 'marker' + lcl_markernum;
    }

    //var image = lcl_marker_url + lcl_pointcolor + '/' + lcl_marker + '.png';
    		var image = new google.maps.MarkerImage("https://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=" + (iRowCount+1) + "|92e415",
        		new google.maps.Size(21, 34),
        		new google.maps.Point(0,0),
        		new google.maps.Point(10, 34));

    markers.push(new google.maps.Marker({
       position:  mappoints[iRowCount],
       map:       map,
       draggable: false,
       animation: google.maps.Animation.DROP,
       icon:      image
    }));

    openInfoWindow(iRowCount);

    return iRowCount;
  }

  function openInfoWindow(iRowCount) {
    infowindow.push(new google.maps.InfoWindow({
       size:     new google.maps.Size(50,50),
       position: mappoints[iRowCount],
       content:  bubbleInfo[iRowCount]
    }));

    google.maps.event.addListener(markers[iRowCount], 'click', function() {
       infowindow[iRowCount].open(map,markers[iRowCount]);
    });
  }
</script>
</head>
<body>
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "    <div id=""map_canvas""></div>" & vbcrlf
  response.write "    <input type=""button"" name=""buttonCloseWindow"" id=""buttonCloseWindow"" value=""Close Window"" />" & vbcrlf
  response.write "    <div>" & vbcrlf
                        ListRequests()
  response.write "    </div>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Sub GetCityPoint( ByRef sLat, ByRef sLng )
   'Get the point to center the map
    Dim sSql, oPoint

    sSql = "Select latitude, longitude from organizations where orgid = " & sOrgID

    Set oPoint = Server.CreateObject("ADODB.Recordset")
    oPoint.Open sSQL, Application("DSN"), 3, 1

    If Not oPoint.EOF Then 
       sLat = oPoint("latitude")
       sLng = oPoint("longitude")
    End If 

    oPoint.close
    Set oPoint = Nothing 
End Sub 

'------------------------------------------------------------------------------
function ListRequests()
  dim sSql, oTickets, bgcolor, iMarkerCount
  dim lcl_marker_img, lcl_marker_url, lcl_marker
  dim lcl_latitude_actionline, lcl_latitude_mobileoption
  dim lcl_longitude_actionline, lcl_longitude_mobileoption

  iMarkerCount               = 0
  bgcolor                    = "#eeeeee"
  lcl_marker_img             = "http://gmaps-samples.googlecode.com/svn/trunk/markers/green/"
  lcl_marker_url             = ""
  lcl_marker                 = ""
  lcl_latitude               = ""
  lcl_longitude              = ""
  lcl_latitude_actionline    = "0.00"
  lcl_latitude_mobileoption  = "0.00"
  lcl_longitude_actionline   = "0.00"
  lcl_longitude_mobileoption = "0.00"

  sSQL = sMapQuery
  sSQL = sSQL & " ORDER BY " & lcl_order_by

  if sOrderBy = "submit_date" then
     sSQL = sSQL & " desc"
  end if

  set oTickets = Server.CreateObject("ADODB.Recordset")
  oTickets.Open sSQL, Application("DSN"), 3, 1

  if not oTickets.eof then
     response.write "<table id=""requestsList"" border=""0"" cellpadding=""2"" cellspacing=""0"" class=""tablelist"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <th>Marker #</th>" & vbcrlf
     response.write "      <th>Tracking Number</th>" & vbcrlf
   	 response.write "      <th>Ticket Name</th>" & vbcrlf
   	 response.write "      <th>Location</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oTickets.eof
        lcl_latitude_actionline    = oTickets("latitude")
        lcl_latitude_mobileoption  = oTickets("mobileoption_latitude")
        lcl_longitude_actionline   = oTickets("longitude")
        lcl_longitude_mobileoption = oTickets("mobileoption_longitude")
        lcl_display_latlng         = ""

       'If CDbl(oTickets("latitude")) <> CDbl(0) Then 
        if lcl_latitude_mobileoption <> "" AND lcl_latitude_mobileoption <> "0" AND lcl_latitude_mobileoption <> "0.00" then
           lcl_latitude  = replace(lcl_latitude_mobileoption,",","")
           lcl_longitude = replace(lcl_longitude_mobileoption,",","")
        else
           lcl_latitude  = lcl_latitude_actionline
           lcl_longitude = lcl_longitude_actionline
        end if

        if lcl_latitude = "" OR lcl_latitude = "0" then
           lcl_latitude = "0.00"
        end if

        if lcl_longitude = "" OR lcl_longitude = "0" then
           lcl_longitude = "0.00"
        end if

        if lcl_latitude <> "0.00" then
           if lcl_mappoints_info <> "" then
              lcl_mappoints_info = lcl_mappoints_info & ", " & vbcrlf
           end if

           iMarkerCount   = iMarkerCount + 1
           bgcolor        = changeBGColor(bgcolor, "#ffffff", "#eeeeee")
           lcl_marker     = "marker" & iMarkerCount

           if iMarkerCount > 99 then
              lcl_marker = "blank"
           end if

           lcl_marker_url = lcl_marker_img & lcl_marker & ".png"

           if left(trim(oTickets("streetname")),9) <> "Latitude:" then
              lcl_display_latlng = "<span style=""color: #800000;"">[" & lcl_latitude & "&nbsp;/&nbsp;" & lcl_longitude & "]</span>"

              if trim(oTickets("streetname")) <> "" then
                 lcl_display_latlng = "<br />" & lcl_display_latlng
              end if
           end if

           response.write "<tr onclick=""infowindow[" & iMarkerCount-1 & "].open(map,markers[" & iMarkerCount-1 & "]);location.href='#';"" bgcolor=""" & bgcolor & """ onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='pointer';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">" & vbcrlf
           response.write "    <td>" & iMarkerCount & "</td>" & vbcrlf
           response.write "    <td>" & oTickets("action_autoid") & replace(FormatDateTime(oTickets("submit_date"),4),":","") & "</td>" & vbcrlf
		         response.write "    <td>" & oTickets("action_formTitle") & "</td>" & vbcrlf
      		   response.write "    <td>" & vbcrlf
           response.write          oTickets("streetname") & lcl_display_latlng & vbcrlf
           response.write "    </td>" & vbcrlf
           response.write "</tr>" & vbcrlf
        end if

        oTickets.movenext
     loop

     response.write vbcrlf & "</table>"

     If iMarkerCount = 0 Then 
        response.write "<p>There are no mappable locations.</p>"
     End If 

  Else 
     response.write "<p>There are no mappable locations.</p>"
  End If

End Function

'------------------------------------------------------------------------------
function getMapPointsInfo(iInfoType)
  dim lcl_return, lcl_mappoints_info, lcl_latitude, lcl_longitude
  dim lcl_infoType, sSQL, oMapPoints
  dim lcl_latitude_actionline, lcl_latitude_mobileoption
  dim lcl_longitude_actionline, lcl_longitude_mobileoption
  dim lcl_display_location, lcl_display_latlng

  lcl_infoType               = "MAPPOINT"
  lcl_return                 = ""
  lcl_mappoints_info         = ""
  lcl_latitude               = ""
  lcl_longitude              = ""
  lcl_latitude_actionline    = "0.00"
  lcl_latitude_mobileoption  = "0.00"
  lcl_longitude_actionline   = "0.00"
  lcl_longitude_mobileoption = "0.00"
  lcl_display_location       = ""
  lcl_display_latlng         = ""

  if iInfoType <> "" then
     lcl_infoType = ucase(iInfoType)
  end if

  sSQL = sMapQuery
  sSQL = sSQL & " ORDER BY " & lcl_order_by

  if sOrderBy = "submit_date" then
     sSQL = sSQL & " desc"
  end if

 	set oMapPoints = Server.CreateObject("ADODB.Recordset")
 	oMapPoints.Open sSQL, Application("DSN"), 3, 1

 	do while not oMapPoints.eof
     lcl_latitude_actionline    = oMapPoints("latitude")
     lcl_latitude_mobileoption  = oMapPoints("mobileoption_latitude")
     lcl_longitude_actionline   = oMapPoints("longitude")
     lcl_longitude_mobileoption = oMapPoints("mobileoption_longitude")
     lcl_display_location       = ""
     lcl_display_latlng         = ""

     if lcl_latitude_mobileoption <> "" AND lcl_latitude_mobileoption <> "0" AND lcl_latitude_mobileoption <> "0.00" then
        lcl_latitude  = replace(lcl_latitude_mobileoption,",","")
        lcl_longitude = replace(lcl_longitude_mobileoption,",","")
     else
        lcl_latitude  = lcl_latitude_actionline
        lcl_longitude = lcl_longitude_actionline
     end if

     if lcl_latitude = "" OR lcl_latitude = "0" then
        lcl_latitude = "0.00"
     end if

     if lcl_longitude = "" OR lcl_longitude = "0" then
        lcl_longitude = "0.00"
     end if

     if lcl_latitude <> "0.00" then
        if lcl_mappoints_info <> "" then
           lcl_mappoints_info = lcl_mappoints_info & ", " & vbcrlf
        end if

       'Determine which type of mappoint information to build and return
        if lcl_infoType = "BUBBLEINFO" then
           lcl_display_location = "<strong>Location: </strong>" & trim(oMapPoints("streetname"))
           lcl_display_latlng   = "<strong>Latitude: </strong>" & lcl_latitude & "<br /><strong>Longitude: </strong>" & lcl_longitude

           if left(trim(oMapPoints("streetname")),9) <> "Latitude:" then
              if trim(oMapPoints("streetname")) <> "" then
                 lcl_display_location = lcl_display_location & "<br />" & lcl_display_latlng
              else
                 lcl_display_location = lcl_display_latlng
              end if
           else
              lcl_display_location = lcl_display_latlng
           end if

           lcl_mappoints_info = lcl_mappoints_info & "'<div>"
           lcl_mappoints_info = lcl_mappoints_info & "<strong>Tracking Number: </strong>" & oMapPoints("action_autoid") & replace(FormatDateTime(oMapPoints("submit_date"),4),":","") & "<br />"
           lcl_mappoints_info = lcl_mappoints_info & "<strong>Ticket Name: </strong>" & oMapPoints("action_formTitle") & "<br />"
           lcl_mappoints_info = lcl_mappoints_info & lcl_display_location
           lcl_mappoints_info = lcl_mappoints_info & "</div>'"
           '<div><strong>Business Name</strong>: Design Mill<br /><strong>Business Address</strong>: 7842 COOPER ROAD<br /><strong>Website</strong>: <a href="http://www.design-mill.com" target="_blank">www.design-mill.com</a><br /><a href="datamgr_info.asp?f=&dm=9">[more details...]</a></div>', 
           '<div><strong>Business Name</strong>: Woodhouse Day Spa<br /><strong>Business Address</strong>: 9370 MONTGOMERY ROAD<br /><strong>Website</strong>: <a href="http://woodhousecincinnati.rtrk.com" target="_blank">woodhousecincinnati.rtrk.com</a><br /><a href="datamgr_info.asp?f=&dm=10">[more details...]</a></div>'
        else
           lcl_mappoints_info = lcl_mappoints_info & "new google.maps.LatLng(" & lcl_latitude & ", " & lcl_longitude & ")"
           'lcl_mappoints_info = lcl_mappoints_info & "Latitude-ActionLine: [" & lcl_latitude_actionline & "] - Latitude-MobileOption: [" & lcl_latitude_mobileoption & "] - Longitude-ActionLine: [" & lcl_longitude_actionline & "] - Longitude-MobileOption: [" & lcl_longitude_mobileoption & "] - latitude: [" & lcl_latitude & "] - longitude: [" & lcl_longitude & "]"
        end if
     end if

     oMapPoints.movenext

  loop

  oMapPoints.close
  set oMapPoints = nothing

  if lcl_mappoints_info <> "" then
     lcl_return = lcl_mappoints_info
  end if

  getMapPointsInfo = lcl_return

end function
%>
