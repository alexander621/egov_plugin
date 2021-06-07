<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="charts_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: print_chart.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Allows the user to open a chart in a print-mode screen
'
' MODIFICATION HISTORY
' 1.0 09/16/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../../admin/outage_feature_offline.asp"
 end if

 sLevel = "../../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"actionline_chartsgraphs") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 set oChartOrg = New classOrganization
 oChartOrg.SetOrgId(session("orgid"))

 lcl_chartid     = ""
 lcl_charttype   = ""
 lcl_chartwidth  = "800"
 lcl_chartheight = "500"

 if request("cid") <> "" then
    lcl_chartid = clng(request("cid"))
 end if

'Get the chart data
 sSQL = "SELECT charttype, chartwidth, chartheight "
 sSQL = sSQL & " FROM egov_charts "
 sSQL = sSQL & " WHERE chartid = " & lcl_chartid

 set oGetChartInfo = Server.CreateObject("ADODB.Recordset")
	oGetChartInfo.Open sSQL, Application("DSN"), 3, 1

 if not oGetChartInfo.eof then
    lcl_charttype   = oGetChartInfo("charttype")
    lcl_chartwidth  = oGetChartInfo("chartwidth")
    lcl_chartheight = oGetChartInfo("chartheight")

    if lcl_charttype = "bar" OR lcl_charttype = "column" then
       lcl_pageurl = "barcolumnchart"
    else
       lcl_pageurl = "piechart"
    end if

    lcl_org_sitename = getOrgVirtualSiteName(session("orgid"))
    lcl_chart_url    = Application("charts_graphs_url") & "/" & lcl_org_sitename & "/admin/charts/" & lcl_pageurl & ".aspx?cid=" & lcl_chartid
 end if

 oGetChartInfo.close
 set oGetChartInfo = nothing
%>
<html>
<head>
  <title>E-Gov Administration Consule {Action Line - Charts and Graphs}</title>
	
 	<link rel="stylesheet" type="text/css" href="../../global.css" />

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<div id="content">
  <div id="centercontent">
<%
 'BEGIN: Display Chart/Graph --------------------------------------------------
  response.write "<div>" & vbcrlf

  if lcl_chart_url <> "" then
     lcl_frame_width  = lcl_chartwidth  + 20
     lcl_frame_height = lcl_chartheight + 20

     response.write "<p>" & vbcrlf
     response.write "<iframe name=""chartsGraphsResults"" id=""chartsGraphsResults"" src=""" & lcl_chart_url & """ width=""" & lcl_frame_width & """ height=""" & lcl_frame_height & """ marginwidth=""0"" marginheight=""0"" hspace=""0"" vspace=""0"" frameborder=""0"" scrolling=""0"" bordercolor=""#ff0000"">You will not see this text if your browser supports IFRAME.</iframe>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if

  response.write "</div>" & vbcrlf
 'END: Display Chart/Graph ----------------------------------------------------
%>
  </div>
</div>

</body>
</html>
