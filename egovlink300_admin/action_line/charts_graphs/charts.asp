<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="charts_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: charts.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Allows the user to setup and access ASP.NET charts and graphs
'
' MODIFICATION HISTORY
' 1.0 09/01/2010 David Boyer - Initial Version
' 1.1 09/01/2011 David Boyer - Added "exclude auto-resolved requests" checkbox option
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

'BEGIN: Setup the search criteria fields --------------------------------------
 lcl_today                         = Date()
 lcl_chartid                       = 0
 lcl_charturl                      = ""
 lcl_sc_fromDate                   = request("fromDate")
 lcl_sc_toDate                     = request("toDate")
 lcl_sc_includedates               = "N"
 lcl_sc_includedates_checked       = ""
 lcl_sc_selectedchart              = ""
 lcl_sc_charttype                  = "pie"
 lcl_sc_showlegend                 = ""
 lcl_sc_legendtitle                = ""
 lcl_sc_chartwidth                 = "800"
 lcl_sc_chartheight                = "500"
 lcl_sc_collectedthreshold         = "5"
 lcl_sc_collectedlabel             = "Other"
 lcl_sc_formtype                   = ""
 lcl_sc_exclude_autoresolved       = ""

 lcl_selected_formtype_all         = ""
 lcl_selected_formtype_public      = ""
 lcl_selected_formtype_internal    = ""

 lcl_selected_charttype1           = ""
 lcl_selected_charttype2           = ""
 lcl_selected_charttype3           = ""
 lcl_selected_charttype4           = ""
 lcl_selected_charttype5           = ""
 lcl_selected_charttype6           = ""
 lcl_selected_charttype7           = ""
 lcl_selected_charttype8           = ""
 lcl_selected_charttype9           = ""

 lcl_checked_exclude_autoresolved  = ""
 lcl_checked_showlegend            = ""
 lcl_checked_showlegend_javascript = "false"


 if request("cid") <> "" then
    lcl_chartid = clng(request("cid"))
 end if

'If empty then default to the current date
 if lcl_sc_toDate = "" OR IsNull(lcl_sc_toDate) then
    lcl_sc_toDate = lcl_today
 end if

 if lcl_sc_fromDate = "" OR IsNull(lcl_sc_fromDate) then
    lcl_sc_fromDate = cdate(Month(lcl_today)& "/1/" & Year(lcl_today))
 end if

 if request("sc_includedates") <> "" then
    lcl_sc_includedates = request("sc_includedates")

    if lcl_sc_includedates = "Y" then
       lcl_sc_includedates_checked = " checked=""checked"""
    end if
 end if

 if request("sc_selectedchart") <> "" then
    lcl_sc_selectedchart = request("sc_selectedchart")
 end if

'Determine if the charttype is "selected" or not
 lcl_selected_charttype1 = isChartTypeSelected("1",lcl_sc_selectedchart)
 lcl_selected_charttype2 = isChartTypeSelected("2",lcl_sc_selectedchart)
 lcl_selected_charttype3 = isChartTypeSelected("3",lcl_sc_selectedchart)
 lcl_selected_charttype4 = isChartTypeSelected("4",lcl_sc_selectedchart)
 lcl_selected_charttype5 = isChartTypeSelected("5",lcl_sc_selectedchart)
 lcl_selected_charttype6 = isChartTypeSelected("6",lcl_sc_selectedchart)
 lcl_selected_charttype7 = isChartTypeSelected("7",lcl_sc_selectedchart)
 lcl_selected_charttype8 = isChartTypeSelected("8",lcl_sc_selectedchart)
 lcl_selected_charttype9 = isChartTypeSelected("9",lcl_sc_selectedchart)

 if request("sc_charttype") <> "" then
    lcl_sc_charttype = request("sc_charttype")
 end if

 if request("sc_showlegend") <> "" then
    lcl_sc_showlegend = request("sc_showlegend")

    if lcl_sc_showlegend = "Y" then
       lcl_checked_showlegend            = " checked=""checked"""
       lcl_checked_showlegend_javascript = "true"
       lcl_sc_legendtitle                = "Legend"
    end if
 end if

 if request("sc_legendtitle") <> "" then
    lcl_sc_legendtitle = request("sc_legendtitle")
 end if

 if request("sc_chartwidth") <> "" then
    lcl_sc_chartwidth = request("sc_chartwidth")
 end if

 if request("sc_chartheight") <> "" then
    lcl_sc_chartheight = request("sc_chartheight")
 end if

 if request("sc_collectedthreshold") <> "" then
    lcl_sc_collectedthreshold = request("sc_collectedthreshold")
 end if

 if request("sc_collectedlabel") <> "" then
    lcl_sc_collectedlabel = request("sc_collectedlabel")
 end if

 if request("sc_formtype") <> "" then
    lcl_sc_formtype = request("sc_formtype")
 end if

 if request("sc_formtype") <> "" then
    lcl_sc_formtype = ucase(request("sc_formtype"))
 end if

 if lcl_sc_formtype = "INTERNAL" then
    lcl_selected_formtype_internal = " selected=""selected"""
 elseif lcl_sc_formtype = "PUBLIC" then
    lcl_selected_formtype_public = " selected=""selected"""
 else
    lcl_selected_formtype_all = " selected=""selected"""
 end if

 if request("sc_exclude_autoresolved") = "Y" then
     lcl_sc_exclude_autoresolved      = request("sc_exclude_autoresolved")
     lcl_checked_exclude_autoresolved = " checked=""checked"""
 end if

'Check to see if we are to display a chart/graph
 if request.ServerVariables("REQUEST_METHOD") = "POST" then
    getChartURL session("orgid"), _
                session("userid"), _
                lcl_sc_fromDate, _
                lcl_sc_toDate, _
                lcl_sc_includeDates, _
                lcl_sc_selectedchart, _
                lcl_sc_charttype, _
                lcl_sc_chartwidth, _
                lcl_sc_chartheight, _
                lcl_sc_showlegend, _
                lcl_sc_legendtitle, _
                lcl_sc_collectedthreshold, _
                lcl_sc_collectedlabel, _
                lcl_sc_collectedlabel, _
                lcl_sc_formtype, _
                lcl_sc_exclude_autoresolved, _
                lcl_chartid, _
                lcl_charturl
 else
    lcl_sc_showlegend                 = "Y"
    lcl_sc_legendtitle                = "Legend"
    lcl_checked_showlegend            = "Y"
    lcl_checked_showlegend_javascript = "true"
 end if

 lcl_charturl = replace(lcl_charturl,"http:","https:")
                             
'END: Setup the search criteria fields ----------------------------------------

'Check to see if the tab index has been declared
 if request("ti") <> "" then
    lcl_tabindex = clng(request("ti"))
 else
    lcl_tabindex = 0
 end if
%>
<html>
<head>
  <title>E-Gov Administration Consule {Action Line - Charts}</title>
	
	<link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../yui/build/tabview/assets/skins/sam/tabview.css" />

	<script language="javascript" src="../../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../../scripts/getdates.js"></script>
 <script language="javascript" src="../../scripts/formvalidation_msgdisplay.js"></script>

	<script type="text/javascript" src="../../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../../yui/element-min.js"></script>  
	<script type="text/javascript" src="../../yui/tabview-min.js"></script>

<script language="javascript">
<!--
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
<% response.write "			tabView.set('activeIndex', " & lcl_tabindex & ");" %>
//			tabView.set('activeIndex', 0); 
		})();

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("../calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
}

function changeRowColor(pID,pStatus) {
  if(pStatus=="OVER") {
     document.getElementById(pID).style.cursor          = "hand";
     document.getElementById(pID).style.backgroundColor = "#93bee1";
  }else{
     document.getElementById(pID).style.cursor          = "";
     document.getElementById(pID).style.backgroundColor = "";
  }
}

function enableDisableLegendTitle() {
  lcl_disable_legendtitle = true;

  if(document.getElementById("sc_showlegend").checked) {
     lcl_disable_legendtitle = false;
  }

  document.getElementById("sc_legendtitle").disabled = lcl_disable_legendtitle;

}

function printChart(iCID) {
  lcl_width  = 1000;
  lcl_height = 600;
  lcl_left   = ((screen.AvailWidth/2)-(lcl_width/2));
  lcl_top    = ((screen.AvailHeight/2)-(lcl_height/2));

  chartURL  = "print_chart.asp";
  chartURL += "?cid=" + iCID;

  eval('window.open("' + chartURL + '", "_chart' + iCID + '", "width=' + lcl_width + ',height=' + lcl_height + ',left=' + lcl_left + ',top=' + lcl_top + ',toolbar=1,statusbar=0,scrollbars=1,menubar=1,resizable=1")');
}

function checkChartTypes() {
  //Determine which chart types are available depending on the chart/graph selected.
  var lcl_chartselected = document.getElementById("sc_selectedchart").value;

  //Setup NOT "pie" chart options  (i.e. "bar" and "column" graphs)
<%
 'Build the "IF" statement
  lcl_charts = ""
  lcl_charts = lcl_charts &    "lcl_chartselected == ""4"""
  lcl_charts = lcl_charts & "|| lcl_chartselected == ""5"""
  lcl_charts = lcl_charts & "|| lcl_chartselected == ""6"""
  lcl_charts = lcl_charts & "|| lcl_chartselected == ""7"""
  lcl_charts = lcl_charts & "|| lcl_chartselected == ""8"""
  lcl_charts = lcl_charts & "|| lcl_chartselected == ""9"""
%>
  if (<%=lcl_charts%>) {
      if (document.getElementById("sc_charttype").style.display == "none" || document.getElementById("sc_charttype").style.display == "") {

          document.getElementById("sc_collectedthreshold").value = "";
          document.getElementById("sc_collectedlabel").value     = "";
          document.getElementById("sc_legendtitle").value        = "<%=lcl_sc_legendtitle%>";
          document.getElementById("sc_showlegend").checked       = <%=lcl_checked_showlegend_javascript%>;

          document.getElementById("sc_charttype").style.display          = "block";
          document.getElementById("sc_showlegend").style.display         = "block";
          document.getElementById("sc_legendtitle").style.display        = "block";
          document.getElementById("sc_collectedthreshold").style.display = "none";
          document.getElementById("sc_collectedlabel").style.display     = "none";

          //document.getElementById("sc_charttype_label").innerHTML          = "Chart Type:";
          document.getElementById("sc_charttype_display").innerHTML        = "";
          document.getElementById("sc_showlegend_label").innerHTML         = "Show Legend:";
          document.getElementById("sc_legendtitle_label").innerHTML        = "Legend Title:";
          document.getElementById("sc_collectedthreshold_label").innerHTML = "&nbsp;";
          document.getElementById("sc_collectedlabel_label").innerHTML     = "&nbsp;";

          lcl_charttype_length = document.getElementById("sc_charttype").length;

          buildChartTypesList("bar");

          //Determine which option is selected
          for (i=0;i<lcl_charttype_length;i++) {
               if (document.getElementById("sc_charttype")[i].value == "<%=lcl_sc_charttype%>") {
                   document.getElementById("sc_charttype")[i].selected = true;
               }
          }
       }

  //Setup "pie" chart options
  }else if (lcl_chartselected == "3") {
      buildChartTypesList("pie");

      document.getElementById("sc_showlegend").checked       = <%=lcl_checked_showlegend_javascript%>;
      document.getElementById("sc_charttype").value          = "pie";
      document.getElementById("sc_legendtitle").value        = "";
      document.getElementById("sc_collectedthreshold").value = "<%=lcl_sc_collectedthreshold%>";
      document.getElementById("sc_collectedlabel").value     = "<%=lcl_sc_collectedlabel%>";

      document.getElementById("sc_charttype").style.display          = "none";
      document.getElementById("sc_showlegend").style.display         = "none";
      document.getElementById("sc_legendtitle").style.display        = "none";
      document.getElementById("sc_collectedthreshold").style.display = "block";
      document.getElementById("sc_collectedlabel").style.display     = "block";

      //document.getElementById("sc_charttype_label").innerHTML          = "&nbsp;";
      document.getElementById("sc_charttype_display").innerHTML        = "Pie Chart";
      document.getElementById("sc_charttype_display").style.color  = '#800000';
      document.getElementById("sc_collectedthreshold_label").innerHTML = "Group Results with Minimum Value:";
      document.getElementById("sc_collectedlabel_label").innerHTML     = "Minimum Value Label:";
      document.getElementById("sc_showlegend_label").innerHTML         = "&nbsp;";
      document.getElementById("sc_legendtitle_label").innerHTML        = "&nbsp;";
  }
}

function buildChartTypesList(iType) {
  var newElem = document.getElementById("sc_charttype");
  newElem.options.length = 0;

  if(iType != "pie") {
     newElem.options[0] = new Option("Bar Graph","bar");
     newElem.options[1] = new Option("Column Graph","column");
  } else {
     newElem.options[0] = new Option("Pie Chart","pie");
  }
}

function viewChart() {
		var daterege         = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
  var lcl_return_false = 0;
  var sIncludeDates    = "N";
  var sShowLegend      = "N";

  if (document.getElementById("fromDate").value != "") {
    		var dateFromOk = daterege.test(document.getElementById("fromDate").value);

    		if (! dateFromOk ) {
          document.getElementById("fromDate").focus();
          inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
          lcl_return_false = lcl_return_false + 1;
      } else {
          clearMsg("fromDateCalPop");
      }
  } else {
      document.getElementById("fromDate").focus();
      inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Required Field Missing: </strong> From Date',10,'fromDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }

  if (document.getElementById("toDate").value != "") {
    		var dateToOk = daterege.test(document.getElementById("toDate").value);

    		if (! dateToOk ) {
          document.getElementById("toDate").focus();
          inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
          lcl_return_false = lcl_return_false + 1;
      } else {
          clearMsg("toDateCalPop");
      }
  } else {
      document.getElementById("toDate").focus();
      inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Required Field Missing: </strong> To Date',10,'toDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }

  if (document.getElementById("sc_collectedthreshold").style.display) {
      if (document.getElementById("sc_collectedthreshold").value != "") {
          if (! Number(document.getElementById("sc_collectedthreshold").value) && document.getElementById("sc_collectedthreshold").value != "0") {
              document.getElementById("sc_collectedthreshold").focus();
              inlineMsg(document.getElementById("sc_collectedthreshold").id,'<strong>Invalid Value: </strong> "Group Results with Minimum Value" must be numeric.',10,'sc_collectedthreshold');
              lcl_return_false = lcl_return_false + 1;
          }else if(document.getElementById("sc_collectedthreshold").value <= 0) {
              document.getElementById("sc_collectedthreshold").focus();
              inlineMsg(document.getElementById("sc_collectedthreshold").id,'<strong>Invalid Value: </strong> "Group Results with Minimum Value" must be greater than zero (0).',10,'sc_collectedthreshold');
              lcl_return_false = lcl_return_false + 1;
          } else {
              clearMsg("sc_collectedthreshold");
          }
      } else {
          clearMsg("sc_collectedthreshold");
      }
  }

  if(lcl_return_false > 0) {
     return false;
  } else {

//     if(document.getElementById("sc_includedates").checked) {
//        sIncludeDates = "Y";
//     }

//     if(document.getElementById("sc_showlegend").checked) {
//        sShowLegend = "Y";
//     }

//     var sParameter = 'isAjaxRoutine=Y';
//     sParameter    += '&orgid='               + encodeURIComponent('<%=session("orgid")%>');
//     sParameter    += '&userid='              + encodeURIComponent('<%=session("userid")%>');
//     sParameter    += '&fromDate='            + encodeURIComponent(document.getElementById("fromDate").value);
//     sParameter    += '&toDate='              + encodeURIComponent(document.getElementById("toDate").value);
//     sParameter    += '&includedates='        + sIncludeDates;
//     sParameter    += '&selectedchart='       + encodeURIComponent(document.getElementById("sc_selectedchart").value);
//     sParameter    += '&charttype='           + encodeURIComponent(document.getElementById("sc_charttype").value);
//     sParameter    += '&chartwidth='          + encodeURIComponent(document.getElementById("sc_chartwidth").value);
//     sParameter    += '&chartheight='         + encodeURIComponent(document.getElementById("sc_chartheight").value);
//     sParameter    += '&showlegend='          + sShowLegend;
//     sParameter    += '&legendtitle='         + encodeURIComponent(document.getElementById("sc_legendtitle").value);
//     sParameter    += '&collectedthreshold='  + encodeURIComponent(document.getElementById("sc_collectedthreshold").value);
//     sParameter    += '&collectedlabel='      + encodeURIComponent(document.getElementById("sc_collectedlabel").value);
//     sParameter    += '&collectedlegendtext=' + encodeURIComponent(document.getElementById("sc_collectedlabel").value);

//     document.getElementById("chartsGraphsResults").width  = document.getElementById("sc_chartwidth").value;
//     document.getElementById("chartsGraphsResults").height = document.getElementById("sc_chartheight").value;

//     doAjax('getchartgraph.asp', sParameter, 'displayChart', 'post', '0');
     document.getElementById("chartGraphSearchCriteria").action = "charts.asp";
     document.getElementById("chartGraphSearchCriteria").submit();
  }

}

function displayChart(p_chart_src) {
  document.getElementById("chartsGraphsResults").src = p_chart_src;
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" class="yui-skin-sam" onload="checkChartTypes();">

<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf

  response.write "<p><font size=""+1""><strong>Action Line: Charts</strong></font></p>" & vbcrlf

 'BEGIN: Tabs (headers) -------------------------------------------------------
  response.write "<div id=""demo"" class=""yui-navset"" style=""width:804px"">" & vbcrlf
  response.write "  <ul class=""yui-nav"">" & vbcrlf
  response.write "  		<li><a href=""#tab1""><em>Charts</em></a></li>" & vbcrlf
  response.write "  		<li><a href=""#tab2""><em>History Log</em></a></li>" & vbcrlf
  response.write "  </ul>" & vbcrlf
  response.write "<div class=""yui-content"">" & vbcrlf
 'END: Tabs (headers) ---------------------------------------------------------

 'BEGIN: Chart/Graphs Results -------------------------------------------------
  'lcl_charturl = "http://dev4.egovlink.com/eclink/admin/charts/piechart.aspx?cid=6"
  'lcl_charturl = ""

  response.write "<form name=""chartGraphSearchCriteria"" id=""chartGraphSearchCriteria"" method=""POST"" action="""">" & vbcrlf
  response.write "  <input type=""hidden"" name=""cid"" id=""cid"" value=""" & lcl_chartid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "<div id=""tab1"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""start"" style=""width:800px;"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "              <tr valign=""top"">" & vbcrlf
  response.write "                  <td nowrap=""nowrap"">Request Submission Date:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & lcl_sc_fromDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('fromDateCalPop');"" />" & vbcrlf
  response.write "                      <a href=""javascript:void doCalendar('From');""><img src=""../../images/calendar.gif"" id=""fromDateCalPop"" border=""0"" style=""cursor:pointer"" onclick=""clearMsg('fromDateCalPop');"" /></a>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      To: " & vbcrlf
  response.write "                      <input type=""text"" name=""toDate"" id=""toDate"" value=""" & lcl_sc_toDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('toDateCalPop');"" />" & vbcrlf
  response.write "                      <a href=""javascript:void doCalendar('To');""><img src=""../../images/calendar.gif"" id=""toDateCalPop"" border=""0"" style=""cursor:pointer"" onclick=""clearMsg('toDateCalPop');"" /></a>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        DrawDateChoices "Date",""
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr valign=""top"">" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""checkbox"" name=""sc_includedates"" id=""sc_includedates"" value=""Y""" & lcl_sc_includedates_checked & " />Include search dates in chart/graph title" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      Form Types:" & vbcrlf
  response.write "                      <select name=""sc_formtype"" id=""sc_formtype"">" & vbcrlf
  response.write "                        <option value="""""         & lcl_selected_formtype_all      & ">&nbsp;</option>" & vbcrlf
  response.write "                        <option value=""PUBLIC"""   & lcl_selected_formtype_public   & ">Public Only</option>" & vbcrlf
  response.write "                        <option value=""INTERNAL""" & lcl_selected_formtype_internal & ">Internal Only</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend>Chart/Graph Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td nowrap=""nowrap"">Charts:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sc_selectedchart"" id=""sc_selectedchart"" onchange=""checkChartTypes()"">" & vbcrlf
  'response.write "                        <option value=""1""" & lcl_selected_charttype1 & ">Monthly Status</option>" & vbcrlf
  'response.write "                        <option value=""2""" & lcl_selected_charttype2 & ">Monthly Status by Department</option>" & vbcrlf
  response.write "                        <option value=""3""" & lcl_selected_charttype3 & ">Most Submitted Requests</option>" & vbcrlf
  response.write "                        <option value=""6""" & lcl_selected_charttype6 & ">Open Items by Department</option>" & vbcrlf
  response.write "                        <option value=""5""" & lcl_selected_charttype5 & ">Open Items Activity by Department (Avg Days Open)</option>" & vbcrlf
  response.write "                        <option value=""8""" & lcl_selected_charttype8 & ">Open Items Activity by Department (Avg Days Last Activity)</option>" & vbcrlf
  response.write "                        <option value=""7""" & lcl_selected_charttype7 & ">Open Items by Form</option>" & vbcrlf
  response.write "                        <option value=""4""" & lcl_selected_charttype4 & ">Open Items Activity by Form (Avg Days Open)</option>" & vbcrlf
  response.write "                        <option value=""9""" & lcl_selected_charttype9 & ">Open Items Activity by Form (Avg Days Last Activity)</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td id=""sc_charttype_label"">Chart Type:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <span id=""sc_charttype_display""></span>" & vbcrlf
  response.write "                      <select name=""sc_charttype"" id=""sc_charttype"">" & vbcrlf
  response.write "                        <option value=""pie"">Pie Chart</option>" & vbcrlf
  response.write "                        <option value=""bar"">Bar Graph</option>" & vbcrlf
  response.write "                        <option value=""column"">Column Graph</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Chart Width:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sc_chartwidth"" id=""sc_chartwidth"">" & vbcrlf
                                          lcl_options_width = 100

                                          do while lcl_options_width < 1001
                                             if lcl_options_width = clng(lcl_sc_chartwidth) then
                                                lcl_selected_width = " selected=""selected"""
                                             else
                                                lcl_selected_width = ""
                                             end if

                                             response.write "<option value=""" & lcl_options_width & """" & lcl_selected_width & ">" & lcl_options_width & "</option>" & vbcrlf

                                             lcl_options_width = lcl_options_width + 100
                                          loop
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td id=""sc_showlegend_label"">Show Legend:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""checkbox"" name=""sc_showlegend"" id=""sc_showlegend"" value=""Y""" & lcl_checked_showlegend & " />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Chart Height:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sc_chartheight"" id=""sc_chartheight"">" & vbcrlf
                                          lcl_options_height = 100

                                          do while lcl_options_height < 1001
                                             if lcl_options_height = clng(lcl_sc_chartheight) then
                                                lcl_selected_height = " selected=""selected"""
                                             else
                                                lcl_selected_height = ""
                                             end if

                                             response.write "<option value=""" & lcl_options_height & """" & lcl_selected_height & ">" & lcl_options_height & "</option>" & vbcrlf

                                             lcl_options_height = lcl_options_height + 100
                                          loop
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td id=""sc_legendtitle_label"">Legend Title:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_legendtitle"" id=""sc_legendtitle"" value=""" & lcl_sc_legendtitle & """ size=""30"" maxlength=""50"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr valign=""top"">" & vbcrlf
  response.write "                  <td colspan=""2"" style=""padding-left:6px"">" & vbcrlf
  response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td id=""sc_collectedthreshold_label"" nowrap=""nowrap"">Group Results with Minimum Value:</td>" & vbcrlf
  response.write "                            <td width=""100%"" style=""padding-left:2px;""><input type=""text"" name=""sc_collectedthreshold"" id=""sc_collectedthreshold"" value=""" & lcl_sc_collectedthreshold & """ size=""4"" maxlength=""5"" onchange=""clearMsg('sc_collectedthreshold');"" /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td colspan=""2"">" & vbcrlf
  response.write "                                <input type=""checkbox"" name=""sc_exclude_autoresolved"" id=""sc_exclude_autoresolved"" value=""Y""" & lcl_checked_exclude_autoresolved & """ />&nbsp;Exclude Auto-Resolved Requests" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                      </table>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td nowrap=""nowrap"" id=""sc_collectedlabel_label"">Minimum Value Label:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_collectedlabel"" id=""sc_collectedlabel"" value=""" & lcl_sc_collectedlabel & """ size=""30"" maxlength=""50"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:790px; margin-top:5pt; margin-bottom:5pt; margin-left:5pt;"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td><input type=""button"" name=""viewChartButton"" id=""viewChartButton"" value=""View Chart"" class=""button"" onclick=""viewChart();"" /></td>" & vbcrlf
  response.write "                <td align=""right"">&nbsp;" & vbcrlf

  if CLng(lcl_chartid) > 0 then
     response.write "                 <input type=""button"" name=""printChartButton"" id=""printChartButton"" value=""Print Chart"" class=""button"" onclick=""printChart('" & lcl_chartid & "');"" />" & vbcrlf
  end if

  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  if lcl_charturl <> "" then
     lcl_frame_width  = lcl_sc_chartwidth  + 20
     lcl_frame_height = lcl_sc_chartheight + 20

     response.write "<p>" & vbcrlf
     'response.write "<iframe name=""chartsGraphsResults"" id=""chartsGraphsResults"" src=""" & lcl_charturl & """ width=""" & lcl_sc_chartwidth & """ height=""" & lcl_sc_chartheight & """ marginwidth=""0"" marginheight=""0"" hspace=""0"" vspace=""0"" frameborder=""0"" scrolling=""YES"" bordercolor=""#ff0000"">You will not see this text if your browser supports IFRAME.</iframe>" & vbcrlf
     response.write "<iframe name=""chartsGraphsResults"" id=""chartsGraphsResults"" src=""" & lcl_charturl & """ width=""" & lcl_frame_width & """ height=""" & lcl_frame_height & """ marginwidth=""0"" marginheight=""0"" hspace=""0"" vspace=""0"" frameborder=""0"" scrolling=""0"" bordercolor=""#ff0000"">You will not see this text if your browser supports IFRAME.</iframe>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if

  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
 'END: Chart/Graphs Results ---------------------------------------------------

 'BEGIN: History Log ----------------------------------------------------------
  response.write "<div id=""tab2"">" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"">" & vbcrlf
  response.write "  <tr align=""left"">" & vbcrlf
  response.write "      <th>Chart</th>" & vbcrlf
  response.write "      <th align=""center"">Chart Type</th>" & vbcrlf
  response.write "      <th>Date Run</th>" & vbcrlf
  response.write "  <tr>" & vbcrlf

  sSQL = "SELECT "
  sSQL = sSQL & " chartid, "
  sSQL = sSQL & " charttype, "
  sSQL = sSQL & " charttitle, "
  sSQL = sSQL & " dateadded, "
  sSQL = sSQL & " createdby "
  sSQL = sSQL & " FROM egov_charts "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND createdby = " & session("userid")
  sSQL = sSQL & " ORDER BY dateadded DESC "

 	set oChartLog = Server.CreateObject("ADODB.Recordset")
	 oChartLog.Open sSQL, Application("DSN"), 3, 1

  lcl_bgcolor = "#ffffff"

  if not oChartLog.eof then
     do while not oChartLog.eof

        lcl_bgcolor         = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_row_onmouseover = " onmouseover=""changeRowColor('row_" & oChartLog("chartid") & "','OVER')"""
        lcl_row_onmouseout  = " onmouseout=""changeRowColor('row_" & oChartLog("chartid") & "','OUT')"""
        lcl_row_onclick     = " onclick=""printChart('" & oChartLog("chartid") & "');"""

        lcl_display_charttitle = ""

        if oChartLog("charttitle") <> "" then
           lcl_display_charttitle = replace(oChartLog("charttitle"),"\n", " ")
        end if

        response.write "  <tr id=""row_" & oChartLog("chartid") & """ bgcolor=""" & lcl_bgcolor & """" & lcl_row_onmouseover & lcl_row_onmouseout & ">" & vbcrlf
        response.write "      <td" & lcl_row_onclick & ">" & lcl_display_charttitle & "</td>" & vbcrlf
        response.write "      <td" & lcl_row_onclick & " align=""center"">" & oChartLog("charttype") & "</td>" & vbcrlf
        response.write "      <td" & lcl_row_onclick & ">" & oChartLog("dateadded")  & "</td>" & vbcrlf
        response.write "  <tr>" & vbcrlf

        oChartLog.movenext
     loop
  end if

  oChartLog.close
  set oChartLog = nothing

  response.write "</table>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: History Log ------------------------------------------------------------

  response.write "</div>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
