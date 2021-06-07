<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: customreports.asp
' AUTHOR:   David Boyer
' CREATED:  11/18/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  11/18/08  David Boyer - Initial Version
'                              * Initial Report: Code Sections
' 1.1 12/08/08 David Boyer - Added "Subscriber List" export
' 1.2 01/05/09 David Boyer	- Added "PDF Field Names" (Action Line Form Creator)
' 1.3 01/13/09 David Boyer - Added "Public Requests Created per Organization" (requested by Peter)
' 1.4 01/14/08 David Boyer - Added "Team Roster"
' 1.5 01/21/09 David Boyer - Added "Help Documentation"
' 1.6 02/05/09 David Boyer - Added "Membership Card Scans" report/export
' 1.7 02/17/09 David Boyer - Added "Print" and "Print Preview" buttons
' 1.8 05/21/09 David Boyer - Added "Click Counts" report/export
' 1.9 11/30/09 David Boyer - Modified "Team Roster" report.  Added "Pants Size".
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("customreports") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'Determine which custom report to run
 lcl_customreport       = ""
 lcl_cr_title           = ""
 lcl_codesection_ids    = ""
 lcl_requestids         = ""
 lcl_userpermission     = ""
 lcl_showmenu           = "N"
 lcl_show_totalrecords  = "N"
 lcl_totalrecords_label = ""
 lcl_export_parameters  = ""

 if UCASE(request("CR")) = "CODESECTIONS" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "Code Violations"
    lcl_query           = session("CR_CODESECTIONS")
    lcl_userpermission  = "customreports_codesections"

    lcl_orghasfeature_action_line_substatus     = orghasfeature("action_line_substatus")
    lcl_userhaspermission_action_line_substatus = userhaspermission(session("userid"),"action_line_substatus")

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "SUBSCRIBERLIST" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "Subscribers"
    lcl_query           = session("CR_SUBSCRIBERLIST")
    lcl_userpermission  = "customreports_subscriberlist"

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "PDFFIELDNAMES" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "PDF Field Names"
    lcl_query           = session("CR_PDFFIELDNAMES")
    lcl_userpermission  = ""

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "CLASSEVENTS_TEAMROSTER" AND session("CR_CLASSEVENTS_TEAMROSTER") <> "" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "Team Roster"
    lcl_query           = session("CR_CLASSEVENTS_TEAMROSTER")
    lcl_userpermission  = "customreports_classesevents_teamroster"

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "MAPPOINTS_EXPORT" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_mappoint_typeid = request("mpt")
    lcl_cr_title        = "Map-Points"
    lcl_query           = session("CR_MAPPOINTS_EXPORT")
    lcl_userpermission  = "customreports_mappoints"

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "CLICKCOUNTS" then
    lcl_customreport       = UCASE(request("CR"))
    lcl_cr_title           = "Click Counts"
    lcl_query              = ""
    lcl_userpermission     = "customreports_clickcounts"
    lcl_show_totalrecords  = "Y"
    lcl_totalrecords_label = " Times Clicked"

   'Determine which "clickcount" query to run.
    if session("CR_CLICKCOUNTS") = "POSTINGS" then
       lcl_cr_title    = lcl_cr_title & " (Postings)"
       lcl_posting_id  = request("posting_id")

      'Setup this parameter so that the export will work properly
       lcl_export_parameters  = "&posting_id=" & lcl_posting_id

       if lcl_posting_id <> "" then
          lcl_query = "SELECT c.postings_clickid, "
          lcl_query = lcl_query & " c.orgid, "
          lcl_query = lcl_query & " c.userid, "
          lcl_query = lcl_query & " (select isnull(eu.userlname,'') + ', ' + isnull(eu.userfname,'') "
          lcl_query = lcl_query &  " from egov_users eu "
          lcl_query = lcl_query &  " where eu.userid = c.userid) AS username, "
          lcl_query = lcl_query & " (select isnull(eu2.useremail,'') "
          lcl_query = lcl_query &  " from egov_users eu2 "
          lcl_query = lcl_query &  " where eu2.userid = c.userid) AS useremail, "
          lcl_query = lcl_query & " c.posting_id, "
          lcl_query = lcl_query & " isnull(jb.jobbid_id,'') AS jobbid_id, "
          lcl_query = lcl_query & " isnull(jb.title,'') AS title, "
          lcl_query = lcl_query & " c.clicked_linkid, "
          lcl_query = lcl_query & " c.clicked_linktext, "
          lcl_query = lcl_query & " c.clicked_linkurl, "
          lcl_query = lcl_query & " c.clicked_date, "
          lcl_query = lcl_query & " (select isnull(eu.userbusinessname,'') "
          lcl_query = lcl_query &  " from egov_users eu "
          lcl_query = lcl_query &  " where eu.userid = c.userid) AS userbusinessname, "
          lcl_query = lcl_query & " (select isnull(eu.userworkphone,'') "
          lcl_query = lcl_query &  " from egov_users eu "
          lcl_query = lcl_query &  " where eu.userid = c.userid) AS userworkphone, "
          lcl_query = lcl_query & " (select isnull(eu.userfax,'') "
          lcl_query = lcl_query &  " from egov_users eu "
          lcl_query = lcl_query &  " where eu.userid = c.userid) AS userfax "
          lcl_query = lcl_query & " FROM egov_clickcounter_postings c "
          lcl_query = lcl_query &      " LEFT OUTER JOIN egov_jobs_bids jb ON c.posting_id = jb.posting_id "
          lcl_query = lcl_query & " WHERE c.orgid = "    & session("orgid")
          lcl_query = lcl_query & " AND c.posting_id = " & lcl_posting_id
          lcl_query = lcl_query & " ORDER BY c.clicked_linktext, c.clicked_date DESC, 4 "
       end if
    end if

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "RSSLOG" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "RSS Send Log"
    'lcl_query           = ""
    'lcl_userpermission  = "customreports_subscriberlist"
    lcl_userpermission  = ""
	lcl_export_parameters = "&ID=" & request("ID")

    lcl_query = "SELECT r.rssid, r.rowid, r.title, r.description, r.rsslink, r.publicationdate, r.createdbyid, r.createdbyname, r.feedid "
    lcl_query = lcl_query & " FROM egov_rss r, egov_rssfeeds f "
    lcl_query = lcl_query & " WHERE r.feedid = f.feedid "
    lcl_query = lcl_query & " AND UPPER(f.feedname) = '" & session("RSSType") & "' "
    lcl_query = lcl_query & " AND r.orgid = " & session("orgid")
    lcl_query = lcl_query & " AND r.rowid = " & request("ID")
    lcl_query = lcl_query & " ORDER BY r.publicationdate DESC "

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "HELP_DOCUMENTATION" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "Help Documentation"
    'lcl_query           = session("CR_HELP_DOCUMENTATION")
    lcl_userpermission  = "customreports_helpdocumentation"
    lcl_showmenu        = "Y"

    lcl_query = "SELECT h.helpdocid, h.orgid, h.helpdoc_name, h.documentid, d.documenturl "
    lcl_query = lcl_query & " FROM egov_helpdocs h "
    lcl_query = lcl_query &      " LEFT OUTER JOIN documents d ON h.documentid = d.documentid "
    lcl_query = lcl_query & " ORDER BY helpdoc_name "

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "PUBLICREQUESTSBYORG" then
    lcl_customreport    = UCASE(request("CR"))
    lcl_cr_title        = "Public Action Line Requests by Organization"
    lcl_userpermission  = ""
    'lcl_query          = session("CR_PDFFIELDNAMES")

   'Get the feature name for Action Line
    lcl_featureid = getFeatureID("action line")

    lcl_query = "SELECT distinct f.orgid, o.orgname, "
    lcl_query = lcl_query & " (SELECT count(r.action_autoid) "
    lcl_query = lcl_query & "  FROM egov_action_request_view r "
    lcl_query = lcl_query & "  WHERE r.orgid = f.orgid "
    lcl_query = lcl_query & "  AND (r.employeesubmitid is null OR r.employeesubmitid = '')) as TotalRequests "
    lcl_query = lcl_query & " FROM egov_organizations_to_features f, organizations o "
    lcl_query = lcl_query & " WHERE f.orgid = o.orgid "
    lcl_query = lcl_query & " AND f.featureid = " & lcl_featureid
    lcl_query = lcl_query & " ORDER BY o.orgname "

'------------------------------------------------------------------------------
 elseif UCASE(request("CR")) = "MEMBERSHIPCARD_SCANS_BY_MEMBERID" then
    lcl_customreport       = UCASE(request("CR"))
    lcl_cr_title           = "Membership Card Scan History (by Member ID)"
    lcl_userpermission     = "customreports_membership_scanlog"
    lcl_export_parameters  = "&memberid=" & request("memberid") & "&rateid=" & request("rateid")
    'lcl_query              = session("CR_MEMBERSHIPCARD_SCANS_BY_MEMBERID")

    lcl_query = "SELECT a.scan_datetime, a.memberid, a. poolpassid, "
    lcl_query = lcl_query & "(select distinct f.lastname+ ', ' + f.firstname "
    lcl_query = lcl_query & " from egov_familymembers f "
    lcl_query = lcl_query & " where f.familymemberid = (select distinct m.familymemberid "
    lcl_query = lcl_query &                           " from egov_poolpassmembers m "
    lcl_query = lcl_query &                           " where m.memberid = a.memberid)) AS membername "
    lcl_query = lcl_query & " FROM egov_pool_attendance_log a "
    lcl_query = lcl_query & " WHERE a.orgid = "  & session("orgid")
    lcl_query = lcl_query & " AND a.memberid = " & request("memberid")
    lcl_query = lcl_query & " AND a.rateid = "   & request("rateid")
    lcl_query = lcl_query & " ORDER BY a.scan_datetime "

    lcl_show_totalrecords  = "Y"
    lcl_totalrecords_label = " Times Scanned"

 end if

'Check for the proper permission(s)
 'if lcl_customreport <> "PDFFIELDNAMES" then
 '   if not UserHasPermission(session("userid"),"customreports") then
	'      response.redirect sLevel & "permissiondenied.asp"
 '   else
       if lcl_userpermission <> "" then
          if not UserHasPermission(session("userid"),lcl_userpermission) then
      	      response.redirect sLevel & "permissiondenied.asp"
          end if
       else
         'Reports that do NOT require permissions to access:
         '  - RSSLOG
         '  - PDFFIELDNAMES
         '  - PUBLICREQUESTSBYORG
          if UCASE(request("CR")) <> "RSSLOG" AND UCASE(request("CR")) <> "PDFFIELDNAMES" AND UCASE(request("CR")) <> "PUBLICREQUESTSBYORG" then
      	      response.redirect sLevel & "permissiondenied.asp"
          end if
       end if
 '   end if
 'end if

'Determine if the user has requested an export
 lcl_export = "N"

'Set up the export
 if UCASE(request("export")) = "Y" then
    sFileName            = lcl_customreport & "_" & replace(replace(replace(replace(replace(Now(),":",""),"\","")," ",""),"AM",""),"PM","") & ".xls"
    response.ContentType = "application/vnd.ms-excel"  'or "x-msexcel"
    response.AddHeader "Content-Disposition", "attachment;filename=" & sFileName
    lcl_export = "Y"
 end if

'Display only for non-export --------------------------------------------------
 if lcl_export <> "Y" then
    response.write "<html>" & vbcrlf
    response.write "<head>" & vbcrlf
    response.write " 	<title>E-Gov Administration Console</title>" & vbcrlf

   'Javascript
    response.write "  <script language=""javascript"" src=""tablesort.js""></script>" & vbcrlf
    response.write "  <script language=""javascript"" src=""../scripts/modules.js""></script>" & vbcrlf
    response.write "  <script language=""javascript"">" & vbcrlf
    response.write "    function launchExport() {" & vbcrlf

    if request("iFormID") <> "" then
       lcl_export_parameters = lcl_export_parameters & "&iFormID=" & request("iFormID")
    end if

    if request("mpt") <> "" then
       lcl_export_parameters = lcl_export_parameters & "&mpt=" & request("mpt")
    end if

    response.write "      window.open('customreports.asp?cr=" & lcl_customreport & "&export=Y" & lcl_export_parameters & "', '_customreports_export', 'toolbar=0,statusbar=0,scrollbars=1,menubar=0');" & vbcrlf
    response.write "    }" & vbcrlf
    response.write "  </script>" & vbcrlf

   'Setup "Print" and "Print Preview" buttons
    response.write "  <script defer>" & vbcrlf
    response.write "		window.onload = function() {" & vbcrlf

    'response.write "    factory.printing.printer           = ""Zebra P330i USB Card Printer"";" & vbcrlf
    response.write "    factory.printing.header            = '';"    & vbcrlf
    response.write "    factory.printing.footer            = '';"    & vbcrlf
    response.write "    factory.printing.portrait          = false;" & vbcrlf
    response.write "    factory.printing.leftMargin        = 0.5;"  & vbcrlf
    response.write "    factory.printing.rightMargin       = 0.5;"  & vbcrlf
    response.write "    factory.printing.topMargin         = 0.5;"  & vbcrlf
    response.write "    factory.printing.bottomMargin      = 0.5;"   & vbcrlf

    response.write "    //enable control buttons" & vbcrlf
    response.write "    var templateSupported = factory.printing.IsTemplateSupported();" & vbcrlf
    response.write "    var controls = idControls.all.tags('input');" & vbcrlf
    response.write "    for ( i = 0; i < controls.length; i++ ) {" & vbcrlf
    response.write "       controls[i].disabled = false;" & vbcrlf
    response.write "       if (templateSupported && controls[i].className == 'ie55' )" & vbcrlf
    response.write "           controls[i].style.display = 'inline';" & vbcrlf
    response.write "       }" & vbcrlf
    response.write "  }" & vbcrlf
    response.write "  </script>" & vbcrlf

   'CSS
    response.write " 	<link rel=""stylesheet"" type=""text/css"" href=""../menu/menu_scripts/menu.css"" />" & vbcrlf
    response.write " 	<link rel=""stylesheet"" type=""text/css"" href=""../global.css"" />" & vbcrlf
    response.write " 	<link rel=""stylesheet"" type=""text/css"" href=""pageprint.css"" media=""print"" />" & vbcrlf

    response.write "</head>" & vbcrlf
    response.write "<body>" & vbcrlf

    ShowHeader sLevel

    if lcl_showmenu = "Y" then
  %>
       <!--#Include file="../menu/menu.asp"-->
  <%
    end if

  'BEGIN: Third Party Print Control
   response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
   response.write "  <input disabled type=""button"" value=""Print the page"" onclick=""factory.printing.Print(true)"" />" & vbcrlf
   response.write "  <input class=""ie55"" disabled type=""button"" value=""Print Preview..."" onclick=""factory.printing.Preview()"" />" & vbcrlf
   response.write "</div>" & vbcrlf

   response.write "<object id=""factory"" viewastext  style=""display:none"" classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814"" codebase=""../includes/smsx.cab#Version=6,3,434,12""></object>" & vbcrlf
  'END: Third Party Print Control

    response.write "<div id=""content"">" & vbcrlf
    response.write "	 <div id=""centercontent"">" & vbcrlf
    response.write "<p><font size=""+1""><strong>" & session("sOrgName") & " - " & lcl_cr_title & "</strong></font></p>" & vbcrlf

    displaybuttons

 end if

'------------------------------------------------------------------------------
 sSQL = lcl_query
	set oCustom = Server.CreateObject("ADODB.Recordset")
	oCustom.Open sSQL, Application("DSN"), 0, 1

	if not oCustom.eof then
   'Display only for non-export -----------------------------------------------
    if lcl_export <> "Y" then
       lcl_bgcolor = "#ffffff"
       iRowCount   = 0

       response.write "<p>" & vbcrlf
       response.write "<div class=""shadow"">" & vbcrlf
       response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"">" & vbcrlf
       response.write "  <tr align=""left"" valign=""bottom"">" & vbcrlf
    else
       response.write "<table border=""1"">" & vbcrlf
       response.write "  <tr>" & vbcrlf

       if lcl_customreport = "SUBSCRIBERLIST" then
          lcl_totalcolumns = 16

          response.write "      <td><strong>" & oCustom("categoryname") & "</strong></td>" & vbcrlf
          displayFillerTDs "N",lcl_totalcolumns-1,"Y"
          displayFillerTDs "Y",lcl_totalcolumns,"Y"

          response.write "  <tr>" & vbcrlf
       end if

    end if

   'Build the column headers per report ---------------------------------------
    if lcl_customreport = "CODESECTIONS" then
       response.write "      <th>Tracking Number</th>"  & vbcrlf
       response.write "      <th>Submit Date</th>"      & vbcrlf
       response.write "      <th>Resolved Date</th>"    & vbcrlf
       response.write "      <th>Status</th>"           & vbcrlf

       if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
          response.write "      <th>Sub-Status</th>"       & vbcrlf
       end if

       response.write "      <th>Property Address</th>" & vbcrlf
       response.write "      <th>Code Sections</th>"    & vbcrlf
       if session("orgid") = "209" then
       	response.write "      <th>Assigned</th>"    & vbcrlf
       end if

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "SUBSCRIBERLIST" then
       response.write "      <th>First Name</th>"       & vbcrlf
       response.write "      <th>Last Name</th>"        & vbcrlf
       response.write "      <th>Email</th>"            & vbcrlf
       response.write "      <th>Street Number</th>"    & vbcrlf
       response.write "      <th>Street Prefix</th>"    & vbcrlf
       response.write "      <th>Street Name</th>"      & vbcrlf
       response.write "      <th>City</th>"             & vbcrlf
       response.write "      <th>State</th>"            & vbcrlf
       response.write "      <th>Zip</th>"              & vbcrlf
       response.write "      <th>Business Name</th>"    & vbcrlf
       response.write "      <th>Business<br />Street Number</th>" & vbcrlf
       response.write "      <th>Business Address</th>" & vbcrlf
       response.write "      <th>Home Phone</th>"       & vbcrlf
       response.write "      <th>Cell Phone</th>"       & vbcrlf
       response.write "      <th>Work Phone</th>"       & vbcrlf
       response.write "      <th>Fax</th>"              & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "PUBLICREQUESTSBYORG" then
       response.write "      <th>Clients</th>"                                & vbcrlf
       response.write "      <th>Total Requests Created via Public-side</th>" & vbcrlf
       response.write "      <th align=""center"">OrgID</th>"                 & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "RSSLOG" then
       response.write "      <th>Title</th>"            & vbcrlf
       response.write "      <th>Description</th>"      & vbcrlf
       response.write "      <th>URL</th>"              & vbcrlf
       response.write "      <th>Publication Date</th>" & vbcrlf
       response.write "      <th>Created By</th>"       & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "HELP_DOCUMENTATION" then
       response.write "      <th>HelpDoc Name<br /><i>(used in code)</i></th>" & vbcrlf
       response.write "      <th align=""center"">DocumentID</th>"             & vbcrlf
       response.write "      <th>Filename</th>"                                & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "CLASSEVENTS_TEAMROSTER" then
      'Check for an "edit display" for the T-shirt label
       if orgHasDisplay(session("orgid"),"class_teamregistration_tshirt_label") then
          lcl_label_tshirt = getOrgDisplay(session("orgid"),"class_teamregistration_tshirt_label")
       else
          lcl_label_tshirt = "T-Shirt"
       end if

       response.write "      <th>First Name Enrollee</th>" & vbcrlf
       response.write "      <th>Last Name</th>"           & vbcrlf
       response.write "      <th>Age</th>"                 & vbcrlf
       response.write "      <th>D.O.B.</th>"              & vbcrlf
       response.write "      <th>Grade</th>"               & vbcrlf
       response.write "      <th>" & lcl_label_tshirt & " Size</th>" & vbcrlf
       response.write "      <th>Pants Size</th>"          & vbcrlf
       response.write "      <th>Home Phome</th>"          & vbcrlf
       response.write "      <th>Parent First Name</th>"   & vbcrlf
       response.write "      <th>Parent Last Name</th>"    & vbcrlf
       response.write "      <th>Emergency Phone</th>"     & vbcrlf
       response.write "      <th>Address</th>"             & vbcrlf
       response.write "      <th>Coach Type</th>"          & vbcrlf
       response.write "      <th>Coach Name</th>"          & vbcrlf
       response.write "      <th>Coach Day Phone</th>"     & vbcrlf
       response.write "      <th>Coach Cell Phone</th>"    & vbcrlf
       response.write "      <th>Coach Email</th>"         & vbcrlf

   '------------------------------------------------------------------------------
    elseif lcl_customreport = "MEMBERSHIPCARD_SCANS_BY_MEMBERID" then
       response.write "      <th align=""center"">Member ID</th>" & vbcrlf
       response.write "      <th>Member</th>"    & vbcrlf
       response.write "      <th>Scan Date</th>" & vbcrlf
       response.write "      <th align=""center"">Pool Pass ID</th>" & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "PDFFIELDNAMES" then
       response.write "      <th>Action Line Section</th>" & vbcrlf
       response.write "      <th>Field Name passed to PDF form</th>" & vbcrlf
       response.write "      <th>Value passed to PDF form</th>" & vbcrlf

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "CLICKCOUNTS" then
       if session("CR_CLICKCOUNTS") = "POSTINGS" then
          response.write "      <th>BID ID</th>" & vbcrlf
          response.write "      <th>Title</th>" & vbcrlf
          response.write "      <th>User</th>" & vbcrlf
          response.write "      <th>Business Name</th>" & vbcrlf
          response.write "      <th>Phone</th>" & vbcrlf
          response.write "      <th>Email</th>" & vbcrlf
          response.write "      <th>Fax</th>" & vbcrlf
          'response.write "      <th>Link Clicked</th>" & vbcrlf
          'response.write "      <th>Click Date</th>" & vbcrlf
       end if

   '---------------------------------------------------------------------------
    elseif lcl_customreport = "MAPPOINTS_EXPORT" then

      'All columns are dynamic in Map-Points.  Retrieve all of the columns for the Map-Point
       sSQL = "SELECT mptf.fieldname "
       sSQL = sSQL & " FROM egov_mappoints_types_fields mptf, egov_mappoints_types mpt "
       sSQL = sSQL & " WHERE mptf.mappoint_typeid = mpt.mappoint_typeid "
       sSQL = sSQL & " AND mptf.mappoint_typeid = " & lcl_mappoint_typeid
       sSQL = sSQL & " ORDER BY mptf.resultsOrder "

     		set oMPT_headers = Server.CreateObject("ADODB.Recordset")
     		oMPT_headers.Open sSQL, Application("DSN"), 3, 1

       if not oMPT_headers.eof then
          do while not oMPT_headers.eof

             response.write "      <th>" & oMPT_headers("fieldname") & "</th>" & vbcrlf

             oMPT_headers.movenext
          loop
       end if

       set oMPT_headers = nothing

   '---------------------------------------------------------------------------
    else
       response.write "      <th>&nbsp;</th>" & vbcrlf
    end if

    response.write "  </tr>" & vbcrlf

    while not oCustom.eof
       iRowCount = iRowCount + 1

       if lcl_customreport <> "PDFFIELDNAMES" then

         'Display only for non-export --------------------------------------------
          lcl_row_events = ""

          if lcl_export <> "Y" then
             lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

            'Setup the javascript events for the row.
             lcl_row_events = " id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"""
 
          else
             lcl_bgcolor = ""
          end if

         'Build the row(s) per report
          response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """" & lcl_row_events & ">" & vbcrlf
       end if

      '------------------------------------------------------------------------
      'Display column data per report
      '------------------------------------------------------------------------
       if lcl_customreport = "CODESECTIONS" then
      '------------------------------------------------------------------------

         'Track all of the codesection_ids AND the request_ids to be used in the summaries.
          if oCustom("submitted_action_code_id") <> "" then
             if lcl_codesection_ids = "" then
                lcl_codesection_ids = oCustom("submitted_action_code_id")
             else
                lcl_codesection_ids = lcl_codesection_ids & ", " & oCustom("submitted_action_code_id")
             end if

             if lcl_requestids = "" then
                lcl_requestids = oCustom("action_autoid")
             else
                lcl_requestids = lcl_requestids & ", " & oCustom("action_autoid")
             end if

          else
             lcl_codesection_ids = lcl_codesection_ids
             lcl_requestids      = lcl_requestids
          end if

         'Format the Submit/Complete Data
          lcl_submitdate   = "&nbsp;"
          lcl_completedate = "&nbsp;"

          if oCustom("submit_date") <> "" then
             lcl_submitdate = formatdatetime(oCustom("submit_date"),vbShortDate)
          end if

          if oCustom("complete_date") <> "" then
             lcl_completedate = formatdatetime(oCustom("complete_date"),vbShortDate)
          end if

          response.write "      <td>" & oCustom("TrackingNumber")         & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & lcl_submitdate   & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & lcl_completedate & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("status")                 & "</td>" & vbcrlf

          if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
              response.write "      <td>" & oCustom("sub_status_desc")        & "</td>" & vbcrlf
          end if

          response.write "      <td>" & oCustom("issuelocation_address")  & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("code_name")              & "</td>" & vbcrlf
	  if session("orgid") = "209" then
          	response.write "      <td>" & oCustom("assignedname")              & "</td>" & vbcrlf
	  end if

      '------------------------------------------------------------------------
       elseif lcl_customreport = "SUBSCRIBERLIST" then
      '------------------------------------------------------------------------
          response.write "      <td>" & oCustom("userfname")           & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userlname")           & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("useremail")           & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userstreetnumber")    & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userstreetprefix")    & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("useraddress")         & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("usercity")            & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userstate")           & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userzip")             & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userbusinessname")    & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userbusinessnumber")  & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userbusinessaddress") & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("userhomephone")) & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("usercell"))      & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("userworkphone")) & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("userfax"))       & "</td>" & vbcrlf

      '------------------------------------------------------------------------
       elseif lcl_customreport = "PDFFIELDNAMES" then
      '------------------------------------------------------------------------
          session("PDFRowCount") = 0
          session("PDF_bgcolor") = lcl_bgcolor
          session("lcl_export")  = lcl_export

          buildPDFRow "<strong>General Purpose Fields:</strong>","",""
          buildPDFRow "","TodaysDate",""

        	'Contact fields
        		sSQL = "SELECT userfname,userlname,userbusinessname,useremail,userhomephone,userfax,useraddress,usercity,userstate,userzip "
          sSQL = sSQL & " FROM egov_users" 

        		set oUserFields = Server.CreateObject("ADODB.Recordset")
        		oUserFields.Open sSQL, Application("DSN"), 3, 1

          buildPDFRow "<strong>Contact Fields:</strong>","",""

        	 for each field in oUserFields.fields
       						if field.name <> "userpassword" then
                buildPDFRow "",field.name,""
       						end if
        		next

          oUserFields.close
        		set oUserFields = nothing

        	'Additional Information field
          buildPDFRow "<strong>Additional Comment Field:</strong>","",""
          buildPDFRow "","Admin_Additional_Comments",""

        	'Tracking Number field
          buildPDFRow "<strong>Tracking Number Field:</strong>","",""
          buildPDFRow "","Tracking Number",""

        	'Code Sections field
          buildPDFRow "<strong>Code Section Field:</strong>","",""
          buildPDFRow "","Code_Sections",""

        	'Issue Location fields
          buildPDFRow "<strong>Issue Location Fields:</strong>","",""

         	sSQL = "SELECT streetnumber,streetaddress,city,state,zip,comments "
          sSQL = sSQL & " FROM egov_action_response_issue_location"

         	set oLocFields = Server.CreateObject("ADODB.Recordset")
         	oLocFields.Open sSQL, Application("DSN"), 3, 1

         	for each field in oLocFields.fields
             buildPDFRow "",field.name,""
         	next

          oLocFields.close
         	set oLocFields = nothing

        	'Form fields
          buildPDFRow "<strong>Form Fields:</strong>","",""

        	'Additional field values
         	sSQL = "SELECT pdfformname, sequence, answerlist, fieldtype "
          sSQL = sSQL & " FROM egov_action_form_questions "
          sSQL = sSQL & " WHERE pdfformname IS NOT NULL "
          sSQL = sSQL & " AND pdfformname <> '' "
          sSQL = sSQL & " AND formid = " & request("iFormID")
          sSQL = sSQL & " ORDER BY sequence "

        	 set oDynamicFields = Server.CreateObject("ADODB.Recordset")
         	oDynamicFields.Open sSQL, Application("DSN"), 3, 1

        	 if not oDynamicFields.eof then
         		  while not oDynamicFields.eof
           		  	if oDynamicFields("pdfformname") <> "" then
                   if oDynamicFields("fieldtype") = 6 then
                      buildPDFRow "",oDynamicFields("answerlist"),"Yes"
                   else
                      buildPDFRow "",oDynamicFields("pdfformname"),""
                   end if
             			end if
     			        oDynamicFields.movenext
           		wend
          end if

          oDynamicFields.close
          set oDynamicFields = nothing

      '------------------------------------------------------------------------
       elseif lcl_customreport = "PUBLICREQUESTSBYORG" then
      '------------------------------------------------------------------------

          response.write "      <td>" & oCustom("orgname")                       & "</td>" & vbcrlf
          response.write "      <td>" & FormatNumber(oCustom("TotalRequests"),0) & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & oCustom("orgid")        & "</td>" & vbcrlf

      '------------------------------------------------------------------------
       elseif lcl_customreport = "RSSLOG" then
      '------------------------------------------------------------------------

          response.write "      <td>" & oCustom("title")           & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("description")     & "</td>" & vbcrlf
          response.write "      <td>" & session("egovclientwebsiteurl") & oCustom("rsslink") & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("publicationdate") & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("createdbyname")   & "</td>" & vbcrlf

      '------------------------------------------------------------------------
       elseif lcl_customreport = "HELP_DOCUMENTATION" then
      '------------------------------------------------------------------------

         'All help documentation is stored in E-Gov Support!
          lcl_replace_string = "/public_documents300/custom/pub/egovsupport/published_documents/"

          response.write "      <td>" & oCustom("helpdoc_name")                               & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & oCustom("documentid")                & "</td>" & vbcrlf
          response.write "      <td>" & replace(oCustom("documenturl"),lcl_replace_string,"") & "</td>" & vbcrlf

      '------------------------------------------------------------------------
       elseif lcl_customreport = "CLASSEVENTS_TEAMROSTER" then
      '------------------------------------------------------------------------

          lcl_birthdate = ""
          lcl_age       = ""

          if oCustom("birthdate") <> "" then
             lcl_birthdate = replace(oCustom("birthdate"),"1/1/1900","")
          end if

          if lcl_birthdate <> "" then
             lcl_age = getCitizenAge(lcl_birthdate)
          end if

          response.write "      <td>" & oCustom("userfname")                        & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("userlname")                        & "</td>" & vbcrlf
          response.write "      <td>" & lcl_age                                     & "</td>" & vbcrlf
          response.write "      <td>" & lcl_birthdate                               & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("rostergrade")                      & "</td>" & vbcrlf
          response.write "      <td align=""left"">" & oCustom("rostershirtsize")   & "</td>" & vbcrlf
          response.write "      <td align=""left"">" & oCustom("rosterpantssize")   & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("userhomephone")) & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("parentfirstname")                  & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("parentlastname")                   & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("parentphone"))   & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("useraddress_complete")             & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("rostercoachtype")                  & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("rostervolunteercoachname")         & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("rostervolunteercoachdayphone"))  & "</td>" & vbcrlf
          response.write "      <td>" & FormatPhoneNumber(oCustom("rostervolunteercoachcellphone")) & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("rostervolunteercoachemail")        & "</td>" & vbcrlf

      '------------------------------------------------------------------------
       elseif lcl_customreport = "MEMBERSHIPCARD_SCANS_BY_MEMBERID" then
      '------------------------------------------------------------------------

          response.write "      <td align=""center"">" & oCustom("memberid")   & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("membername")                  & "</td>" & vbcrlf
          response.write "      <td>" & oCustom("scan_datetime")               & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & oCustom("poolpassid") & "</td>" & vbcrlf


      '------------------------------------------------------------------------
       elseif lcl_customreport = "MAPPOINTS_EXPORT" then
      '------------------------------------------------------------------------

          sSQL = "SELECT mpv.mp_valueid, "
          sSQL = sSQL & " mpv.orgid, "
          sSQL = sSQL & " mpv.mappoint_typeid, "
          sSQL = sSQL & " mpv.mappointid, "
          sSQL = sSQL & " mpv.mp_fieldid, "
          sSQL = sSQL & " mpv.fieldtype, "
          sSQL = sSQL & " mpv.fieldvalue, "
          sSQL = sSQL & " mptf.fieldname "
          sSQL = sSQL & " FROM egov_mappoints_values AS mpv "
          sSQL = sSQL &      " INNER JOIN egov_mappoints_types_fields AS mptf ON mpv.mappoint_typeid = mptf.mappoint_typeid "
          sSQL = sSQL &      " AND mpv.mp_fieldid = mptf.mp_fieldid "
          sSQL = sSQL & " WHERE mpv.mappointid = " & oCustom("mappointid")
          sSQL = sSQL & " AND mpv.mappoint_typeid = " & lcl_mappoint_typeid
          sSQL = sSQL & " ORDER BY mptf.resultsOrder "

        	 set oGetMPTValues = Server.CreateObject("ADODB.Recordset")
         	oGetMPTValues.Open sSQL, Application("DSN"), 3, 1

          if not oGetMPTValues.eof then
             do while not oGetMPTValues.eof
                if oGetMPTValues("fieldtype") <> "" then
                   lcl_td_nowrap = " nowrap=""nowrap"""
                else
                   lcl_td_nowrap = ""
                end if

                response.write "      <td" & lcl_td_nowrap & ">" & oGetMPTValues("fieldvalue")    & "</td>" & vbcrlf

                oGetMPTValues.movenext
             loop
          end if

          'oGetMPTValues.close
          set oGetMPTValues = nothing

      '------------------------------------------------------------------------
       elseif lcl_customreport = "CLICKCOUNTS" then
      '------------------------------------------------------------------------

          if session("CR_CLICKCOUNTS") = "POSTINGS" then

             if oCustom("username") = ", " then
                lcl_username = ""
             else
                lcl_username = oCustom("username")
             end if

             if trim(oCustom("useremail")) <> "" then
                lcl_useremail = trim(oCustom("useremail"))
             else
                lcl_useremail = ""
             end if

             if oCustom("userworkphone") <> "" then
                lcl_userworkphone = FormatPhoneNumber(oCustom("userworkphone"))
             else
                lcl_userworkphone = ""
             end if

             if oCustom("userfax") <> "" then
                lcl_userfax = FormatPhoneNumber(oCustom("userfax"))
             else
                lcl_userfax = ""
             end if

             response.write "      <td>" & oCustom("jobbid_id")            & "</td>" & vbcrlf
             response.write "      <td>" & oCustom("title")                & "</td>" & vbcrlf
             response.write "      <td nowrap=""nowrap"">" & lcl_username  & "</td>" & vbcrlf
             response.write "      <td>" & oCustom("userbusinessname")     & "</td>" & vbcrlf
             response.write "      <td>" & lcl_userworkphone               & "</td>" & vbcrlf
             response.write "      <td>" & lcl_useremail                   & "</td>" & vbcrlf
             response.write "      <td>" & lcl_userfax                     & "</td>" & vbcrlf
             'response.write "      <td>" & oCustom("clicked_date")         & "</td>" & vbcrlf
             'response.write "      <td>" & oCustom("clicked_linktext")     & "</td>" & vbcrlf
          end if

      '------------------------------------------------------------------------
       end if
      '------------------------------------------------------------------------

       if lcl_customreport <> "PDFFIELDNAMES" then
          response.write "  </tr>" & vbcrlf
       end if

    			oCustom.movenext
    wend

    response.write "</table>" & vbcrlf

   'Display only for non-export -----------------------------------------------
    if lcl_export <> "Y" then
       response.write "</div>" & vbcrlf

       showTotalRecords lcl_show_totalrecords, lcl_export, lcl_totalrecords_label, iRowCount

       response.write "</p>" & vbcrlf

    else

       showTotalRecords lcl_show_totalrecords, lcl_export, lcl_totalrecords_label, iRowCount

    end if
   '---------------------------------------------------------------------------

		  oCustom.close
  		set oCustom = nothing 

   'If any requests have any code sections then display the summary section.
    if lcl_codesection_ids <> "" then
       displaySummary lcl_codesection_ids, lcl_requestids, lcl_export
    end if

 else
  		response.write "<font style=""color:#ff0000;font-weight:bold"">No records exist.</font>" & vbcrlf
 end if

'Display only for non-export --------------------------------------------------
 if lcl_export <> "Y" then

    displaybuttons

    response.write " 	</div>" & vbcrlf
    response.write "</div>" & vbcrlf

    response.write "<!--#Include file=""../admin_footer.asp""-->  " & vbcrlf

    response.write "</body>" & vbcrlf
    response.write "</html>" & vbcrlf
 end if

'------------------------------------------------------------------------------
 sub displaybuttons()

   response.write "<p>" & vbcrlf
   response.write "<input type=""button"" name=""sClose"" id=""sClose"" value=""Close Window"" class=""button"" onclick=""window.close()"" />" & vbcrlf
   response.write "<input type=""button"" name=""sRefresh"" id=""sRefresh"" value=""Refresh Results"" class=""button"" onclick=""window.location.reload();"" />" & vbcrlf
   response.write "<input type=""button"" name=""sExport"" id=""sExport"" value=""Export Results"" class=""button"" onclick=""launchExport()"" />" & vbcrlf
   response.write "</p>" & vbcrlf

 end sub

'------------------------------------------------------------------------------
sub displaySummary(p_codesection_ids,p_requestids,p_isExport)

  if p_codesection_ids <> "" then
     sSQL = "SELECT distinct action_code_id, code_name, "
     sSQL = sSQL & " (select count(submitted_request_id) "
     sSQL = sSQL &  " from egov_submitted_request_code_sections sc "
     sSQL = sSQL &  " where submitted_action_code_id = action_code_id "
     sSQL = sSQL &  " and submitted_request_id IN (" & p_requestids & ")) AS total_requests "
     sSQL = sSQL & " FROM egov_actionline_code_sections "
     sSQL = sSQL & " WHERE action_code_id IN (" & p_codesection_ids & ")"
     sSQL = sSQL & " AND orgid = " & session("orgid")

    	set oCodes = Server.CreateObject("ADODB.Recordset")
    	oCodes.Open sSQL, Application("DSN"), 0, 1

     if not oCodes.eof then

        response.write "<p class=""page_break_before""><font size=""+1""><strong>Summary</strong></font></p>" & vbcrlf

        if p_isExport <> "Y" then
           response.write "<div class=""shadow"">" & vbcrlf
           response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"">" & vbcrlf
        else
           response.write "<table border=""0"">" & vbcrlf
        end if

        response.write "  <tr align=""left"">" & vbcrlf
        response.write "      <th>Code Section</th>" & vbcrlf
        response.write "      <th width=""150"" align=""center"">Total Requests Found On</th>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        lcl_bgcolor = "#ffffff"

        while not oCodes.eof
           if p_isExport <> "Y" then
              lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
           else
              lcl_bgcolor = ""
           end if

           response.write "  <tr align=""left"" valign=""bottom"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           response.write "      <td>" & oCodes("code_name") & "</td>" & vbcrlf
           response.write "      <td width=""150"" align=""center"">" & oCodes("total_requests") & "</td>" & vbcrlf
           response.write "  </tr>" & vbcrlf

           oCodes.movenext
        wend

        response.write "</table>" & vbcrlf

        if p_isExport <> "Y" then
           response.write "</div>" & vbcrlf
        end if

     end if

     oCodes.close
     set oCodes = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub displayFillerTDs(iOpenTR,iTotalColumns,iCloseTR)

  if iTotalColumns <> "" then
     if isnumeric(iTotalColumns) then
        if iTotalColumns > 0 then

          'Check to see if a TR is to be opened
           if UCASE(iOpenTR) = "Y" then
              response.write "<tr>" & vbcrlf
           end if

          'Determine the number of "filler" TDs needed
           for i = 1 to iTotalColumns
              response.write "    <td></td>" & vbcrlf
           next

          'Check to see if a TR is to be closed
           if UCASE(iCloseTR) = "Y" then
              response.write "</tr>" & vbcrlf
           end if
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
sub buildPDFRow(iColumnValue1,iColumnValue2,iColumnValue3)
  lcl_row_events   = ""
  lcl_columnvalue1 = "&nbsp;"
  lcl_columnvalue2 = "&nbsp;"
  lcl_columnvalue3 = "[data entered by user]"

  if iColumnValue1 <> "" then
     lcl_columnvalue1 = iColumnValue1
  end if

  if iColumnValue2 <> "" then
     lcl_columnvalue2 = iColumnValue2
  end if

  if iColumnValue3 <> "" then
     lcl_columnvalue3 = iColumnValue3
  end if

  if session("lcl_export") <> "Y" then
     lcl_bgcolor = changeBGColor(session("PDF_bgcolor"),"","")
     session("PDF_bgcolor") = lcl_bgcolor
     iRowCount   = session("PDFRowCount") + 1

    'Setup the javascript events for the row.
     lcl_row_events = " id=""" & iRowCount & """" ' onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"""
 
  else
     lcl_bgcolor = ""
  end if

 'Build the row(s) per report
  response.write "  <tr bgcolor=""" & lcl_bgcolor & """" & lcl_row_events & ">" & vbcrlf
  response.write "      <td>" & lcl_columnvalue1 & "</td>" & vbcrlf
  response.write "      <td>" & lcl_columnvalue2 & "</td>" & vbcrlf
  response.write "      <td>" & lcl_columnvalue3 & "</td>" & vbcrlf
  response.write "  </tr>" & vbclrf

end sub

'------------------------------------------------------------------------------
function getFeatureID(iFeature)
  lcl_return = 0

  if iFeature <> "" then
     sSQL = "SELECT featureid "
     sSQL = sSQL & " FROM egov_organization_features "
     sSQL = sSQL & " WHERE UPPER(feature) = '" & UCASE(iFeature) & "'"

     set oFID = Server.CreateObject("ADODB.Recordset")
     oFID.Open sSQL, Application("DSN"), 0, 1

     if not oFID.eof then
        lcl_return = oFID("featureid")
     end if

     oFID.close
     set oFID = nothing

  end if

  getFeatureID = lcl_return

end function

'------------------------------------------------------------------------------
sub showTotalRecords(iShowTotalRecords, iIsExport, iTotalRecordsLabel, iTotal)

  if iShowTotalRecords = "Y" then
     if iIsExport <> "Y" then
        response.write "<div>" & vbcrlf
     end if

     response.write "<strong>Total" & iTotalRecordsLabel & ": </strong>[" & iTotal & "]" & vbcrlf

     if iIsExport <> "Y" then
        response.write "</div>" & vbcrlf
     end if

  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
  set dtb = Server.CreateObject("ADODB.Recordset")
  dtb.Open sSQLi, Application("DSN"), 0, 1

end sub
%>
