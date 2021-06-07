<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: view_donotknock.asp
' AUTHOR: David Boyer
' CREATED: 12/03/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  User can view "Do Not Knock" List(s).
'
' MODIFICATION HISTORY
' 1.0  12/03/2010 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
 session("redirectlang") = "Return to ""Do Not Knock"" List"
 sTitle = "View ""Do Not Knock"" List"

'If they do not have a userid set, take them to the login page automatically
 if request.cookies("userid") = "" or request.cookies("userid") = "-1" then
	   session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
   	response.redirect "user_login.asp?p=dnk"
 end if

 Dim sUserType, iUserid, sResidentDesc
 sUserType = "P"
 iUserid   = request.cookies("userid")

'Determine which list(s) the user has access to view
 lcl_userid            = request.cookies("userid")
 lcl_canViewPeddlers   = checkAccessToList(lcl_userid, iorgid, "peddlers")
 lcl_canViewSolicitors = checkAccessToList(lcl_userid, iorgid, "solicitors")

 if NOT lcl_canViewPeddlers AND NOT lcl_canViewSolicitors then
    response.redirect "default_page.asp"
 end if

'Determine which page title to display
 if lcl_canViewPeddlers then
    if lcl_canViewSolicitors then
       lcl_title = "Peddlers/Solicitors ""Do Not Knock"" Lists"
    else
       lcl_title = "Peddlers ""Do Not Knock"" List"
    end if
 else
    if lcl_canViewSolicitors then
       lcl_title = "Solicitors ""Do Not Knock"" List"
    end if
 end if
%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName & " - " & sTitle %></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	<script language="javascript" src="scripts/modules.js"></script>
	<script language="javascript" src="scripts/easyform.js"></script>  

<script language="javascript">
<!--
function openWin2(url, name) {
    popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function exportList() {
  lcl_url  = "view_donotknock_export.asp";
  lcl_url += "?vp=<%=lcl_canViewPeddlers%>";
  lcl_url += "&vs=<%=lcl_canViewSolicitors%>";

  openWin2(lcl_url,"export");
}
//-->
</script>
</head>

<!--#include file="include_top.asp"-->
<%
'BEGIN: Body Content ---------------------------------------------------------
 response.write "<font class=""pagetitle"">Welcome to " & sOrgName & " ""Do Not Knock"" List</font><br />" & vbcrlf

 RegisteredUserDisplay( "" )

 response.write "<div id=""content"">" & vbcrlf
 response.write "  <div id=""centercontent"">" & vbcrlf

'Retrieve search options
 lcl_sc_streetname = ""
 lcl_sc_listtype   = ""

 if request("sc_streetname") <> "" then
    lcl_sc_streetname = request("sc_streetname")
 end if

 if request("sc_listtype") <> "" then
    lcl_sc_listtype = request("sc_listtype")
 end if

'BEGIN: Build search criteria -------------------------------------------------
 response.write "<p>" & vbcrlf
 response.write "<fieldset>" & vbcrlf
 response.write "  <legend>Search Options&nbsp;</legend>" & vbcrlf
 response.write "  <p>" & vbcrlf
 response.write "    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 response.write "      <form name=""searchForm"" id=""searchForm"" action=""view_donotknock.asp"" method=""post"">" & vbcrlf
 response.write "      <tr>" & vbcrlf
 response.write "          <td>Address:&nbsp;</td>" & vbcrlf
 response.write "          <td>" & vbcrlf
                               displayStreetName_dropdown iOrgID, lcl_sc_streetname
 response.write "          </td>" & vbcrlf
 response.write "      </tr>" & vbcrlf

 if lcl_canViewPeddlers AND lcl_canViewSolicitors then

    lcl_listtype_selected_all        = ""
    lcl_listtype_selected_peddlers   = ""
    lcl_listtype_selected_solicitors = ""

    if lcl_sc_listtype = "PEDDLERS" then
       lcl_listtype_selected_peddlers = " selected=""selected"""
    elseif lcl_sc_listtype = "SOLICITORS" then
       lcl_listtype_selected_solicitors = " selected=""selected"""
    else
       lcl_listtype_selected_all = " selected=""selected"""
    end if

    response.write "      <tr>" & vbcrlf
    response.write "          <td>List Type:&nbsp;</td>" & vbcrlf
    response.write "          <td>" & vbcrlf
    response.write "              <select name=""sc_listtype"" id=""sc_listtype"">" & vbcrlf
    response.write "                <option value=""ALL"""        & lcl_listtype_selected_all        & ">All</option>" & vbcrlf
    response.write "                <option value=""PEDDLERS"""   & lcl_listtype_selected_peddlers   & ">Peddlers List</option>" & vbcrlf
    response.write "                <option value=""SOLICITORS""" & lcl_listtype_selected_solicitors & ">Solicitors List</option>" & vbcrlf
    response.write "              </select>" & vbcrlf
    response.write "          </td>" & vbcrlf
    response.write "      </tr>" & vbcrlf
 end if

 response.write "      <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "      <tr>" & vbcrlf
 response.write "          <td colspan=""2"">" & vbcrlf
 response.write "              <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
 response.write "          </td>" & vbcrlf
 response.write "      </tr>" & vbcrlf
 response.write "      </form>" & vbcrlf
 response.write "    </table>" & vbcrlf
 response.write "  </p>" & vbcrlf
 response.write "</fieldset>" & vbcrlf
 response.write "</p>" & vbcrlf
'END: Build search criteria ---------------------------------------------------

'BEGIN: Build the query based on what the user has access to view -------------
 lcl_linecount = 0
 lcl_bgcolor   = "#ffffff"

 sSQL = "SELECT DISTINCT "
 sSQL = sSQL & " userstreetnumber, "
 sSQL = sSQL & " useraddress, "
 sSQL = sSQL & " userunit, "
 sSQL = sSQL & " isOnDoNotKnockList_peddlers, "
 sSQL = sSQL & " isOnDoNotKnockList_solicitors, "
 sSQL = sSQL & " dbo.breakOut_StreetAddress('STREETNUMBER',useraddress) as db_streetnumber, "
 sSQL = sSQL & " dbo.breakOut_StreetAddress('STREETNAME',useraddress) as db_streetname "
 sSQL = sSQL & " FROM egov_users "
 sSQL = sSQL & " WHERE orgid = " & iOrgID
 sSQL = sSQL & " AND useraddress <> '' "
 sSQL = sSQL & " AND useraddress IS NOT NULL "

 if lcl_canViewPeddlers then
     if lcl_canViewSolicitors then
        if lcl_sc_listtype = "PEDDLERS" then
           sSQL = sSQL & " AND isOnDoNotKnockList_peddlers = 1 "
        elseif lcl_sc_listtype = "SOLICITORS" then
           sSQL = sSQL & " AND isOnDoNotKnockList_solicitors = 1 "
        else
           sSQL = sSQL & " AND (isOnDoNotKnockList_peddlers = 1 OR isOnDoNotKnockList_solicitors = 1) "
        end if
     else
        sSQL = sSQL & " AND isOnDoNotKnockList_peddlers = 1 "
     end if
 else
     if lcl_canViewSolicitors then
        sSQL = sSQL & " AND isOnDoNotKnockList_solicitors = 1 "
     end if
 end if

 if lcl_sc_streetname <> "" then
    sSQL = sSQL & " AND UPPER(useraddress) LIKE ('%" & UCASE(lcl_sc_streetname) & "%') "
 end if

 sSQL = sSQL & " ORDER BY dbo.breakOut_StreetAddress('STREETNAME',useraddress), "
 sSQL = sSQL &          " dbo.breakOut_StreetAddress('STREETNUMBER',useraddress), "
 sSQL = sSQL &          " userunit "

'Build the results table
 if sSQL <> "" then

   'Set up the session variable for the export
    session("DONOTKNOCK_QUERY") = sSQL
	response.write "<!--" & sSQL & "-->"

   	set oKnockList = Server.CreateObject("ADODB.Recordset")
   	oKnockList.Open sSQL, Application("DSN"), 3, 1

    if not oKnockList.eof then
       do while not oKnockList.eof
          lcl_linecount = lcl_linecount + 1
          lcl_bgcolor   = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

          if lcl_linecount = 1 then
             response.write "<p>" & vbcrlf
             response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width:800px;"">" & vbcrlf
             response.write "  <tr>" & vbcrlf
             response.write "      <td style=""font-size:12pt; font-weight:bold"">" & lcl_title & "</td>" & vbcrlf
             response.write "      <td align=""right"">" & vbcrlf
             response.write "          <input type=""button"" name=""exportListButton"" id=""exportListButton"" class=""button"" value=""Export List"" onclick=""exportList()"" />" & vbcrlf
             response.write "      </td>" & vbcrlf
             response.write "  </tr>" & vbcrlf
             response.write "</table>" & vbcrlf
             response.write "</p>" & vbcrlf

             response.write "<div class=""transactionreportshadow"" style=""width:800px"">" & vbcrlf
             response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""transactionreport"" style=""width:800px"">" & vbcrlf
             response.write "  <tr align=""center"">" & vbcrlf
             response.write "      <td class=""transaction_header"" width=""100%"" align=""left"">Address (Resident Unit)</td>" & vbcrlf

             if lcl_canViewPeddlers AND lcl_canViewSolicitors then
                response.write "      <td class=""transaction_header"" nowrap=""nowrap"">On Peddler<br />List</td>" & vbcrlf
                response.write "      <td class=""transaction_header"" nowrap=""nowrap"">On Solicitor<br />List</td>" & vbcrlf
             end if

             response.write "  </tr>" & vbcrlf
          end if

         'Format fields to display
          lcl_display_useraddress         = "&nbsp;"
          lcl_display_userunit            = ""
          lcl_display_isOnList_peddlers   = "&nbsp;"
          lcl_display_isOnList_solicitors = "&nbsp;"

          if oKnockList("useraddress") <> "" then
             lcl_display_useraddress = oKnockList("useraddress")
          end if

          if oKnockList("userunit") <> "" then
             lcl_display_userunit    = oKnockList("userunit")
             lcl_display_useraddress = lcl_display_useraddress & " (" & lcl_display_userunit & ")"
          end if

          if lcl_canViewPeddlers then
             if oKnockList("isOnDoNotKnockList_peddlers") then
                lcl_display_isOnList_peddlers = "Y"
             end if
          end if

          if lcl_canViewSolicitors then
             if oKnockList("isOnDoNotKnockList_solicitors") then
                lcl_display_isOnList_solicitors = "Y"
             end if
          end if

          response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
          response.write "      <td width=""100%"" align=""left"">" & lcl_display_useraddress & "</td>" & vbcrlf

          if lcl_canViewPeddlers AND lcl_canViewSolicitors then
             response.write "      <td nowrap=""nowrap"">" & lcl_display_isOnList_peddlers   & "</td>" & vbcrlf
             response.write "      <td nowrap=""nowrap"">" & lcl_display_isOnList_solicitors & "</td>" & vbcrlf
          end if

          response.write "  </tr>" & vbcrlf

          oKnockList.movenext
       loop
    end if

    oKnockList.close
    set oKnockList = nothing

    if lcl_linecount > 0 then
       response.write "</table>" & vbcrlf
       response.write "</div>" & vbcrlf
    end if
 end if

'Build a "no record available" table, if needed
 if sSQL = "" OR lcl_linecount < 1 then
    response.write "<p>" & vbcrlf
    response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""transactionreport"" style=""width:800px"">" & vbcrlf
    response.write "  <tr><td class=""transaction_header"" width=""100%"" align=""left"">&nbsp;</td></tr>" & vbcrlf
    response.write "  <tr><td align=""left"">No Addresses Available</td></tr>" & vbcrlf
    response.write "</table>" & vbcrlf
    response.write "</p>" & vbcrlf
 end if

'END: Build the query based on what the user has access to view ---------------

 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "<p><br />&nbsp;<br />&nbsp;</p>" & vbcrlf
'END: Body Content -----------------------------------------------------------
%>
<!--#include file="include_bottom.asp"-->
<%
'------------------------------------------------------------------------------
'function checkAccessToList(iUserID, iOrgID, iListType)

' lcl_return = False

'Determine the list type
' if iListType <> "" then

'    sSQL = "SELECT isDoNotKnockVendor_" & iListType & " as 'ListAccess' "
'    sSQL = sSQL & " FROM egov_users "
'    sSQL = sSQL & " WHERE orgid = " & iOrgID
'    sSQL = sSQL & " AND userid = " & iUserID

'   	set oCheckListAccess = Server.CreateObject("ADODB.Recordset")
' 	  oCheckListAccess.Open sSQL, Application("DSN"), 3, 1

'    if not oCheckListAccess.eof then
'       lcl_return = oCheckListAccess("ListAccess")
'    end if

'    oCheckListAccess.close
'    set oCheckListAccess = nothing

' end if

' checkAccessToList = lcl_return

'end function

'------------------------------------------------------------------------------
sub displayStreetName_dropdown(p_orgid, iCurrentValue)
	Dim sSQL, sCompareName

 sStreetNumber  = ""
 sStreetName    = ""
 sCompareStreet = ""

 if iCurrentValue <> "" then
    BreakOutAddress iCurrentValue, sStreetNumber, sStreetName
    sCompareStreet = sStreetName
 end if

 sSQL = "SELECT DISTINCT "
 'sSQL = sSQL & " useraddress "
 sSQL = sSQL & " dbo.breakOut_StreetAddress('STREETNAME',useraddress) as useraddress "
' sSQL = sSQL & ", userunit "
 sSQL = sSQL & " FROM egov_users "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
 sSQL = sSQL & " AND useraddress <> '' "
 sSQL = sSQL & " AND useraddress IS NOT NULL "
 sSQL = sSQL & " AND (isOnDoNotKnockList_peddlers = 1 "
 sSQL = sSQL & "  OR  isOnDoNotKnockList_solicitors = 1) "
 sSQL = sSQL & " ORDER BY dbo.breakOut_StreetAddress('STREETNAME',useraddress) "
	
	set oSCStreetName = Server.CreateObject("ADODB.Recordset")
	oSCStreetName.Open sSQL, Application("DSN"), 0, 1

	if not oSCStreetName.eof then
  		response.write "<select name=""sc_streetname"" id=""sc_streetname"">" & vbcrlf
    response.write "  <option value=""""></option>" & vbcrlf

    do while not oSCStreetName.eof
       lcl_useraddress         = ""
       'lcl_userunit            = ""
       sStreetNumber           = ""
       sStreetName             = ""
       lcl_streetname_selected = ""

       'if trim(oSCStreetName("useraddress")) <> "" then
          'lcl_useraddress = oSCStreetName("useraddress")
       'end if
       sStreetName = oSCStreetName("useraddress")

       'BreakOutAddress lcl_useraddress, sStreetNumber, sStreetName

       if sStreetName = sCompareStreet then
          lcl_streetname_selected = " selected=""selected"""
       end if

       response.write "  <option value=""" & sStreetName & """" & lcl_streetname_selected & ">" & sStreetName & "</option>" & vbcrlf

       oSCStreetName.movenext
    loop

    response.write "</select>" & vbcrlf
 end if

	oSCStreetName.close
	set oSCStreetName = nothing 

end sub
%>
