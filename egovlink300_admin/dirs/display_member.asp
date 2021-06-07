<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: display_member.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays a list of admin users
'
' MODIFICATION HISTORY
' 1.0 ??/??/??	  ????        - INITIAL VERSION
' 1.1 08/04/2009 David Boyer - Added "Delegate"
' 1.2	08/28/2009	Steve Loar - Added Rental supervisor pick for Menlo Park Rentals project
' 1.3 05/28/2010 David Boyer - Removed "Delete" button and "checkboxes" per Peter's request (task: 833)
' 1.4	10/14/2011	Steve Loar - Chenaged to not show those flagged as deleted.
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim lcl_orghasfeature_rental_supervisors, lcl_checked_rental_supervisors
	dim pagesize, totalpages, RA, totalrecords, groupname, thisname, conn, rs, strSQL, CName
	dim numstartid, numendid, i, EventOrNot, Str_Bgcolor, username, password, str_image, editurl, FullName

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

 if not UserHasPermission(session("userid"),"edit users") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the search parameters
 lcl_sc_firstname = request("sc_firstname")
 lcl_sc_lastname  = request("sc_lastname")
 lcl_sc_orderby   = request("sc_orderby")
 lcl_group_id     = request("groupid")

'Set up the ORDER BY
 if lcl_sc_orderby <> "" then
    if lcl_sc_orderby = "lastname" then
       lcl_orderby = "UPPER(lastname), UPPER(firstname)"
    elseif lcl_sc_orderby = "firstname" then
       lcl_orderby = "UPPER(firstname), UPPER(lastname)"
    elseif lcl_sc_orderby = "email" then
       lcl_orderby = "UPPER(email), UPPER(lastname), UPPER(firstname)"
    end if
 else
    lcl_orderby = "UPPER(lastname), UPPER(firstname)"
 end if

'Set up the link parameters for the return url for the search criteria options
 lcl_return_url_parameters = ""

 if lcl_sc_firstname <> "" then
    lcl_return_url_parameters = "sc_firstname=" & lcl_sc_firstname
 end if

 if lcl_sc_lastname <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_lastname=" & lcl_sc_lastname
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_lastname=" & lcl_sc_lastname
    end if
 end if

 if lcl_sc_orderby <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_orderby=" & lcl_sc_orderby
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_lorderby=" & lcl_sc_orderby
    end if
 end if

 if lcl_group_id <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "groupid=" & lcl_group_id
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&groupid=" & lcl_group_id
    end if
 end if

 if lcl_return_url_parameters <> "" then
    lcl_return_url_parameters = "&" & REPLACE(lcl_return_url_parameters,"%","<<PER>>")
 end if

'Convert the % in the search criteria
 lcl_sc_firstname = REPLACE(request("sc_firstname"),"<<PER>>","%")
 lcl_sc_lastname  = REPLACE(request("sc_lastname"),"<<PER>>","%")

'Check for org features
 lcl_orghasfeature_class_supervisors = orghasfeature("class supervisors")
 lcl_orghasfeature_rental_supervisors = orghasfeature("create edit rentals")
 lcl_orghasfeature_admin_locations   = orghasfeature("admin locations")
 lcl_orghasfeature_action_line       = orghasfeature("action line")

'Check for user permissions
 lcl_userhaspermission_edit_users  = userhaspermission(session("userid"),"edit users")
 lcl_userhaspermission_action_line = userhaspermission(session("userid"),"action line")
%>
<html>
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Administration Console {Edit User}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<style type="text/css">
.fieldset
{
   margin: 5px;
   padding: 5px;
   border-radius: 6px;
}

.fieldset legend
{
   padding: 4px 8px;
   border: 1pt solid #808080;
   border-radius: 6px;
   font-size: 1.25em;
   color: #800000;
}

.buttonRow
{
   padding-bottom: 5px;
}

.buttonRow a:hover
{
   text-decoration: none;
}

#buttonPrevious,
#buttonNext,
#buttonEditMembership
{
   cursor: pointer;
}

#searchOptionsTable
{
   margin: 10px;
   width: 80% !important;
}

#searchOptionsTable td
{
   padding: 2px;
}
</style>

	<script type="text/javascript" src="../scripts/modules.js"></script>
	<script type="text/javascript" src="../scripts/selectAll.js"></script>
	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
 <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

	<script type="text/javascript">
	<!--
		function ChangeSupervisor( iUserId ) 
		{
			//alert( 'User = ' + iUserId );
			doAjax('setclasssupervisor.asp', 'userid=' + iUserId, '', 'get', '0');
		}

		function ChangeRentalSupervisor( iUserId ) 
		{
			//alert( 'User = ' + iUserId );
			doAjax('setrentalsupervisor.asp', 'userid=' + iUserId, '', 'get', '0');
		}

		function SupervisorSet( sReturn ) 
		{
			// Nothing happens here
			alert( sReturn );
		}
		function openWin1(url, name) 
		{
			popupWin = window.open(url, name, "resizable,width=550,height=450");
		}

		function openWin2(url, name) 
		{
			//popupWin = window.open(url, name, "resizable,width=380,height=300");
			popupWin = window.open(url, name, "resizable,width=500,height=350");
		}

	//-->
	</script>
</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  'groupmode  = 1, display individual group
		'groupmode  = 2, display all member
		'totalpages = 1
		HowManyMembersDisplayed = 0

 	thisname    = request.servervariables("script_name")
		pagesize    = GetUserPageSize(Session("UserId"))  'Steve Loar 2/5/2007
		currentpage = 1

		if not isempty(request("currentpage")) then
					currentpage = request("currentpage")
 	end if

  response.write "<div id=""content"">" & vbcrlf
  response.write " 	<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "	 <tr>" & vbcrlf
  response.write "  				<td valign=""top"">" & vbcrlf
                            DisplayRecords session("orgid"), _
                                           currentpage
  response.write "	     </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<!--#include file="footer.asp"-->
<%
'------------------------------------------------------------------------------
sub DisplayRecords(iOrgID, _
                   iCurrentPage)

  dim iShowCount, iRowCount, sSQL, sOrgID, sCurrentPage

	 iShowCount   = 0
  sOrgID       = 0
  sCurrentPage = 1
	 thisname     = request.servervariables("script_name")

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iCurrentPage <> "" then
     if not containsApostrophe(iCurrentPage) then
        sCurrentPage = clng(iCurrentPage)
     end if
  end if

	 set conn = Server.CreateObject("ADODB.Connection")
	 conn.Open Application("DSN")

	 set rs = Server.CreateObject("ADODB.Recordset")
	 set rs.ActiveConnection = conn

	 rs.CursorLocation = 3 
	 rs.CursorType     = 3 

 	if trim(lcl_group_id) <> "" then
     GroupMode = 1

   		sSQL = "SELECT u.userid, "
     sSQL = sSQL & " u.firstname, "
     sSQL = sSQL & " u.lastname, "
     sSQL = sSQL & " ISNULL(email,'') AS email, "
     sSQL = sSQL & " username, "
     sSQL = sSQL & " password, "
   		sSQL = sSQL & " ISNULL(locationid,0) AS locationid, "
     sSQL = sSQL & " isclasssupervisor, "
     sSQL = sSQL & " groupname, "
   		sSQL = sSQL & " (SELECT u2.firstname + ' ' + u2.lastname FROM users u2 WHERE u2.userid = u.delegateid) AS delegate_name, "
   		sSQL = sSQL & " (SELECT u2.email FROM users u2 WHERE u2.userid = u.delegateid) AS delegate_email, "
     sSQL = sSQL & " isrentalsupervisor "
   		sSQL = sSQL & " FROM users u, "
     sSQL = sSQL &      " usersgroups ug, "
     sSQL = sSQL &      " groups g "
   		sSQL = sSQL & " WHERE u.userid = ug.userid "
     sSQL = sSQL & " AND g.groupid = ug.groupid "
     sSQL = sSQL & " AND isdeleted = 0 "
   		sSQL = sSQL & " AND u.orgid = " & sOrgID

		   if lcl_sc_firstname <> "" then
     			sSQL = sSQL & " AND UPPER(firstname) like ('%" & dbsafe(UCASE(lcl_sc_firstname)) & "%') "
   		end if

   		if lcl_sc_lastname <> "" then
			     sSQL = sSQL & " AND UPPER(lastname) like ('%" & dbsafe(UCASE(lcl_sc_lastname)) & "%') "
   		end if

   		if not IsRootAdmin( Session("UserId") ) then
     			sSQL = sSQL & " AND isrootadmin <> 1 "
	   	end if

   		sSQL = sSQL & " AND ug.groupid = " & lcl_group_id
   		sSQL = sSQL & " ORDER BY " & lcl_orderby
 	else 
   		GroupMode = 2

   		sSQL = "SELECT userid, "
     sSQL = sSQL & " u.firstname, "
     sSQL = sSQL & " u.lastname, "
     sSQL = sSQL & " ISNULL(email,'') AS email, "
     sSQL = sSQL & " username, "
     sSQL = sSQL & " password, "
   		sSQL = sSQL & " ISNULL(locationid,0) AS locationid, "
     sSQL = sSQL & " isclasssupervisor, "
     sSQL = sSQL & " '' as groupname, "
   		sSQL = sSQL & " (SELECT u2.firstname + ' ' + u2.lastname FROM users u2 WHERE u2.userid = u.delegateid) AS delegate_name, "
   		sSQL = sSQL & " (SELECT u2.email FROM users u2 WHERE u2.userid = u.delegateid) AS delegate_email, "
     sSQL = sSQL & " isrentalsupervisor "
   		sSQL = sSQL & " FROM users u "
   		sSQL = sSQL & " WHERE (username IS NOT NULL) "
     sSQL = sSQL & " AND isdeleted = 0 "
   		sSQL = sSQL & " AND username <> '' "

   		if lcl_sc_firstname <> "" then
     			sSQL = sSQL & " AND UPPER(firstname) like ('%" & dbsafe(UCASE(lcl_sc_firstname)) & "%') "
   		end if

    	if lcl_sc_lastname <> "" then
     			sSQL = sSQL & " AND UPPER(lastname) like ('%" & dbsafe(UCASE(lcl_sc_lastname)) & "%') "
     end if

   		if not IsRootAdmin( Session("UserId") ) then
     			sSQL = sSQL & " AND isrootadmin <> 1 "
    	end if

   		sSQL = sSQL & " AND u.orgid = " & sOrgID
   		sSQL = sSQL & " ORDER BY " & lcl_orderby

 	end if

 	rs.Open sSQL

 'BEGIN: The following code dealing with the recordcount=0 --------------------
 	if rs.recordcount = 0 then
    	rs.close

    	if groupmode = 1 then
      		sSQL = "SELECT groupname FROM groups WHERE groupid = " & CLng(lcl_group_id)

      		rs.Open sSQL

      		CName     = "Department:&nbsp;" & rs("groupname") 
      		groupname = rs("groupname")

      		rs.close
	    else
      		CName = "Department:&nbsp;" & langDiaplyMember 
    	end if

    	statistics sOrgID, _
                CName, _
                groupmode

    	displayButtons groupmode, _
                    sCurrentPage

    	response.write "<p>" & vbcrlf
     response.write "<div class=""shadow"">" & vbcrlf
   	 response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tablelist"">" & vbcrlf
    	response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;" & langUser & "</th>" & vbcrlf
     response.write "      <th align=""center"">&nbsp;" & langTypeEmail & "</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
   	 response.write "  <tr>" & vbcrlf
     response.write "      <td align=""left"" colspan=""2"">" & vbcrlf
     response.write "          &nbsp;&nbsp;<font size=""1"" color=""#FF0000""><b>" & langNoRecords & "</b></font>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf

    	exit sub
 	end if
 'END: The following code dealing with the recordcount=0 ----------------------

 	rs.movefirst
 	totalrecords = rs.RecordCount
 	TotalPages   = (totalrecords \ pagesize) + 1  '\means integer/integer

	 if totalrecords Mod pagesize=0 and TotalPages > 0 then
     TotalPages=TotalPages-1
  end if

	 if totalrecords <= pagesize then
     TotalPages = 1
  end if

	 if TotalPages < 1 then
     TotalPages = 1
  end if

	 if isNumeric(sCurrentPage) then
   		if sCurrentPage < 1 then
        sCurrentPage = 1
     end if

	 	  if clng(sCurrentPage) > clng(TotalPages) then
        sCurrentPage = TotalPages
     end if
	 else
   		sCurrentPage = 1
	 end if

	 numstartid	= (sCurrentPage-1) * PageSize

	 if numstartid + PageSize < totalrecords then
   		numendid = numstartid + pagesize - 1
	 else
	 	  numendid = totalrecords - 1
	 end if

	 RA = rs.GetRows

	 if groupmode = 1 then
   		CName     = "Department:&nbsp;" & RA(8,i)
		   groupname = RA(8,i)
	 else
   		CName     = "Department:&nbsp;" & langDiaplyMember
	 end If

	 statistics sOrgID, _
             CName, _
             groupmode

	 displayButtons groupmode, _
                 sCurrentPage

 'BEGIN: Display all users -----------------------------------------------------
 	response.write "<form name=""allMembers"" id=""allMembers"" method=""post"" action=""delete_multipleuser.asp?previousURL=" & thisname & "&Extra=" & request.querystring & """>" & vbcrlf
 	response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" id=""adminuserlist"">" & vbcrlf
 	response.write "  <tr>" & vbcrlf
 	response.write "      <th align=""left"" width=""40%"">" & langUser & "</th>" & vbcrlf

 	if lcl_orghasfeature_class_supervisors then
   		response.write "      <th align=""left"">Class Supervisor</th>" & vbcrlf
 	end If

 	if lcl_orghasfeature_rental_supervisors then
   		response.write "<th align=""left"">Rental<br />Supervisor</th>"
 	end if

 	if lcl_orghasfeature_admin_locations then
   		response.write "      <th align=""left"">Location</th>" & vbcrlf
 	end if

 	response.write "      <th>&nbsp;</th>" & vbcrlf
 	response.write "      <th align=""left"">&nbsp;" & langTypeEmail & "</th>" & vbcrlf
 	response.write "      <th align=""left"">Edit Groups</th>" & vbcrlf

  if lcl_orghasfeature_action_line AND lcl_userhaspermission_action_line then
 	   response.write "      <th align=""left"">Delegate</th>" & vbcrlf
  end if

 	response.write "  </tr>"

 	j = numstartid	 'For alt row colors
 	i = numstartid  '0, 20, 40 usually
 	iShowCount  = i
 	iRowCount   = 0
  lcl_bgcolor = "#eeeeee"

 	do while iShowCount <= numendid
   		lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
  	 	username    = RA(4,i)
   		password    = RA(5,i)

   		if (isnull(username)) or (username="") then
 		     	str_image="<img src=""../images/.gif"" />" & vbcrlf
   		else
 		     	str_image="<img src=""../images/newuser.gif"" />" & vbcrlf
   		end if

   		editurl  = "update_user.asp?userid=" & RA(0,i) & "&currentpage=" & sCurrentPage & lcl_return_url_parameters
   		FullName = trim(RA(2,i))&",&nbsp;" & trim(RA(1,i))

   		if lcl_userhaspermission_edit_users AND NOT IsRootAdmin( RA(0,i) ) then
	 	     j = j + 1

     			HowManyMembersDisplayed = HowManyMembersDisplayed + 1
     			iRowCount               = iRowCount + 1

     			response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
     			response.write "      <td onClick=""location.href='" & editurl & "';"">" & FullName & "</td>" & vbcrlf

    				if lcl_orghasfeature_class_supervisors then
      					if RA(7,i) then
        						lcl_checked_class_supervisors = " checked=""checked"""
      					else
        						lcl_checked_class_supervisors = ""
           end if

      					response.write "<td align=""center"">"
      					response.write "<input type=""checkbox"" name=""issupervisor"" value=""" & RA(0,i) & """ onclick=""ChangeSupervisor(" & RA(0,i) & ");""" & lcl_checked_class_supervisors & " />" 
      					response.write "</td>"
    				end if
				
    				if lcl_orghasfeature_rental_supervisors then
      					if RA(11,i) then
        						lcl_checked_rental_supervisors = " checked=""checked"""
      					else
        						lcl_checked_rental_supervisors = ""
      					end if

      					response.write "<td align=""center"">"
      					response.write "<input type=""checkbox"" name=""isrentalsupervisor"" value=""" & RA(0,i) & """ onclick=""ChangeRentalSupervisor(" & RA(0,i) & ");""" & lcl_checked_rental_supervisors & " />" 
		      			response.write "</td>"
   				 end if

     			if lcl_orghasfeature_admin_locations then
        			response.write "      <td onClick=""location.href='" & editurl & "';"" nowrap=""nowrap"">" & GetUserLocation( RA(6,i) ) & "</td>" & vbcrlf
     			end if

     			if RA(3,i) <> "" then
   			     response.write "      <td align=""center""><a href=""mailto:" & RA(3,i) & """><img src=""../images/newmail_small.gif"" border=""0"" /></a></td>" & vbcrlf
     			else
        			response.write "      <td onclick=""location.href='" & editurl & "';"">&nbsp;</td>" & vbcrlf
    	 		end if

        response.write "      <td onClick=""location.href='" & editurl & "';"">&nbsp;" & RA(3,i) & "</td>" & vbcrlf
     			response.write "      <td><em><a href=""javascript:openWin2('ManageMemberGroup.asp?userid="&RA(0,i)&"','_blank')"">"&langEdit&"</a></em></td>" & vbcrlf

       'Set up the delegate
        if lcl_orghasfeature_action_line AND lcl_userhaspermission_action_line then
           lcl_delegate       = ""
           lcl_delegate_name  = ""
           lcl_delegate_email = ""

           if trim(RA(9,i)) <> "" then
              lcl_delegate = trim(RA(9,i))
           end if

          'Check for delegate's email
           if trim(RA(10,i)) <> "" then
              lcl_delegate_email = trim(RA(10,i))
           end if

           if lcl_delegate_name <> "" then
              lcl_delegate = lcl_delegate_name
           end if

           if lcl_delegate_email <> "" then
              if lcl_delegate <> "" then
                 lcl_delegate = lcl_delegate & "<br />[<span style=""color:#800000"">" & lcl_delegate_email & "]</span>"
              else
                 lcl_delegate = lcl_delegate_email
              end if
           end if

           response.write "      <td>" & lcl_delegate & "</td>" & vbcrlf
        end if

     			response.write "  </tr>" & vbcrlf

     			iShowCount = iShowCount + 1

   		end if

   		i = i + 1
  loop

 	if HowManyMembersDisplayed = 0 then
   		response.write "  <tr style=""height:25px;"">" & vbcrlf
     response.write "      <td colspan=""5"">No users that you can view!</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
 	end if
	
 	response.write "</table>" & vbcrlf
 	response.write "</form>" & vbcrlf

 	displayButtons groupmode, _
                 sCurrentPage

 	rs.close
 	set rs = nothing

end sub

'------------------------------------------------------------------------------
sub statistics(iOrgID, _
               CName, _
               iGroupMode)

  dim lcl_orgid

  lcl_orgid = 0

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        lcl_orgid = clng(iOrgID)
     end if
  end if

	'Select order by value
 	if lcl_sc_orderby = "lastname" then
   		lcl_lastname_selected  = " selected"
   		lcl_firstname_selected = ""
   		lcl_email_selected     = ""
 	elseif lcl_sc_orderby = "firstname" then
   		lcl_lastname_selected  = ""
   		lcl_firstname_selected = " selected"
   		lcl_email_selected     = ""
 	elseif lcl_sc_orderby = "email" then
   		lcl_lastname_selected  = ""
   		lcl_firstname_selected = ""
   		lcl_email_selected     = " selected"
 	else
   		lcl_lastname_selected  = " selected"
   		lcl_firstname_selected = ""
   		lcl_email_selected     = ""
  end if

 	response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
 	response.write "  <tr>" & vbcrlf
 	response.write "      <td><font size=""+1""><strong>" & CName & "</strong></font>" & vbcrlf

 	if iGroupMode = 1 then
   		'response.write "<br /><img src=""../images/arrow_2back.gif"" align=""absmiddle"" />&nbsp;<a href=""display_committee.asp"">Back to Department List</a>" & vbcrlf
 		  response.write "<br /><input type=""button"" name=""returnButton"" id=""returnButton"" class=""button"" value=""Back to Department List"" onclick=""location.href='display_committee.asp';"" />" & vbcrlf
 	end if

 	response.write "      </td>" & vbcrlf
 	response.write "  </tr>" & vbcrlf
 	response.write "</table>" & vbcrlf
 	response.write "<fieldset class=""fieldset"">" & vbcrlf
 	response.write "  <legend>Search Options</legend>" & vbcrlf
 	response.write "  <form name=""search_sort_form"" value=""display_member.asp"">" & vbcrlf
 	response.write "<table id=""searchOptionsTable"" border=""0"" cellspacing=""0"">" & vbcrlf
 	response.write "  <tr valign=""top"">" & vbcrlf
 	response.write "      <td width=""50%"">" & vbcrlf
 	response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
 	response.write "            <tr>" & vbcrlf
 	response.write "                <td>First Name:</td>" & vbcrlf
 	response.write "                <td width=""75%""><input type=""text"" name=""sc_firstname"" value=""" & lcl_sc_firstname & """ size=""15"" maxlength=""25"" /></td>" & vbcrlf
 	response.write "            </tr>" & vbcrlf
 	response.write "            <tr>" & vbcrlf
 	response.write "                <td>Last Name:</td>" & vbcrlf
 	response.write "                <td width=""75%""><input type=""text"" name=""sc_lastname"" value=""" & lcl_sc_lastname & """ size=""15"" maxlength=""25"" /></td>" & vbcrlf
 	response.write "            </tr>" & vbcrlf
 	response.write "          </table>" & vbcrlf
 	response.write "      </td>" & vbcrlf
 	response.write "      <td>" & vbcrlf
 	response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 	response.write "            <tr>" & vbcrlf
 	response.write "                <td>Department:</td>" & vbcrlf
 	response.write "                <td>" & vbcrlf
 	response.write "                    <select name=""groupid"">" & vbcrlf
 	response.write "                      <option value=""""></option>" & vbcrlf

 	sSQLg = "SELECT groupid, groupname "
 	sSQLg = sSQLg & " FROM groups "
 	sSQLg = sSQLg & " WHERE orgid = " & session("orgid")
 	sSQLg = sSQLg & " ORDER BY UPPER(groupname) "

 	set rsg = Server.CreateObject("ADODB.Recordset")
 	rsg.Open sSQLg, Application("DSN"), 1, 3

 	if not rsg.eof then
	   	do while not rsg.eof
      		if lcl_group_id = "" then
       				lcl_selected = ""
     			else
       				if rsg("groupid") = CLng(lcl_group_id) then
         					lcl_selected = " selected"
       				else
         					lcl_selected = ""
       				end if
     			end if

      		response.write "<option value=""" & rsg("groupid") & """" & lcl_selected & ">" & rsg("groupname") & "</option>" & vbcrlf

     			rsg.movenext
    	loop 
  end if

 	response.write "                    </select>" & vbcrlf
 	response.write "                </td>" & vbcrlf
 	response.write "            </tr>" & vbcrlf
	 response.write "            <tr>" & vbcrlf
	 response.write "                <td>Sort:</td>" & vbcrlf
	 response.write "                <td>" & vbcrlf
	 response.write "                    <select name=""sc_orderby"">" & vbcrlf
	 response.write "                      <option value=""lastname"""  & lcl_lastname_selected  & ">Last Name</option>" & vbcrlf
	 response.write "                      <option value=""firstname""" & lcl_firstname_selected & ">First Name</option>" & vbcrlf
	 response.write "                      <option value=""email"""     & lcl_email_selected     & ">Email</option>" & vbcrlf
	 response.write "                    </select>" & vbcrlf
	 response.write "                </td>" & vbcrlf
	 response.write "            </tr>" & vbcrlf
	 response.write "          </table>" & vbcrlf
	 response.write "      </td>" & vbcrlf
	 response.write "  </tr>" & vbcrlf
	 response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
	 response.write "  <tr><td colspan=""2""><input type=""submit"" name=""searchButton"" id=""searchButton"" value=""SEARCH"" class=""button"" /></td></tr>" & vbcrlf
	 response.write "</table>" & vbcrlf
	 response.write "  </form>" & vbcrlf
	 response.write "</fieldset>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayButtons(iGroupMode, _
                   iCurrentPage)

  dim sAddItOnURL, sGroupMode, lcl_currentpage

  sAddItOnURL     = ""
  sGroupMode      = 2
  lcl_currentpage = 1
  sButtonDisabledPrevious = " disabled=""disabled"""
  sButtonDisabledNext     = " disabled=""disabled"""

  if iGroupMode <> "" then
     if not containsApostrophe(iGroupMode) then
        sGroupMode = clng(iGroupMode)
     end if
  end if

  if iCurrentPage <> "" then
     if not containsApostrophe(iCurrentPage) then
        lcl_currentpage = clng(iCurrentPage)
     end if
  end if

 	if sGroupMode = 1 then
   		sAddItOnURL = "groupid=" & lcl_group_id & "&"
 	end if

 'Determine which buttons are enabled based on the curent page
 	if clng(lcl_currentpage) > 1 then
     lcl_buttonPrev_style = ""
  end if

  if clng(lcl_currentpage) < clng(totalpages) then
     lcl_buttonNext_style = ""
  end if

	 response.write "<div class=""buttonRow"">" & vbcrlf

 'BEGIN: Button - Previous ----------------------------------------------------
 	if clng(lcl_currentpage) > 1 then
     sButtonDisabledPrevious = ""

   		response.write "<a href=""" & thisname & "?" & sAddItOnURL & "currentpage=" & (lcl_currentpage-1) & "&" & REPLACE(lcl_return_url_parameters,"&"&REPLACE(sAddItOnURL,"&",""),"") & """>"
  end if

  response.write "<input type=""button"" name=""buttonPrevious"" id=""buttonPrevious"" value=""Prev " & pagesize & """" & sButtonDisabledPrevious & " />" & vbcrlf

 	if clng(lcl_currentpage) > 1 then
     response.write "</a>" & vbcrlf
  end if
 'END: Button - Previous ------------------------------------------------------

 'BEGIN: Button - Next --------------------------------------------------------
 	if clng(lcl_currentpage) < clng(totalpages) then
     sButtonDisabledNext = ""

	   	response.write "<a href=""" & thisname & "?" & sAddItOnURL & "currentpage=" & (lcl_currentpage+1) & "&" & REPLACE(lcl_return_url_parameters,"&"&REPLACE(sAddItOnURL,"&",""),"") & """>"
 	end if

 response.write "  <input type=""button"" name=""buttonNext"" id=""buttonNext"" value=""Next " & pagesize & """" & sButtonDisabledNext & " />" & vbcrlf

 	if clng(lcl_currentpage) < clng(totalpages) then 
	   	response.write "</a>"
 	end if
 'END: Button - Next ----------------------------------------------------------

 'BEGIN: Button - Edit Membership ---------------------------------------------
 	if lcl_group_id <> "" then
     response.write "&nbsp;&nbsp;" & vbcrlf
     response.write "<input type=""button"" name=""buttonEditMembership"" id=""buttonEditMembership"" value=""Edit Membership"" onclick=""openWin2('ManageCommitteeMember.asp?groupid=" & lcl_group_id & "','_blank');"" />" & vbcrlf
 	end if
 'END: Button - Edit Membership -----------------------------------------------

  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
Function IsRootAdmin( ByVal iUserId )
	Dim sSQL, oRs, blnReturnValue

	'SET DEFAULT
	blnReturnValue = False

	'If the root is viewing this, then let them see their own record, otherwise No.
	If Session("UserID") <> iUserId Then
		sSQL = "SELECT COUNT(userid) AS root_count "
		sSQL = sSQL & " FROM users  WHERE orgid = " & Session("OrgID")
		sSQL = sSQL & " AND userid = " & iUserId
		sSQL = sSQL & " AND isrootadmin = 1"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 0, 1

  		If clng(oRs("root_count")) > 0 Then
   			'the ORGANIZATION HAS the FEATURE
   			blnReturnValue = True
  		End If

  		oRs.Close 
  		Set oRs = Nothing

	End If 

	IsRootAdmin = blnReturnValue

End Function 


%>
