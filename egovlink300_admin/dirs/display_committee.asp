<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<% 
 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "departments") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

 dim pagesize, totalpages,totalrecords,groupmode
 dim thisname, Currentpage, rs, strSQL, conn, numstartid, numendid
 dim strDirectory, conn2, rs2, strSQL2, AllUsers, totalUsers, strUser
 dim totalContacts, strContact, AdditonURL, deleteurl
 dim RA, i, GroupNumber, editurl, temp, Str_Bgcolor, HowManyCommitteeDisplayed
 dim bCanEdit

 'groupmode=1, display individual group
 'groupmode=2, display all member
 HowManyCommitteeDisplayed = 0
 pagesize   = GetUserPageSize( session("userid") ) ' Steve Loar 2/21/2007
 totalpages = 1
 thisname   = request.servervariables("script_name")

 if request("currentpage") <> "" AND isNumeric(request("currentpage")) then
	   currentpage = clng(request("currentpage"))

 			if clng(currentpage) < 1 then
    	  currentpage = 1
    end if

 else
	   currentpage = 1
 end if

'Retrieve the search criteria
 lcl_sc_department = request("sc_department")

'Display the screen message
 lcl_message = ""

 if request("success") = "SN" then
    lcl_message = "Successfully Created..."
 elseif request("success") = "SD" then
    lcl_message = "Successfully Deleted..."
 end if

 if lcl_message <> "" then
    lcl_message = "<strong class=""screenMsg"">*** " & lcl_message & " ***</strong>"
 else
    lcl_message = "&nbsp;"
 end if
%>
<html>
<head>
  <title><%=langBSCommittees%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

<style type="text/css">
.fieldset
{
   border-radius: 6px;
}

.fieldset legend
{
   border: 1pt solid #808080;
     border-radius: 6px;
   padding: 4px 8px;
   font-size: 1.25em;
   color: #800000;
}

.screenMsg
{
   color: #ff0000;
}

#buttonPrevious,
#buttonNext,
#buttonNewDepartment,
#buttonDelete
{
   cursor: pointer;
}

#departmentList
{
   width: 800px;
}

#departmentList tbody td
{
   white-space: nowrap;
}
</style>
  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function() {
   $('#sc_department').focus();

   $('#buttonPrevious').click(function() {
      changePage('PREVIOUS');
   });

   $('#buttonNext').click(function() {
      changePage('NEXT');
   });
});

function changePage(iDirection)
{
   var lcl_url;
   var lcl_currentPage;

   lcl_currentPage = 0;

   if($('#currentPage').val() != '')
   {
      lcl_currentPage = eval($('#currentPage').val());
   }

   if(iDirection != '')
   {
      if(iDirection == 'PREVIOUS')
      {
         lcl_currentPage = (lcl_currentPage - 1);
      }
      else
      {
         lcl_currentPage = (lcl_currentPage + 1);
      }
   }

   lcl_url = '<%=thisname%>?<%=AdditonURL%>currentpage=' + lcl_currentPage;

   location.href = lcl_url;
}

function openWin2(url, name) {
		popupWin = window.open(url, name,"resizable,width=500,height=250");
}

function confirmDelete() {
  iRtn = confirm('<%=langWanttoDelete%>');
  if (iRtn) {
      document.getElementById("DeleteCommittee").submit();
  }
}
//-->
</script>
</head>

<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp" //-->
<%
  response.write "<body>" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & session("sOrgName") & "&nbsp;Departments</strong></font><br />" & vbcrlf
  response.write "          <div id=""dir_info""></div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">" & lcl_message & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Search Options</legend>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "  <form name=""search_form"" method=""post"" action=""display_committee.asp"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          Department:&nbsp;<input type=""text"" name=""sc_department"" id=""sc_department"" value=""" & lcl_sc_department & """ size=""30"" maxlength=""150"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
  response.write "  <tr><td><input type=""submit"" value=""SEARCH"" /></td></tr>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
  response.write "</p>" & vbcrlf

 'page size, RA, pagerecord, currentpage values must be declared to global variables.
  bCanEdit = True 
  thisname = request.servervariables("script_name")

 	set conn = Server.CreateObject("ADODB.Connection")
	 conn.Open Application("DSN")

 	set rs = Server.CreateObject("ADODB.Recordset")
 	set rs.ActiveConnection = conn
 	rs.CursorLocation = 3
	 rs.CursorType     = 3 

 	strSQL = "SELECT groupid, orgid, groupname, groupdescription "
  strSQL = strSQL & " FROM groups g "
  strSQL = strSQL & " WHERE g.orgid = " & session("OrgID")
  strSQL = strSQL & " AND isInactive <> 1 "

  if lcl_sc_department <> "" then
     strSQL = strSQL & " AND UPPER(g.groupname) LIKE ('%" & UCASE(lcl_sc_department) & "%') "
  end if

  strSQL = strSQL & " ORDER BY groupname"

 	rs.Open strSQL

 'Get the totals
	 totalrecords = rs.RecordCount
 	totalpages   = (totalrecords \ pagesize) + 1  '\means integer/integer

 	if totalpages < 1 then
  	  totalpages = 1
  end if

 	if clng(currentpage) >= clng(totalpages) then
     currentpage = totalpages
 	else
	    currentpage = 1
  end if

 	displayButtons currentpage

  response.write "<input type=""hidden"" name=""currentPage"" id=""currentPage"" value=""" & currentpage & """ />" & vbcrlf

 	response.write "<table id=""departmentList"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tablelist"">" & vbcrlf
 	response.write "  <form name=""DeleteCommittee"" id=""DeleteCommittee"" method=""post"" action=""delete_committee.asp?currentpage=" & currentpage & """>" & vbcrlf
  response.write "  <thead>" & vbcrlf
 	response.write "  <tr style=""height:25px;"">" & vbcrlf
 	response.write "      <th align=""left"">" & vbcrlf
  response.write "          <input class=""listcheck"" type=""checkbox"" name=""chkSelectAll"" onClick=""selectAll('DeleteCommittee', this.checked, 'delete')"" />" & vbcrlf
  response.write "      </th>" & vbcrlf
  'response.write "      <th width=""1"">&nbsp;</th>" & vbcrlf
  response.write "      <th align=""left"">" & langDirectory   & "</th>" & vbcrlf
  response.write "      <th align=""left"">" & langDescription & "</th>" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "      <th>"                & langEntries     & "</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <thead>" & vbcrlf
  response.write "  <tbody>" & vbcrlf


	'------- the following code dealing with the recordcount=0-------
 	if rs.recordcount = 0 then
	   	rs.close

     response.write "  <tr>" & vbcrlf
   		response.write "      <td colspan=""5"">There are no departments to view!</td>" & vbcrlf
  			response.write "  </tr>" & vbcrlf

  else

    	rs.movefirst

  	 	numstartid	= (currentpage-1) * pagesize
  		 numendid	  = IIf(numstartid + pagesize < totalrecords, numstartid+pagesize- 1, totalrecords - 1)

  		'---------- the following will display the whole table ----------
  	 	RA = rs.GetRows()

  	  if lcl_sc_department = "" then
     		 statistics session("orgid"), _
                   lcl_sc_department

  	     lcl_total_user_cnt = 0
  	  end if

  	 	for i = numstartid to numendid
     	 		GroupNumber        = entry(RA(0,i), _
                                    session("orgid"), _
                                    lcl_sc_department)

        	lcl_total_user_cnt = lcl_total_user_cnt + GroupNumber
  	    		editurl            = "display_member.asp?groupid="&RA(0,i)
     	 		temp               = "<br />CanView"&RA(2,i)

  	    		HowManyCommitteeDisplayed = HowManyCommitteeDisplayed+1
     	   Str_Bgcolor = changeBGColor(Str_Bgcolor,"#eeeeee","#ffffff")

  	    		response.write "  <tr bgcolor=" & Str_Bgcolor & ">" & vbcrlf
     				response.write "      <td>" & vbcrlf
        	response.write "          <input class=""listcheck"" type=""checkbox"" name=""delete"" value="""&RA(0,i)&""" />" & vbcrlf
  	      response.write "      </td>" & vbcrlf
     	 		'response.write "      <td style=""padding:0px;"">" & vbcrlf
        	'response.write "          <img src=""../images/newgroup.gif"" border=""0"">" & vbcrlf
  	    		'response.write "      </td>" & vbcrlf
     	 		response.write "      <td>" & vbcrlf
    		  	response.write "          <a href="""&editurl&""" class=""hotspot"" onmouseover=""tooltip.show('Click to view users');"" onmouseout=""tooltip.hide();"">" & RA(2,i) & "</a>" & vbcrlf
         response.write "      </td>" & vbcrlf
    		  	response.write "      <td>" & LEFT(RA(3,i),100) & "</td>" & vbcrlf
     	 		response.write "      <td>" & vbcrlf
  	    		'response.write "          &nbsp;&nbsp;" & vbcrlf
     	 		'response.write "          <a href=""Update_committee.asp?groupid=" & RA(0,i) & """><img src=""../images/edit.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Click to edit');"" onmouseout=""tooltip.hide();""></a>" & vbcrlf
    		  	'response.write "          &nbsp;" & vbcrlf
  	    		'response.write "          <a href=""" & getcommitteeemails(RA(0,i)) & """ class=""hotspot"" onmouseover=""tooltip.show('Click to send e-mail');"" onmouseout=""tooltip.hide();""><img src=""../images/newmail_small.gif"" border=""0""></a>" & vbcrlf
     	 		response.write "          <a href=""Update_committee.asp?groupid=" & RA(0,i) & """><input type=""button"" name=""buttonEdit" & i & """ id=""buttonEdit" & i & """ value=""Edit"" style=""cursor: pointer;"" /></a>" & vbcrlf
         response.write "          <a href=""" & getcommitteeemails(RA(0,i)) & """><input type=""button"" name=""buttonSendEmail" & i & """ id=""buttonSendEmail" & i & """ value=""Send Email"" style=""cursor: pointer;"" /></a>" & vbcrlf
     	 		response.write "      </td>" & vbcrlf
  	    		response.write "      <td align=""center"">" & GroupNumber & "</td>" & vbcrlf
     	 		response.write "  </tr>" & vbcrlf
  	  next

   	rs.close
   	set rs = nothing
   	conn.close
   	set conn = nothing

 	end if

  response.write "  </tbody>" & vbcrlf
 	response.write "</form>" & vbcrlf
 	response.write "</table>" & vbcrlf

'---------- end of displaying the whole table ----------
'Since users can search on departments we now need to recalculate
'the total user count since the call in the loop will set this value to the count of users
'in the first record.
 if lcl_sc_department <> "" then
    strDirectory = getDirectoryLabel(totalrecords)
    strUser      = getUserLabel(lcl_total_user_cnt)

  		response.write "<script>document.all.dir_info.innerHTML = """ & totalrecords & " " & strDirectory & ", " & lcl_total_user_cnt & " " & strUser & """</script>" & vbcrlf
 end if
%>

  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'------------------------------------------------------------------------------
function entry(groupid, _
               iOrgID, _
               iSCDepartment)

 dim sOrgID, sSCDepartment

 sOrgID        = 0
 sSCDepartment = ""

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iSCDepartment <> "" then
     sSCDepartment = ucase(iSCDepartment)
     sSCDepartment = dbsafe(sSCDepartment)
  end if

	' -100 means total records
	' if a groupid>0 it will show individual group
	 set conn2 = Server.CreateObject("ADODB.Connection")
	 conn2.Open Application("DSN")
	 set rs2 = Server.CreateObject("ADODB.Recordset")
	 set rs2.ActiveConnection = conn2
	 rs2.CursorLocation = 3 
	 rs2.CursorType     = 3 

 'Build WHERE CLAUSE for search criteria.
  if sSCDepartment <> "" then
     sSCDepartment = "'%" & sSCDepartment & "%'"

     lcl_where_clause = " AND u.userid IN (select distinct ug.userid "
     lcl_where_clause = lcl_where_clause & " from usersgroups ug, groups g "
     lcl_where_clause = lcl_where_clause & " where ug.groupid = g.groupid "
     lcl_where_clause = lcl_where_clause & " and g.orgid = " & sOrgID
     lcl_where_clause = lcl_where_clause & " and UPPER(g.groupname) LIKE (" & sSCDepartment & "))"
  else
     lcl_where_clause = ""
  end if

 	if groupid = -100 then
   		strSQL2 = "SELECT count(userid) as totaluser "
     strSQL2 = strSQL2 & " FROM users u "
     strSQL2 = strSQL2 & " WHERE username is not null "
     strSQL2 = strSQL2 & " AND username <> '' "
     strSQL2 = strSQL2 & " AND u.orgid=" & sOrgID
     strSQL2 = strSQL2 & " AND (isrootadmin is null or isrootadmin = 0)"
     strSQL2 = strSQL2 & lcl_where_clause
 	elseif groupid = -200 then
   		strSQL2 = "SELECT count(userid) as totaluser "
     strSQL2 = strSQL2 & " FROM users u "
     strSQL2 = strSQL2 & " WHERE ((username is null) or username='') "
     strSQL2 = strSQL2 & " AND u.orgid = " & sOrgID
     strSQL2 = strSQL2 & " AND (isrootadmin is null or isrootadmin = 0)"
     strSQL2 = strSQL2 & lcl_where_clause
 	else
 		  strSQL2 = "SELECT count(ug.userid) as totaluser "
     strSQL2 = strSQL2 & " FROM usersgroups ug "
     strSQL2 = strSQL2 &      " INNER JOIN users u ON u.userid = ug.userid "
     strSQL2 = strSQL2 & " WHERE ug.groupid = " & groupid
     strSQL2 = strSQL2 & " AND u.orgid = " & sOrgID
     strSQL2 = strSQL2 & " AND (isrootadmin is null or isrootadmin = 0)"
 	end if

 	rs2.Open strSQL2
 	AllUsers = rs2("totaluser")
 	rs2.close
 	conn2.close
 	set rs2   = nothing
 	set conn2 = nothing
 	entry     = AllUsers
end function 

'------------------------------------------------------------------------------
sub statistics(iOrgID, _
               iSCDepartment)

 	totalUsers = entry(-100, _
                     iOrgID, _
                     iSCDepartment)

  strDirectory = getDirectoryLabel(totalrecords)
  strUser      = getUserLabel(totalUsers)

  totalContacts = entry(-200, _
                        iOrgID, _
                        iSCDepartment)

	 if totalContacts > 1 then
   		strContact = langContacts
	 else
   		strContact = langContact
	 end if

	 if bCanEdit then
   		response.write "<script>document.all.dir_info.innerHTML = """ & totalrecords & " " & strDirectory & ", " & totalUsers & " " & strUser & """</script>" & vbcrlf
	 end if

end sub

'------------------------------------------------------------------------------
sub displayButtons(iCurrentPage)
 	response.write "<div style=""font-size:10pt; padding-bottom:10px;"">" & vbcrlf

  'if clng(currentpage) > clng(1) then 
		'   response.write "<a href="""&thisname&"?"&AdditonURL&"currentpage="&(currentpage-1)&""">"
  ' 		response.write "<img src=""../images/arrow_back.gif"" align=""absmiddle"" border=""0"">&nbsp;"&langPrev&" "&pagesize&"</a>" & vbcrlf
 	'else
	 '  	response.write "<img src=""../images/arrow_back.gif"" align=""absmiddle"" border=""0"">&nbsp;<font color=""#999999"">"&langPrev&" "&pagesize&"</font>" & vbcrlf
 	'end if

 	'if clng(currentpage) < clng(totalpages) then 
	 '  	response.write "&nbsp;&nbsp;<a href='"&thisname&"?"&AdditonURL&"currentpage="&(currentpage+1)&"'>"&langNext&" "&pagesize
  '   response.write	"&nbsp;<img src=""../images/arrow_forward.gif"" align=""absmiddle"" border=""0""></a>" & vbcrlf
 	'else
	 '  	response.write "&nbsp;&nbsp;" & vbcrlf
  '   response.write "<font color=""#999999"">"&langNext&" "&pagesize&"</font>"
  '   response.write	"&nbsp;<img src=""../images/arrow_forward.gif"" align=""absmiddle"" border=""0"">" & vbcrlf
	 'end if

  sButtonDisabledPrevious = " disabled=""disabled"""
  sButtonDisabledNext     = " disabled=""disabled"""

  if clng(iCurrentPage) > clng(1) then
     sButtonDisabledPrevious = ""
  end if

  if clng(iCurrentPage) < clng(totalpages) then
     sButtonDisabledNext = ""
  end if

  response.write "<input type=""button"" name=""buttonPrevious"" id=""buttonPrevious"" value=""Previous " & pagesize & """" & sButtonDisabledPrevious & " />" & vbcrlf
  response.write "<input type=""button"" name=""buttonNext"" id=""buttonNext"" value=""Next "             & pagesize & """" & sButtonDisabledNext     & " />" & vbcrlf

	'---------- the following is the additional convenient links showing on the top of table ----------
 	response.write	"&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
  response.write "<input type=""button"" name=""new_department"" id=""buttonNewDepartment"" value=""New Department"" onClick=""location.href='register_committee.asp'"" />" & vbcrlf
'  response.write "<img src=""../images/newgroup.gif"" width=""16"" height=""16"" align=""absmiddle"">&nbsp;"
'  response.write "<a href=""register_committee.asp"">New Department</a>" & vbcrlf

'  if clng(record) > 0 and bCanEdit then
  if clng(1) > 0 and bCanEdit then
     response.write "<input type=""button"" name=""delete"" id=""buttonDelete"" value=""Delete"" onClick=""confirmDelete();"" />" & vbcrlf
'     response.write "<img src=""../images/small_delete.gif"" align=""absmiddle"" border=""0"">&nbsp;" & vbcrlf
'     response.write "<a href=""javascript:document.all.DeleteCommittee.submit();"" onClick=""javascript: return confirm('" & langWanttoDelete & "');"">" & langDelete & "</a>" & vbcrlf
  end if

	 response.write "</div>" & vbcrlf
	'------- end of convenient links showing on the top of table-------------------------------------

end  sub

'------------------------------------------------------------------------------
function IIf(bCheck, sTrue, sFalse)
 	if bCheck then
     IIf = sTrue
  else
     IIf = sFalse

  end if
end function 

'------------------------------------------------------------------------------
' GET ALL EMAIL ADDRESSES FOR THIS COMMITTEE
function getcommitteeemails(iCommitteeID)
  Set conn = Server.CreateObject("ADODB.Connection")
	 conn.Open Application("DSN")
	 Set rs = Server.CreateObject("ADODB.Recordset")
	 Set rs.ActiveConnection = conn
	 rs.CursorLocation = 3 
	 rs.CursorType = 3 

 	strSQL = "SELECT u.userid, firstname,lastname, email, username, password, groupname "
  strSQL = strSQL & " FROM users u, usersgroups ug, groups g "
  strSQL = strSQL & " WHERE u.userid = ug.userid "
  strSQL = strSQL & " AND g.groupid = ug.groupid "
  strSQL = strSQL & " AND u.orgid = " & session("OrgID")
  strSQL = strSQL & " AND ug.groupid = " & iCommitteeID

 	rs.Open strSQL
	
	 emaillist = "mailto:"

	 do while not rs.eof
   		if trim(rs("email")) <> "" then
     			emaillist = emaillist + rs("email") + ";"
   		end if

    	rs.movenext
 	loop 

 	getcommitteeemails = emaillist

end function

'------------------------------------------------------------------------------
function getDirectoryLabel(p_totalrecords)
  lcl_return = "Department"

  if p_totalrecords > 1 then
   		lcl_return = lcl_return & "s"
 	end if

  getDirectoryLabel = lcl_return

end function

'------------------------------------------------------------------------------
function getUserLabel(p_total_user_cnt)
  lcl_return = ""

 	if p_total_user_cnt > 1 then
   		lcl_return = langUsers
 	else
	   	lcl_return = langUser
 	end if

  getUserLabel = lcl_return

end function
'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "') "
 	Set rsi = Server.CreateObject("ADODB.Recordset")
	 rsi.Open sSQLi, Application("DSN"), 0, 1

end sub
%>
