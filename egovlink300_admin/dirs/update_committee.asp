<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

if NOT UserHasPermission( Session("UserId"), "departments" ) then
	  response.redirect sLevel & "permissiondenied.asp"
end if

'Determine if the user is updating the record
 if UCASE(request("sAction")) = "UPDATE" then
   'Retrieve the value entered
   	newgroupname        = replace(trim(request("groupname")),"'","''")
   	newgroupdescription = replace(trim(request("groupdescription")),"'","''")
   	grouptype           = replace(trim(request("grouptype")),"'","''")

   'Update the record
   	sSQLu = "UPDATE groups SET "
    sSQLu = sSQLu & " groupname = '"        & dbsafe(newgroupname)        & "', "
    sSQLu = sSQLu & " groupdescription = '" & dbsafe(newgroupdescription) & "', "
    sSQLu = sSQLu & " grouptype = '"        & dbsafe(grouptype)           & "' "
    sSQLu = sSQLu & " WHERE groupid = " & clng(trim(request("groupid")))

   	set rsu = Server.CreateObject("ADODB.Recordset")
    rsu.Open sSQLu, Application("DSN") , 3, 1

    set rsu = nothing

    response.redirect "update_committee.asp?groupid=" & request("groupid") & "&success=SU"

 else

	 '-- check is the group id is entered or not ---------
  	if trim(request("groupid")) = "" then
      response.redirect "display_committee.asp"
 	 else
	 	   sSQL = "SELECT groupid, groupname, groupdescription, grouptype "
      sSQL = sSQL & " FROM groups "
      sSQL = sSQL & " WHERE groupid = " & clng(trim(request("groupid")))

     	set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sSQL, Application("DSN") , 3, 1

     	if not rs.eof then
         groupid          = rs("groupid")
         groupname        = rs("groupname")
         groupdescription = rs("groupdescription")
         grouptype        = rs("grouptype")
      else
         response.redirect "display_committee.asp"
    	 end if

      set rs = nothing

 end if

'Determine if there is a screen message to display
 lcl_onload = ""

 if request("success") = "SU" then
    lcl_onload = "displayScreenMsg('Successfully Updated...');"
 elseif request("success") = "NE" then
    lcl_onload = "displayScreenMsg('Group Name already exists...');"
 else
    lcl_onload = "&nbsp;"
 end if
%>

<html>
<head>
  <title><%=langBSCommittees%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
 	<script language="javascript" src="../scripts/textareamaxlength.js"></script>

<script language="javascript">
<!--
function CheckCommitteeField() {
		if (document.UpdateCommittee.GroupName.value == "") {
//   			alert("Group name is required");
  			 inlineMsg(document.getElementById("groupname").id,'<strong>Required Field Missing: </strong>Group Name',10,'groupname');
    		document.UpdateCommittee.GroupName.focus();
    		return false;
		}
		return true;
}

function doGroupsAccess() {
		x = (screen.width-450)/2;
		y = (screen.height-400)/2;
		win = window.open("ManageCommitteeAccess2.asp?groupid=<%=Request("groupid")%>", "disc_members", "width=450,height=350,status=0,menubar=0,scrollbars=1,toolbar=0,left="+x+",top="+y+",z-lock=yes");
		win.focus();
}

function openWin2(url, name) {
		popupWin = window.open(url, name,"resizable,width=380,height=300");
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
//-->
</script>
</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();<%=lcl_onload%>">

<!-- #include file="dir_constants.asp"-->

<div id="content">
  <div id="centercontent">

<table border="0" cellpadding="10" cellspacing="0" width="100%">
  <tr>
      <td>
          <font size="+1"><b><%=lanUpdateCommitteeTitle%></b></font><br />
          <input type="button" name="returnButton" id="returnButton" value="Back to Department List" class="button" onclick="location.href='display_committee.asp'" />
          <!-- <img src='../images/arrow_back.gif' align='absmiddle'> <a href="display_committee.asp">Back to Department List</a> -->
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
  <tr>
      <td colspan="2" valign="top">

          <form name="UpdateCommittee" id="UpdateCommittee" method="post" action="update_committee.asp">
          		<input type="hidden" name="GroupID" value="<%=groupid%>">

          <% displayButtons %>

          <table border="0" width="100%"   class='tablelist' cellpadding='5' cellspacing='0'>
          	 <tr>
                <th colspan="2" width="100" align="left"><%=langUpdate%>&nbsp;<%=langCommittee%></th>
            </tr>
          		<tr>
          		    <td width="10%" valign="top"><%=langGroup%>:</td>
          		    <td width="80%"><input type="text" name="GroupName" id="groupname" value="<%=groupname%>" size="50" maxlength="50" onchange="clearMsg('groupname');" /></td>
          		</tr>
          		<tr>
              		<td width="10%" valign="top"><%=langDescription%>:</td>
              		<td width="80%"><textarea rows="2" cols="50" name="GroupDescription" maxlength="150"><%=groupdescription%></textarea></td>  
          		</tr>
          		<tr>
              		<td width="10%" valign="top"><strong>Group Type:</strong></td>
              		<td width="80%">  
                  		<select name="grouptype">
                    <%
                      if grouptype = 2 then
                         lcl_selected = " SELECTED"
                      else
                         lcl_selected = ""
                      end if
                    %>
        		          	 <option value="2"<%=lcl_selected%>>Department</option>
                  		</select>
                </td>
            </tr>
          	</table>

          <% displayButtons %>

          </form>

<% end if %>

      </td>
  </tr>
</table>

 </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayButtons()

  response.write "<div style=""font-size:10px; padding-bottom:5px;"">" & vbcrlf
  response.write "  <input type=""button"" name=""sAction"" id=""cancel"" value=""Cancel"" class=""button"" onclick=""history.back();"" />" & vbcrlf
  response.write "  <input type=""submit"" name=""sAction"" id=""update"" value=""Update"" class=""button"" onclick=""return CheckCommitteeField();"" />" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = replace(p_value,"'","''")
  end if

  dbsafe = lcl_return

end function
%>
