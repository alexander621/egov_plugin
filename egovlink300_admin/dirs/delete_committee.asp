<!-- #include file="../includes/common.asp" //-->
<% 
 sLevel = "../" ' Override of value from common.asp

 if NOT UserHasPermission( Session("UserId"), "departments" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 if request("currentpage") <> "" AND isNumeric(request("currentpage")) then
	   currentpage = clng(request("currentpage"))

 			if clng(currentpage) < 1 then
    	  currentpage = 1
    end if

 else
	   currentpage = 1
 end if

'Delete/Inactivate each record selected
 for each delete in request.form("delete")
    	id=clng(delete)

    'Second check to see if this group has been assigned to any request forms that HAVE requests submitted again them.
     if checkGroupOnFormsExists(id) = "Y" then

       'If YES then inactivate the group.
        sSQL2 = "UPDATE groups SET isInactive = 1 "
        sSQL2 = sSQL2 & " WHERE groupid = " & id

     else

       'If NO then delete the group. 
        sSQL2 = "DELETE FROM groups WHERE groupid = " & id

     end if

     set rs2 = Server.CreateObject("ADODB.Recordset")
     rs2.Open sSQL2, Application("DSN"), 3, 1

     removeUsersOffGroup(id)
     removeGroupOffForms(id)

     set rs1 = nothing
     set rs2 = nothing

Next

response.redirect("display_committee.asp?currentpage=" & currentpage & "&success=SD")

'------------------------------------------------------------------------------
function checkGroupOnFormsExists(p_id)
  lcl_return = "N"

  sSQL = "SELECT COUNT(DISTINCT('Y')) as lcl_exists "
  sSQL = sSQL & " FROM egov_action_request_forms f, egov_actionline_requests r "
  sSQL = sSQL & " WHERE f.action_form_id = r.category_id "
  sSQL = sSQL & " AND r.orgid = " & session("orgid")
  sSQL = sSQL & " AND f.deptid = " & p_id

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

  if clng(rs("lcl_exists")) > clng(0) then
     lcl_return = "Y"
  end if

  set rs = nothing

  checkGroupOnFormsExists = lcl_return

end function

'------------------------------------------------------------------------------
sub removeUsersOffGroup(p_id)

    sSQLd2 = "DELETE FROM usersgroups WHERE groupid = " & id

    set rsd2 = Server.CreateObject("ADODB.Recordset")
    rsd2.Open sSQLd2, Application("DSN"), 3, 1

    set rsd2 = nothing

end sub

'------------------------------------------------------------------------------
sub removeGroupOffForms(p_id)

   sSQL = "UPDATE egov_action_request_forms SET deptid = NULL "
   sSQL = sSQL & " WHERE deptid = " & id

   set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sSQL, Application("DSN"), 3, 1

   set rs = nothing

end sub
%>
