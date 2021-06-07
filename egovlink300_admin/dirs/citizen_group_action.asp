<!-- #include file="../includes/common.asp" //-->
<%
 if request("action") = "DELETE" then

    lcl_delete_count = 0
    lcl_success_msg  = ""

    for each delete in request("delete")
       	lcl_groupid = delete

        if lcl_groupid <> "" then
           lcl_delete_count = lcl_delete_count + 1

          	sSQL = "DELETE FROM citizengroups WHERE groupid=" & lcl_groupid

          	set oDeleteCitizenGroup = Server.CreateObject("ADODB.Recordset")
          	oDeleteCitizenGroup.Open sSQL, Application("DSN"), 3, 1

           set oDeleteCitizenGroup = nothing
        end if
    next

    if lcl_delete_count > 0 then
       lcl_success_msg = "?success=SD"
    end if

    lcl_redirect_url = "display_citizen_groups.asp" & lcl_success_msg

 else
   'Setup values
    if request("groupid") <> "" then
       lcl_groupid = request("groupid")
    else
       lcl_groupid = 0
    end if

    if request("groupname") <> "" then
       lcl_groupname = "'" & dbsafe(request("groupname")) & "'"
    else
       lcl_groupname = "NULL"
    end if

    if request("groupdescription") <> "" then
       lcl_groupdescription = "'" & dbsafe(request("groupdescription")) & "'"
    else
       lcl_groupdescription = "NULL"
    end if

    if request("grouptype") <> "" then
       lcl_grouptype = "'" & dbsafe(request("grouptype")) & "'"
    else
       lcl_grouptype = "NULL"
    end if

    if request("orgid") <> "" then
       lcl_orgid = request("orgid")
    else
       lcl_orgid = 0
    end if

    if lcl_groupid > 0 then

      	sSQL = "update citizengroups set "
       sSQL = sSQL & "groupname = "        & lcl_groupname        & ", "
       sSQL = sSQL & "groupdescription = " & lcl_groupdescription & ", "
       sSQL = sSQL & "grouptype = "        & lcl_grouptype
       sSQL = sSQL & " WHERE groupid = " & lcl_groupid

      	set oUpdateCitizenGroup = Server.CreateObject("ADODB.Recordset")
      	oUpdateCitizenGroup.Open sSQL, Application("DSN"), 3, 1

       set oUpdateCitizenGroup = nothing

       lcl_redirect_url = "citizen_group_maint.asp?groupid=" & lcl_groupid & "&success=SU"

    else
       sSQL = "INSERT INTO citizengroups ("
       sSQL = sSQL & "orgid, "
       sSQL = sSQL & "groupname, "
       sSQL = sSQL & "groupdescription, "
       sSQL = sSQL & "grouptype "
       sSQL = sSQL & ") VALUES ("
       sSQL = sSQL & lcl_orgid            & ", "
       sSQL = sSQL & lcl_groupname        & ", "
       sSQL = sSQL & lcl_groupdescription & ", "
       sSQL = sSQL & lcl_grouptype
       sSQL = sSQL & ") "

      	set oInsertCitizenGroup = Server.CreateObject("ADODB.Recordset")
      	oInsertCitizenGroup.Open sSQL, Application("DSN"), 3, 1

       set oInsertCitizenGroup = nothing

       lcl_redirect_url = "display_citizen_groups.asp?success=SA"

    end if
 end if

 response.redirect lcl_redirect_url
%>