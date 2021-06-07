<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!--#include file="../include_top_functions.asp"-->
<%
 lcl_orgid            = 0
 lcl_userid           = 0
 lcl_dm_ownerid       = 0
 lcl_isAdmin          = 0
 lcl_ownertype       = "EDITOR"
 lcl_lastModifiedDate = ""
 lcl_isAjax           = "N"

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("userid") <> "" then
    if isnumeric(request("userid")) then
       lcl_userid = request("userid")
    end if
 end if

 if request("dm_ownerid") <> "" then
    if isnumeric(request("dm_ownerid")) then
       lcl_dm_ownerid = request("dm_ownerid")
    end if
 end if

 if request("changetype") <> "" then
    lcl_ownertype = ucase(request("changetype"))
    lcl_ownertype = "'" & dbsafe(lcl_ownertype) & "'"
 end if

 lcl_lastModifiedDate = "'" & ConvertDateTimetoTimeZone(lcl_orgid) & "'"

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 end if

 if lcl_dm_ownerid <> "" then
    sSQL = "UPDATE egov_dm_owners SET "
    sSQL = sSQL & " ownertype = "               & lcl_ownertype & ", "
    sSQL = sSQL & " isLastModifiedByAdmin = "   & lcl_isAdmin   & ", "
    sSQL = sSQL & " lastmodifiedbyid = "        & lcl_userid    & ", "
    sSQL = sSQL & " lastmodifiedbydate = "      & lcl_lastModifiedDate
    sSQL = sSQL & " WHERE dm_ownerid = " & lcl_dm_ownerid

   	set oChangeOwnerType = Server.CreateObject("ADODB.Recordset")
  	 oChangeOwnerType.Open sSQL, Application("DSN"), 3, 1

    set oChangeOwnerType = nothing

    if lcl_isAjax = "Y" then
       if request("changetype") = "OWNER" then
          response.write "CTO"
       else
          response.write "CTE"
       end if
    end if

 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update section order - Error in AJAX Routine"
    'else
    '   response.write "datamgr_types_maint.asp?dm_typeid=" & lcl_dm_typeid & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>