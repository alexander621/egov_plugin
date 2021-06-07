<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
 lcl_userid             = 0
 lcl_orgid              = 0
 lcl_categoryid         = 0
 lcl_isApproved         = 0
 lcl_approvedDeniedDate = "'" & ConvertDateTimetoTimeZone() & "'"
 lcl_isAjax             = "N"

 if request("userid") <> "" then
    if isnumeric(request("userid")) then
       lcl_userid = request("userid")
    end if
 end if

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("categoryid") <> "" then
    if isnumeric(request("categoryid")) then
       lcl_categoryid = request("categoryid")
    end if
 end if

 if request("isApproved") then
    lcl_isApproved = 1
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 end if

 if lcl_categoryid <> "" then
    sSQL = "UPDATE egov_dm_categories SET "
    sSQL = sSQL & " isApproved = "          & lcl_isApproved & ", "
    sSQL = sSQL & " approvedeniedbyid = "   & lcl_userid     & ", "
    sSQL = sSQL & " approvedeniedbydate = " & lcl_approvedDeniedDate
    sSQL = sSQL & " WHERE categoryid = " & lcl_categoryid

   	set oApproveDenyCategory = Server.CreateObject("ADODB.Recordset")
  	 oApproveDenyCategory.Open sSQL, Application("DSN"), 3, 1

    set oApproveDenyCategory = nothing

    if lcl_isAjax = "Y" then
       if lcl_isApproved then
          response.write "approved"
       else
          response.write "denied"
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