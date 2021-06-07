<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
 lcl_userid     = 0
 lcl_orgid      = 0
 lcl_mappointid = 0
 lcl_categoryid = 0
 lcl_isAjax     = "N"

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

 if request("mappointid") <> "" then
    if isnumeric(request("mappointid")) then
       lcl_mappointid = request("mappointid")
    end if
 end if

 if request("categoryid") <> "" then
    if isnumeric(request("categoryid")) then
       lcl_categoryid = request("categoryid")
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 end if

 if lcl_mappointid <> "" then
    sSQL = "UPDATE egov_mappoints SET "
    sSQL = sSQL & " categoryid = " & lcl_categoryid
    sSQL = sSQL & " WHERE mappointid = " & lcl_mappointid

   	set oUpdateMPCategory = Server.CreateObject("ADODB.Recordset")
  	 oUpdateMPCategory.Open sSQL, Application("DSN"), 3, 1

    if lcl_isAjax = "Y" then
       response.write "Successfully Updated..."
    end if
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update category - Error in AJAX Routine"
    'else
    '   response.write "mappoints_types_maint.asp?mappointid=" & lcl_mappointid & "&success=AJAX_ERROR"
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