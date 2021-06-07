<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="../class/classOrganization.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
 Dim oUpdateDMValue

 set oUpdateDMValue = New classOrganization

 lcl_userid       = 0
 lcl_orgid        = 0
 lcl_dm_typeid    = 0
 lcl_dmid         = 0
 lcl_dm_sectionid = 0
 lcl_dm_fieldid   = 0
 lcl_dm_valueid   = 0
 lcl_fieldvalue   = ""
 lcl_isAjax       = "N"

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

 if request("dm_typeid") <> "" then
    if isnumeric(request("dm_typeid")) then
       lcl_dm_typeid = request("dm_typeid")
    end if
 end if

 if request("dmid") <> "" then
    if isnumeric(request("dmid")) then
       lcl_dmid = request("dmid")
    end if
 end if

 if request("dm_sectionid") <> "" then
    if isnumeric(request("dm_sectionid")) then
       lcl_dm_sectionid = request("dm_sectionid")
    end if
 end if

 if request("dm_fieldid") <> "" then
    if isnumeric(request("dm_fieldid")) then
       lcl_dm_fieldid = request("dm_fieldid")
    end if
 end if

 if request("dm_valueid") <> "" then
    if isnumeric(request("dm_valueid")) then
       lcl_dm_valueid = request("dm_valueid")
    end if
 end if

 if request("fieldvalue") <> "" then
    lcl_fieldvalue = request("fieldvalue")
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 end if

 if lcl_dm_valueid <> "" then
    maintainDMValues lcl_userid, lcl_orgid, lcl_dm_typeid, lcl_dmid, lcl_dm_sectionid, _
                     lcl_dm_fieldid, lcl_dm_valueid, lcl_fieldvalue

    if lcl_isAjax = "Y" then
       response.write "Success"
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