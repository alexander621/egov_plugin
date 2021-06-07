<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
 lcl_orgid         = 0
 lcl_classid       = 0
 lcl_accessorytype = ""
 lcl_isAjaxRoutine = "Y"

 if request("orgid") <> "" then
    lcl_orgid = request("orgid")
 end if

 if request("classid") <> "" then
    lcl_classid = request("classid")
 end if

 if request("accessorytype") <> "" then
    lcl_accessorytype = request("accessorytype")
 end if

 if request("isAjaxRoutine") <> "" then
    lcl_isAjaxRoutine = request("isAjaxRoutine")
 end if
 
 if lcl_orgid > 0 AND lcl_classid > 0 AND lcl_accessorytype <> "" then
    assignDefaultAccessoryOptions lcl_orgid, lcl_classid, lcl_accessorytype, lcl_isAjaxRoutine
 end if

'------------------------------------------------------------------------------
sub assignDefaultAccessoryOptions(p_orgid, p_classid, p_accessorytype, lcl_isAjaxRoutine)
  lcl_total_accessories = 0

 'Check to see if this accessory type for this class already has values assigned
  'sSQL = "SELECT COUNT(*) as total_accessories "
  'sSQL = sSQL & " FROM egov_class_teamroster_accessories_to_class "
  'sSQL = sSQL & " WHERE isDisabled = 0 "
  'sSQL = sSQL & " AND orgid = " & p_orgid
  'sSQL = sSQL & " AND classid = " & p_classid

  sSQL = "SELECT COUNT(*) as total_accessories "
  sSQL = sSQL & " FROM egov_class_teamroster_accessories_to_class "
  sSQL = sSQL & " WHERE orgid = " & p_orgid
  sSQL = sSQL & " AND classid = " & p_classid

  set oTotalAccessories = Server.CreateObject("ADODB.Recordset")
  oTotalAccessories.Open sSQL, Application("DSN"), 0, 1

  if not oTotalAccessories.eof then
     lcl_total_accessories = oTotalAccessories("total_accessories")
  end if

  oTotalAccessories.close
  set oTotalAccessories = nothing

 'If no accessories have been assigned then assign the defaults.
 'Defaults are the records on "egov_class_teamroster_accessories_to_class" with "orgid = 0" and "classid = 0"
  if lcl_total_accessories < 1 then
     sSQL = "SELECT accessoryid  "
     sSQL = sSQL & " FROM egov_class_teamroster_accessories "
     sSQL = sSQL & " WHERE UPPER(accessorytype) = '" & UCASE(p_accessorytype) & "' "
     sSQL = sSQL & " ORDER BY displayorder "

     set oDefaultAccessories = Server.CreateObject("ADODB.Recordset")
     oDefaultAccessories.Open sSQL, Application("DSN"), 3, 1

     lcl_count = 0

     if not oDefaultAccessories.eof then
        do while not oDefaultAccessories.eof
           lcl_count = lcl_count + 1

           sSQL = "INSERT INTO egov_class_teamroster_accessories_to_class ("
           sSQL = sSQL & "orgid, "
           sSQL = sSQL & "classid, "
           sSQL = sSQL & "accessoryid "
           sSQL = sSQL & ") VALUES ("
           sSQL = sSQL & p_orgid   & ", "
           sSQL = sSQL & p_classid & ", "
           sSQL = sSQL & oDefaultAccessories("accessoryid")
           sSQL = sSQL & ") "

           set oInsertAccessories = Server.CreateObject("ADODB.Recordset")
           oInsertAccessories.Open sSQL, Application("DSN"), 3, 1

           set oInsertAccessories = nothing

           oDefaultAccessories.movenext
        loop
     end if

     oDefaultAccessories.close
     set oDefaultAccessories = nothing
  end if

  if lcl_isAjaxRoutine = "Y" then
     response.write "success"
  end if

end sub

'--------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1
end sub
%>