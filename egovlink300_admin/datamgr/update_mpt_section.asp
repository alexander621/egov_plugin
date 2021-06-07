<!-- #include file="../includes/common.asp" //-->
<%
 lcl_orgid           = 0
 lcl_mappoint_typeid = 0
 lcl_sectionid       = 0
 lcl_sectionlocation = "L"
 lcl_sectionorder    = 1
 lcl_isActiveByOrg   = false
 lcl_isActiveByUser  = false

 if request("mappoint_typeid") <> "" then
    if isnumeric(request("mappoint_typeid")) then
       lcl_mappoint_typeid = request("mappoint_typeid")
    end if
 end if

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("sectionid") <> "" then
    if isnumeric(request("sectionid")) then
       lcl_sectionid = request("sectionid")
    end if
 end if

 if request("sectionlocation") <> "" then
    lcl_sectionlocation = ucase(request("sectionlocation"))
 end if

 if request("sectionorder") <> "" then
    if isnumeric(request("sectionorder")) then
       lcl_sectionorder = request("sectionorder")
    end if
 end if

 if request("isActiveByOrg") = "Y" then
    lcl_isActiveByOrg = true
 end if

 if request("isActiveByUser") = "Y" then
    lcl_isActiveByUser = true
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_mappoint_typeid > 0 then
    updateMPTSection lcl_orgid, lcl_mappoint_typeid, lcl_sectionid, lcl_sectionlocation, _
                     lcl_sectionorder, lcl_isActiveByOrg, lcl_isActiveByUser, lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update section order - Error in AJAX Routine"
    else
       response.write "mappoints_types_maint.asp?mappoint_typeid=" & lcl_mappoint_typeid & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub updateMPTSection(iOrgID, iMapPointTypeID, iSectionID, iSectionLocation, iSectionOrder, _
                     iSectionActiveByOrg, iSectionActiveByUser, iIsAjax)

  if iMapPointTypeID <> "" AND iSectionID <> "" then
     if iSectionLocation <> "" then
        sSectionLocation = ucase(iSectionLocation)
        sSectionlocation = "'" & dbsafe(sSectionLocation) & "'"
     else
        sSectionLocation = "'L'"
     end if

     if iSectionOrder <> "" then
        sSectionOrder = iSectionOrder
     else
        sSectionOrder = "1"
     end if

     if iSectionActiveByOrg then
        sSectionActiveByOrg = 1
     else
        sSectionActiveByOrg = 0
     end if

     if iSectionActiveByUser then
        sSectionActiveByUser = 1
     else
        sSectionActiveByUser = 0
     end if

    '1. need to check to see if the section exists for the mappoint type
    '2. if "no" then insert the record
    '3. if "yes" then update the record

     if iSectionID > 0 then

        checkSectionExistsOnMPT iOrgID, iMapPointTypeID, iSectionID, lcl_section_on_mpt

        if lcl_section_on_mpt then
           sSQL = "UPDATE egov_mappoints_types_sections SET "
           sSQL = sSQL & "sectionlocation = " & sSectionLocation
           sSQL = sSQL & ", sectionorder = "    & sSectionOrder

           if iSectionActiveByOrg then
              sSQL = sSQL & ", isActive_byOrg = "  & sSectionActiveByOrg
           end if

           if iSectionActiveByUser then
              sSQL = sSQL & ", isActive_byUser = "  & sSectionActiveByUser
           end if

           sSQL = sSQL & " WHERE mp_sectionid = " & iSectionID

           lcl_success   = "SU"
           lcl_isAjaxmsg = "Sucessfully Updated"

        else
           if iSectionActive then
              sSQL = "INSERT INTO egov_mappoints_types_sections ("
              sSQL = sSQL & "mappoint_typeid "
              sSQL = sSQL & ", orgid "
              sSQL = sSQL & ", sectionid "
              sSQL = sSQL & ", sectionlocation "
              sSQL = sSQL & ", sectionorder "

              if iSectionActiveByOrg then
                 sSQL = sSQL & ", isActive_byOrg "
              end if

              if iSectionActiveByUser then
                 sSQL = sSQL & ", isActive_byUser "
              end if

              sSQL = sSQL & ") VALUES ("
              sSQL = sSQL & iMapPointTypeID
              sSQL = sSQL & ", " & iOrgID
              sSQL = sSQL & ", " & iSectionID
              sSQL = sSQL & ", " & sSectionLocation
              sSQL = sSQL & ", " & sSectionOrder

              if iSectionActiveByOrg then
                 sSQL = sSQL & ", " & sSectionActiveByOrg
              end if

              if iSectionActiveByUser then
                 sSQL = sSQL & ", " & sSectionActiveByOrg
              end if

              sSQL = sSQL & ") "

              lcl_success   = "SU"
              lcl_isAjaxmsg = "Sucessfully Updated"

           end if

        end if

        if sSQL <> "" then
          	set oMaintainMPTSection = Server.CreateObject("ADODB.Recordset")
         	 oMaintainMPTSection.Open sSQL, Application("DSN"), 3, 1

           set oMaintainMPTSection = nothing
        end if

     else
        lcl_success   = "ERROR"
        lcl_isAjaxMsg = "ERROR: No SectionID"
     end if

  else
     lcl_success   = "ERROR"
     lcl_isAjaxMsg = "ERROR"
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  'else
  '   response.redirect "list_faq.asp?faqtype=" & iFAQType & "&success=" & lcl_success
  end if

end sub

'------------------------------------------------------------------------------
sub checkSectionExistsOnMPT(ByVal iOrgID, ByVal iMapPointTypeID, ByVal iSectionID, ByRef lcl_section_on_mpt)

  lcl_section_on_mpt = false

  if iMapPointTypeID <> "" AND iSectionID <> "" then
     sSQL = "SELECT mp_sectionid "
     sSQL = sSQL & " FROM egov_mappoints_types_sections "
     sSQL = sSQL & " WHERE orgid = " & iOrgID
     sSQL = sSQL & " AND mappoint_typeid = " & iMapPointTypeID
     sSQL = sSQL & " AND sectionid = " & iSectionID

    	set oCheckSectionExists = Server.CreateObject("ADODB.Recordset")
   	 oCheckSectionExists.Open sSQL, Application("DSN"), 3, 1

     if not oCheckSectionExists.eof then
        lcl_section_on_mpt = true
     end if

     oCheckSectionExists.close
     set oCheckSectionExists = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>