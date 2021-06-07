<%
'------------------------------------------------------------------------------
'sub buildMapPointLayout(iLayoutID, iMapPointTypeID, iMapPointID, iDisplayFieldsetLegend, _
sub buildDMLayout(iLayoutID, iDMTypeID, iDMID, iDisplayFieldsetLegend, iDisplayFieldsetBorder, _
                  iDisplayAvailableSections, iSectionMode)

 'Setup how to display the FieldSets that hold the layout and unused sections
  sDisplayFieldsetLegend    = True
  sDisplayFieldsetBorder    = True
  sDisplayAvailableSections = True
  sSectionMode              = ""
  lcl_fieldset_class        = "layoutfieldset"

  if iDisplayFieldsetLegend <> "" then
     sDisplayFieldsetLegend = iDisplayFieldsetLegend
  end if

  if iDisplayFieldsetBorder <> "" then
     sDisplayFieldsetBorder = iDisplayFieldsetBorder
  end if

  if not sDisplayFieldsetBorder then
     lcl_fieldset_class = "layoutfieldset_noborder"
  end if

  if iDisplayAvailableSections <> "" then
     sDisplayAvailableSections = iDisplayAvailableSections
  end if

  if iSectionMode <> "" then
     sSectionMode = ucase(iSectionMode)
  end if

 'Determine which layout to build
  if iLayoutID <> "" then
     if checkLayoutExists(iLayoutID) then
        lcl_layoutid = iLayoutID
     end if
  else
     lcl_layoutid = getOriginalLayoutID()
  end if

  getLayoutInfo lcl_layoutid, lcl_layoutname, lcl_isOriginalLayout, lcl_useLayoutSections, _
                lcl_totalcolumns, lcl_columnwidth_left, lcl_columnwidth_middle, lcl_columnwidth_right

  if lcl_totalcolumns < 1 then
     lcl_totalcolumns = 1
  end if

 'Determine the total width of the table (in pixels)
  lcl_totalLayoutWidth = lcl_columnwidth_left + lcl_columnwidth_middle + lcl_columnwidth_right

 'BEGIN: Display all active sections in current MapPoint Type Layout ----------
  response.write "<fieldset class=""" & lcl_fieldset_class & """>" & vbcrlf

  if lcl_displayFieldsetLegend then
     response.write "  <legend>Current Layout</legend>" & vbcrlf
  end if

  'response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""respTable"" style=""max-width:" & lcl_totalLayoutWidth & "px;"">" & vbcrlf
  'response.write "    <tr valign=""top"">" & vbcrlf

                          lcl_displayUnusedSections = false
                          lcl_wrapWithTDTags        = true
                          lcl_total_draggable_items = 0

                          buildLayoutColumns lcl_totalcolumns, "L", lcl_wrapWithTDTags, iDMTypeID, iDMID, _
                                             lcl_columnwidth_left, lcl_displayUnusedSections, lcl_total_draggable_items, _
                                             sSectionMode, lcl_total_draggable_items

                          if lcl_totalcolumns > 1 then
                             if lcl_totalcolumns > 2 then
                                buildLayoutColumns lcl_totalcolumns, "M", lcl_wrapWithTDTags, iDMTypeID, iDMID, _
                                                   lcl_columnwidth_middle, lcl_displayUnusedSections, lcl_total_draggable_items, _
                                                   sSectionMode, lcl_total_draggable_items
                             end if

                             buildLayoutColumns lcl_totalcolumns, "R", lcl_wrapWithTDTags, iDMTypeID, iDMID, _
                                                lcl_columnwidth_right, lcl_displayUnusedSections, lcl_total_draggable_items, _
                                                sSectionMode, lcl_total_draggable_items
                          end if

  'response.write "    </tr>" & vbcrlf
  'response.write "  </table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
 'END: Display all active sections in current MapPoint Type Layout ------------

  response.write "<p>&nbsp;</p>" & vbcrlf

 'BEGIN: Display all "unused" and "active" sections ---------------------------
  if sDisplayAvailableSections then
     response.write "<p>" & vbcrlf
     response.write "<fieldset class=""" & lcl_fieldset_class & """>" & vbcrlf

     if lcl_displayFieldsetLegend then
        response.write "  <legend>Available Sections</legend>" & vbcrlf
     end if

     response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:" & lcl_totalLayoutWidth/2 & "px;"">" & vbcrlf
     response.write "    <tr>" & vbcrlf
     response.write "        <td class=""column"" id=""column" & lcl_totalcolumns + 1 & """ colspan=""" & lcl_totalcolumns & """>" & vbcrlf
                                 lcl_displayUnusedSections = true
                                 lcl_wrapWithTDTags        = false
                                 lcl_totalcolumns          = lcl_totalcolumns + 1

                                 buildLayoutColumns lcl_totalcolumns, "L", lcl_wrapWithTDTags, iDMTypeID, iDMID, _
                                                    lcl_columnwidth_left, lcl_displayUnusedSections, lcl_total_draggable_items, _
                                                    sSectionMode, lcl_total_draggable_items

     response.write "        </td>" & vbcrlf
     response.write "    </tr>" & vbcrlf
     response.write "  </table>" & vbcrlf
     response.write "</fieldset>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if
 'END: Display all "unused" and "active" sections -----------------------------

  response.write "<input type=""hidden"" name=""totalcolumns"" id=""totalcolumns"" value=""" & lcl_totalcolumns & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "<input type=""hidden"" name=""totalitems"" id=""totalitems"" value=""" & lcl_total_draggable_items & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildLayoutColumns(ByVal iTotalColumns, ByVal iColumnLocation, ByVal iWrapWithTDTags, ByVal iDMTypeID, _
                       ByVal iDMID, ByVal iColumnWidth, ByVal iDisplayUnusedSections, ByVal iTotalDraggableItems, _
                       ByVal iSectionMode, ByRef lcl_total_draggable_items)

  if iTotalColumns <> "" then
     lcl_totalcolumns = iTotalColumns
  else
     lcl_totalcolumns = 1
  end if

  if iSectionMode <> "" then
     sSectionMode = ucase(iSectionMode)
  else
     sSectionMode = ""
  end if

 'Based on the total columns, determine if we need to look for "hidden" sections
 '  (i.e. sections that would display in the "middle" of a 3-column layout, but not appear in a 2-column layout
 'L = Left, M = Middle, R = Right
  if iColumnLocation <> "" then
     if iColumnLocation = "L" then
        lcl_columnlocation = ucase(iColumnLocation)

        if lcl_totalcolumns = 2 then
           lcl_columnlocation_sql = "'L','M'"
        elseif lcl_totalcolumns < 2 then
           lcl_columnlocation_sql = "'L','M','R'"
        else
           lcl_columnlocation_sql = "'L'"
        end if
     else
         lcl_columnlocation     = ucase(iColumnLocation)
         lcl_columnlocation_sql = "'" & dbsafe(lcl_columnlocation) & "'"
     end if
  else
     lcl_columnlocation     = "L"
     lcl_columnlocation_sql = "'L'"
  end if

 'Get the column number for the ID
  lcl_columnnumber = getColumnNumber(lcl_totalcolumns, lcl_columnlocation)

  if iColumnWidth > 0 then
     if iWrapWithTDTags then
        'response.write "<td class=""column"" id=""column" & lcl_columnnumber & """ style=""width:" & iColumnWidth & "px;"">" & vbcrlf
        response.write "<div class=""layoutcolumn"" style=""width:" & iColumnWidth & "px;"">" & vbcrlf
     end if

    'Retrieve all of the sections for the column identified
    'OR retrieve all of the "unused" and "active" sections available
    '*** NOTE: in the "unused" and "active" sections, we do NOT want to pull any sections that are "Account Info" sections.
              'These sections are to be used elsewhere in the code and are set up individually for each MapPoint Type.
     if iDisplayUnusedSections then
        sSQL = "SELECT 0 as dm_sectionid, "
        sSQL = sSQL & " dms.sectionid, "
        sSQL = sSQL & " dms.sectionname, "
        sSQL = sSQL & " dms.sectiontype, "
        sSQL = sSQL & " dms.isAccountInfoSection, "
        sSQL = sSQL & " 'N' as sectionIsActive "
        sSQL = sSQL & " FROM egov_dm_sections dms "
        sSQL = sSQL & " WHERE dms.sectionid NOT IN (select distinct dmts.sectionid "
        sSQL = sSQL &                              " from egov_dm_types_sections dmts "
        sSQL = sSQL &                              " where dmts.dm_typeid = " & iDMTypeID
        sSQL = sSQL &                              " and dmts.isActive = 1) "
        sSQL = sSQL & " AND dms.isActive = 1 "
        'sSQL = sSQL & " AND dms.isAccountInfoSection = 0 "
        sSQL = sSQL & " ORDER BY sectionname "
     else
        'sSQL = "SELECT dmts.dm_sectionid as sectionid, "
        sSQL = "SELECT dmts.dm_sectionid, "
        sSQL = sSQL & " dmts.sectionid, "
        sSQL = sSQL & " dms.sectionname, "
        sSQL = sSQL & " dms.sectiontype, "
        sSQL = sSQL & " dms.isAccountInfoSection, "
        sSQL = sSQL & " 'Y' as sectionIsActive "
        sSQL = sSQL & " FROM egov_dm_types_sections dmts "
        sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
        sSQL = sSQL &                 " ON dmts.sectionid = dms.sectionid "
        sSQL = sSQL &                 " AND dms.isActive = 1 "
        sSQL = sSQL & " WHERE dmts.dm_typeid = " & iDMTypeID
        sSQL = sSQL & " AND upper(dmts.sectionlocation) IN (" & lcl_columnlocation_sql & ") "
        sSQL = sSQL & " AND dmts.isActive = 1 "
        sSQL = sSQL & " ORDER BY sectionorder "
     end if
'response.write sSQL  'LEFT OFF HERE!!!!
     set oBuildSectionColumn = Server.CreateObject("ADODB.Recordset")
     oBuildSectionColumn.Open sSQL, Application("DSN"), 3, 1

     if not oBuildSectionColumn.eof then
        lcl_sectioncount          = 0
        lcl_total_draggable_items = iTotalDraggableItems

        do while not oBuildSectionColumn.eof
           lcl_sectioncount          = lcl_sectioncount + 1
           lcl_total_draggable_items = lcl_total_draggable_items + 1

           buildSection iDMTypeID, iDMID, _
                        oBuildSectionColumn("dm_sectionid"), _
                        oBuildSectionColumn("sectionid"), _
                        oBuildSectionColumn("sectionname"), _
                        oBuildSectionColumn("sectiontype"), _
                        oBuildSectionColumn("sectionIsActive"), _
                        oBuildSectionColumn("isAccountInfoSection"), _
                        lcl_columnlocation, lcl_sectioncount, lcl_total_draggable_items, sSectionMode

           oBuildSectionColumn.movenext
        loop
     end if

     oBuildSectionColumn.close
     set oBuildSectionColumn = nothing

'CANNOT DO THIS YET BECAUSE WE DO NOT HAVE A DM_FIELDID
    'Now we need to check to see if there is a MapPoints Value record for this field.
    'If not then we need to create one.
'     if iDMID <> "" then
'        lcl_dmvalue_exists = checkdmvalueExists(lcl_dm_fieldid)

'        if not lcl_dmvalue_exists then
'           lcl_dmid = 0

'           maintainMapPointValues(iOrgID, iDMTypeID, iDMID, iDMSectionID, iMPFieldID, idmvalueID, iFieldValue)
'        end if
'     end if

     if iWrapWithTDTags then
        'response.write "</td>" & vbcrlf
        response.write "</div>" & vbcrlf
     end if
  end if

end sub

'------------------------------------------------------------------------------
sub buildSection(iDMTypeID, iDMID, iDMSectionID, iSectionID, iSectionName, iSectionType, _
                 iSectionIsActive, iIsAccountInfoSection, iSectionLocation, iSectionOrder, _
                 iTotalDraggableItems, iSectionMode)

  if iSectionID <> "" then
     lcl_sectiontype          = ""
     lcl_sectionmode          = ""
     lcl_isAccountInfoSection = False

     if iSectionType <> "" then
        lcl_sectiontype = ucase(iSectionType)
     end if

     if iSectionMode <> "" then
        lcl_sectionmode = ucase(iSectionMode)
     end if

     if iIsAccountInfoSection <> "" then
        lcl_isAccountInfoSection = iIsAccountInfoSection
     end if

    'Determine which section to build
     if lcl_sectionmode = "DRAG" then
        displaySection_noInfo iDMTypeID, iDMSectionID, iSectionID, iSectionName, iSectionIsActive, _
                              iSectionLocation, iSectionOrder, iTotalDraggableItems
     else
        if lcl_sectiontype = "GOOGLE_MAP_DOT" then
           'response.write "<div class=""section"" id=""section" & iSectionID & """>" & vbcrlf
           response.write "<div id=""map_canvas_dot_navigation"" align=""right""><a href=""http://maps.google.com/support/bin/static.py?page=guide.cs&guide=21670&topic=21671&answer=144350"" color=""#0000ff"" target=""_blank"">How to navigate in Google Maps</a></div>" & vbcrlf
           response.write "<div id=""map_canvas_dot"">&nbsp;</div>" & vbcrlf
           response.write "<div id=""map_canvas_dot_getDirections"" align=""center""><input type=""button"" value=""Get Directions"" class=""button"" onclick=""openGoogleURL();"" /></div>" & vbcrlf
           'response.write "</div>" & vbcrlf
        elseif lcl_sectiontype = "GOOGLE_MAP_STREETVIEW" then
           'response.write "<div class=""section"" id=""section" & iSectionID & """>" & vbcrlf
           response.write "<div id=""map_canvas_streetview_navigation"" align=""right""><a href=""http://maps.google.com/support/bin/static.py?page=guide.cs&guide=21670&topic=21671&answer=144350"" color=""#0000ff"" target=""_blank"">How to navigate in Google Maps</a></div>" & vbcrlf
           response.write "<div id=""map_canvas_streetview"">&nbsp;</div>" & vbcrlf
           response.write "<div id=""map_canvas_streetview_getDirections"" align=""center""><input type=""button"" value=""Get Directions"" class=""button"" onclick=""openGoogleURL();"" /></div>" & vbcrlf
           response.write "<div id=""map_canvas_streetview_note"" class=""redText"" align=""center"">NOTE: Street-Level View is approximate.  You may need to rotate the image</div>&nbsp;" & vbcrlf
           'response.write "</div>" & vbcrlf
        elseif lcl_sectiontype = "CATEGORIES_SECTION" then
           displaySection_categories iSectionID, iSectionName, iDMID
        else
           displaySection_default iDMID, iDMTypeID, iSectionID, lcl_isAccountInfoSection, iSectionMode
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
sub displaySection_noInfo(iDMTypeID, iDMSectionID, iSectionID, iSectionName, _
                          iSectionActive, iSectionLocation, iSectionOrder, iTotalDraggableItems)

 'Do NOT display the "Default" section name
  lcl_sectionname     = "&nbsp;"
  lcl_sectionlocation = "L"

  if iSectionName <> "" then
     if ucase(iSectionName) <> "DEFAULT" then
        lcl_sectionname = iSectionName
     end if
  end if

  if iSectionLocation <> "" then
     lcl_sectionlocation = ucase(iSectionLocation)
  end if

  response.write "<div class=""dragbox"" id=""dragbox" & iTotalDraggableItems & """>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"" class=""dragbox-content"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <h2>" & lcl_sectionname & "</h2>" & vbcrlf
  response.write "          <input type=""hidden"" name=""dm_sectionid_"    & iTotalDraggableItems & """ id=""dm_sectionid_"    & iTotalDraggableItems & """ value=""" & iDMSectionID        & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""sectionid_"       & iTotalDraggableItems & """ id=""sectionid_"       & iTotalDraggableItems & """ value=""" & iSectionID          & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""sectionlocation_" & iTotalDraggableItems & """ id=""sectionlocation_" & iTotalDraggableItems & """ value=""" & lcl_sectionlocation & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""sectionorder_"    & iTotalDraggableItems & """ id=""sectionorder_"    & iTotalDraggableItems & """ value=""" & iSectionOrder       & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""sectionactive_"   & iTotalDraggableItems & """ id=""sectionactive_"   & iTotalDraggableItems & """ value=""" & iSectionActive      & """ />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td align=""center""><p>[Click on Section Header]<br />to drag-n-drop</p></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displaySection_default(iDMID, iDMTypeID, iSectionID, iIsAccountInfoSection, iSectionMode)

  dim lcl_scripts, lcl_hide_section

  lcl_hide_section = "Y"

 'Get all of the section info
  getSectionInfo iSectionMode, iSectionID, lcl_sectionname, lcl_sectiontype

  response.write "<div class=""section"" id=""section" & iSectionID & """>" & vbcrlf

'  response.write "<div class=""section"" id=""section" & iSectionID & """>" & vbcrlf
'  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""section-content"">" & vbcrlf
'  response.write "  <input type=""hidden"" name=""sectionid_" & iSectionID & """ id=""sectionid_" & iSectionID & """ value=""" & iSectionID & """ />" & vbcrlf

 'Do NOT display the "Default" section name
  'if ucase(lcl_sectionname) = "DEFAULT" then
  '   lcl_sectionname = ""
  'end if

 'Only show the sectionname row if a sectionname exists.
'  if lcl_sectionname <> "" then
'     response.write "  <tr>" & vbcrlf
'     response.write "      <th align=""left"">" & vbcrlf
'     response.write "          <h2>" & lcl_sectionname & "</h2>" & vbcrlf
'     response.write "      </th>" & vbcrlf
'     response.write "  </tr>" & vbcrlf
'  end if

'  response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf

  sSQL = "SELECT dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQL = sSQL & " dmtf.section_fieldid, "
  sSQL = sSQL & " dmtf.displayFieldName, "
  sSQL = sSQL & " dmsf.fieldname, "
  sSQL = sSQL & " dmsf.fieldtype, "
  sSQL = sSQL & " dmsf.isMultiLine, "
  sSQL = sSQL & " dmsf.hasAddLinkButton, "

  if iDMID <> "" then
     sSQL = sSQL & " dmv.fieldvalue "
  else
     sSQL = ssQL & " '' as fieldvalue "
  end if

  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &                 " ON dmsf.section_fieldid = dmtf.section_fieldid "
  sSQL = sSQL &                 " AND dmsf.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &                 " ON dmts.dm_sectionid = dmtf.dm_sectionid "
  sSQL = sSQL &                 " AND dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                 " AND dmts.sectionid = " & iSectionID
  sSQL = sSQL &                 " AND dmts.isActive = 1 "

  if iDMID <> "" then
     sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_values dmv "
     sSQL = sSQL &                      " ON dmv.dm_fieldid = dmtf.dm_fieldid "
     sSQL = sSQL &                      " AND dmv.dm_typeid = " & iDMTypeID
     sSQL = sSQL &                      " AND dmv.dmid = " & iDMID
  end if

  sSQL = sSQL & " WHERE dmtf.dm_typeid = " & iDMTypeID
  sSQL = sSQL & " AND dmtf.displayInInfoPage = 1 "
  sSQL = sSQL & " ORDER BY dmsf.displayOrder, dmtf.resultsOrder "

  set oEditDMTSectionFields = Server.CreateObject("ADODB.Recordset")
  oEditDMTSectionFields.Open sSQL, Application("DSN"), 3, 1

  if not oEditDMTSectionFields.eof then
     iRowCount              = 0
     lcl_field_maxlength    = "4000"
     lcl_previous_fieldname = ""
     lcl_scripts            = ""

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""section-content"">" & vbcrlf
     response.write "  <input type=""hidden"" name=""sectionid_" & iSectionID & """ id=""sectionid_" & iSectionID & """ value=""" & iSectionID & """ />" & vbcrlf

    'Only show the sectionname row if a sectionname exists.
     if lcl_sectionname <> "" then
        response.write "  <tr>" & vbcrlf
        response.write "      <th align=""left"">" & vbcrlf
        response.write "          <h2>" & lcl_sectionname & "</h2>" & vbcrlf
        response.write "      </th>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

     response.write "  <tr>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf

     do while not oEditDMTSectionFields.eof
        iRowCount      = iRowCount + 1
        lcl_fieldvalue = ""
        lcl_fieldname  = ""
        lcl_fieldtype  = ""

       'ONLY display the row if a value exists OR if we are in "EDIT" mode.
        if oEditDMTSectionFields("fieldvalue") <> "" OR iSectionMode = "EDIT" then
           lcl_hide_section = "N"

           if oEditDMTSectionFields("fieldtype") <> "" then
              lcl_fieldtype = oEditDMTSectionFields("fieldtype")
           end if

           if oEditDMTSectionFields("fieldname") <> "" then
              lcl_fieldname = oEditDMTSectionFields("fieldname")
           end if

           if oEditDMTSectionFields("fieldvalue") <> "" then
              lcl_fieldvalue = oEditDMTSectionFields("fieldvalue")

              if instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
                 lcl_fieldvalue = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)
              else
                 lcl_fieldvalue = replace(lcl_fieldvalue,chr(13),"<br />")
                 lcl_fieldvalue = replace(lcl_fieldvalue,chr(10),"")
              end if
           end if

          'If the user is in "EDIT" mode then we need to show all of the fields even if they do not have data in them.
          'ONLY when the user is NOT in "EDIT" mode do we hide fields that have no data.
           if iSectionMode = "EDIT" then
              lcl_fieldname = "<strong>" & lcl_fieldname & "</strong>"
           else
              if oEditDMTSectionFields("displayFieldName") then
                 lcl_fieldname = "<strong>" & lcl_fieldname & "</strong>"
              else
                'Trying to determine if the previous row had a label displayed or not.
                'If it DID then we need to put a filler in for this row.
                 if lcl_previous_fieldname <> "" AND iRowCount > 1 then
                    lcl_fieldname = "&nbsp;"
                 else
                    lcl_fieldname = ""
                 end if
              end if
           end if

           'if oEditDMTSectionFields("fieldtype") <> "" then
           '   lcl_fieldtype = oEditDMTSectionFields("fieldtype")
           'end if

           response.write "  <tr valign=""top"">" & vbcrlf

           if lcl_fieldname <> "" then
              response.write "      <td nowrap=""nowrap"">" & lcl_fieldname & "</td>" & vbcrlf
              response.write "      <td width=""100%"">" & lcl_fieldvalue
           else
              response.write "      <td colspan=""2"">" & lcl_fieldvalue
           end if

          'If this field has a fieldtype = "ADDRESS" then we need to build a hidden field
          'to store the value in for google maps link(s)
           if lcl_fieldtype = "ADDRESS" _
           OR lcl_fieldtype = "CITY" _
           OR lcl_fieldtype = "STATE" then
              lcl_google_address = replace(lcl_fieldvalue," ","+")

              response.write "<input type=""hidden"" name=""google" & lcl_fieldtype & """ id=""google" & lcl_fieldtype & """ value=""" & lcl_google_address & """ size=""20"" />" & vbcrlf
              'response.write "          <input type=""button"" name=""openGoogleMap"" id=""openGoogleMap"" class=""button"" value=""Open in Google Maps"" onclick=""openGoogleMap()"" />" & vbcrlf
           end if

           response.write "      </td>" & vbcrlf
           response.write "  </tr>" & vbcrlf

           lcl_previous_fieldname = lcl_fieldname
        end if

        oEditDMTSectionFields.movenext
     loop

     response.write "          </table>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
  end if

  oEditDMTSectionFields.close
  set oEditDMTSectionFields = nothing

'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf
'  response.write "</div>" & vbcrlf

 'Determine which "mode" we are in to determine which buttons are displayed
  if iSectionMode = "PUBLIC_VIEW" then
     lcl_iMode = iSectionMode
  else
     if iIsAccountInfoSection then
        lcl_iMode = "ACCOUNTINFO_DRAGDROP"
     else
        if iRowCount > 0 then
           lcl_iMode = "SECTION_VIEW"
        else
           lcl_iMode = "SECTION_VIEW NOEDIT"
        end if
     end if
  end if

  displayButtonsSection lcl_iMode, iDMTypeID, iSectionID, lcl_isActiveByUser

  response.write "</div>" & vbcrlf

  if lcl_hide_section = "Y" AND iSectionMode <> "EDIT" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write "  document.getElementById(""section" & iSectionID & """).style.display='none';" & vbcrlf
     response.write "</script>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub displaySection_categories(p_SectionID, p_SectionName, p_DMID)

  dim sSectionID, sSectionName, sDMID, sParentCategoryID, iRowCount

  sSectionID        = 0
  sSectionName      = p_SectionName
  sDMID             = 0
  sParentCategoryID = 0
  iRowCount         = 0

  if p_SectionID <> "" then
     sSectionID = clng(p_SectionID)
  end if

  if p_DMID <> "" then
     sDMID = clng(p_DMID)
  end if

  getCategoryInfoByDMID sDMID, sParentCategoryID, sParentCategoryName

  if sParentCategoryName <> "" then
     response.write "<div class=""section"" id=""section" & sSectionID & """>" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""section-content"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">" & vbcrlf
     response.write "          <input type=""hidden"" name=""sectionid_" & sSectionID & """ id=""sectionid_" & sSectionID & """ value=""" & sSectionID & """ />" & vbcrlf
     response.write "          <h2>" & sSectionName & "</h2>" & vbcrlf
     response.write "      </th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf        
     response.write "      <td style=""padding-left:4px"">" & vbcrlf
     response.write "          <strong>Category: </strong>" & sParentCategoryName & "<br />" & vbcrlf

     sSQL = "SELECT dtc.dmid_categoryid, "
     sSQL = sSQL & " dtc.dm_typeid, "
     sSQL = sSQL & " dtc.categoryid, "
     sSQL = sSQL & " c.categoryname "
     sSQL = sSQL & " FROM egov_dmdata_to_dmcategories dtc "
     sSQL = sSQL &      " INNER JOIN egov_dm_categories c ON c.categoryid = dtc.categoryid "
     sSQL = sSQL & " WHERE dtc.dmid = " & sDMID
     sSQL = sSQL & " AND c.parent_categoryid = " & sParentCategoryID
     sSQL = sSQL & " AND c.isActive = 1 "
     sSQL = sSQL & " AND c.isApproved = 1 "
     sSQL = sSQL & " ORDER BY UPPER(c.categoryname) "

     set oDisplayCategoryInfo = Server.CreateObject("ADODB.Recordset")
     oDisplayCategoryInfo.Open sSQL, Application("DSN"), 3, 1

     if not oDisplayCategoryInfo.eof then
        do while not oDisplayCategoryInfo.eof

           iRowCount = iRowCount + 1

           if iRowCount = 1 then
              response.write "<strong>Sub-Categories: </strong>" & oDisplayCategoryInfo("categoryname") & vbcrlf 
           end if

           if iRowCount > 1 then
              response.write ", " & oDisplayCategoryInfo("categoryname")
           end if

           oDisplayCategoryInfo.movenext
        loop

     end if

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf

     oDisplayCategoryInfo.close
     set oDisplayCategoryInfo = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub getCategoryInfoByDMID(ByVal iDMID, ByRef sParentCategoryID, ByRef sParentCategoryName)
  dim lcl_return, sDMID

  lcl_return          = 0
  sParentCategoryID   = 0
  sParentCategoryName = ""

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  sSQL = "SELECT dmd.categoryid, "
  sSQL = sSQL & " c.categoryname "
  sSQL = sSQL & " FROM egov_dm_data dmd "
  sSQL = sSQL &      " INNER JOIN egov_dm_categories c ON c.categoryid = dmd.categoryid "
  sSQL = sSQL & " WHERE dmd.dmid = " & sDMID

  set oGetCategoryInfoByDMID = Server.CreateObject("ADODB.Recordset")
  oGetCategoryInfoByDMID.Open sSQL, Application("DSN"), 3, 1

  if not oGetCategoryInfoByDMID.eof then
     sParentCategoryID   = oGetCategoryInfoByDMID("categoryid")
     sParentCategoryName = oGetCategoryInfoByDMID("categoryname")
  end if

  oGetCategoryInfoByDMID.close
  set oGetCategoryInfoByDMID = nothing

end sub

'------------------------------------------------------------------------------
sub getSectionInfo(ByVal iSectionMode, ByVal iSectionID, ByRef lcl_sectionname, ByRef lcl_sectiontype)

  lcl_sectionname        = ""
  lcl_sectiontype        = ""
  lcl_displaySectionName = 0

  if iSectionID <> "" then

    'Get the Section Info
     sSQL = "SELECT sectionname, "
     sSQL = sSQL & " sectiontype, "
     sSQL = sSQL & " displaySectionName "
     sSQL = sSQL & " FROM egov_dm_sections "
     sSQL = sSQL & " WHERE sectionid = " & iSectionID
     sSQL = sSQL & " AND isActive = 1 "

     set oGetSectionInfo = Server.CreateObject("ADODB.Recordset")
     oGetSectionInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetSectionInfo.eof then
        lcl_sectionname        = oGetSectionInfo("sectionname")
        lcl_sectiontype        = oGetSectionInfo("sectiontype")
        lcl_displaySectionName = oGetSectionInfo("displaySectionName")
     end if

     oGetSectionInfo.close
     set oGetSectionInfo = nothing
  end if

 'Determine if the section name is displayed.
 'If not then since this is the admin-site we want to still display it, but in a format that identifies it as "hidden"
 'We also do NOT want to show the name if we are display the section to the public
  if lcl_sectionname <> "" then
     if not lcl_displaySectionName then
        if iSectionMode = "PUBLIC_VIEW" then
           lcl_sectionname = ""
        else
           lcl_sectionname = "<em>" & lcl_sectionname & "</em>"
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
sub displayButtonsSection(iMode, iDMTypeID, iSectionID, iIsActiveByUser)

  response.write "<div id=""displayButtonsSection"" align=""center"" style=""margin-bottom: 5px;"">" & vbcrlf

  if iMode = "SECTION_VIEW" _
  OR iMode = "ACCOUNTINFO_VIEW" then
     response.write "  <input type=""button"" name=""maintainSectionButton"" id=""maintainSectionButton"" class=""button"" value=""Edit"" onclick=""editSection('" & iSectionID & "');"" />" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub dtb_debug(iValue)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(iValue,"'","''") & "') "

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub
%>
