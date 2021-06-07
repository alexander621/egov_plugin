<%
'------------------------------------------------------------------------------
function getCustomReportID(iReportType, iOrgID, iUserID, iIsDefault)
  lcl_return         = ""
  lcl_customreportid = 0

  sSQL = "SELECT cr.customreportid, cr.reporttypeid, cr.reportname, cr.isuserdefault "
  sSQL = sSQL & " FROM egov_customreports cr, egov_customreports_reporttypes rt "
  sSQL = sSQL & " WHERE cr.reporttypeid = rt.reporttypeid "

 'Search for Default Fields
 'A "FALSE" is passed into THESE calls because we want to skip THIS set and just use the query.
  if iIsDefault then
    'Check for a default record for the org.
     lcl_customreportid = getCustomReportID("ACTIONLINE - DEFAULTS", session("orgid"), 0, False)

    'If one does NOT exist specifically FOR the org then grab the system default record.
    'The system default record will have: userid = 0, orgid = 0
     if lcl_customreportid = "" then
        lcl_customreportid = getCustomReportID("ACTIONLINE - DEFAULTS", 0,0, False)
     end if

     sSQL = sSQL & " AND cr.customreportid = " & lcl_customreportid

  else
     sSQL = sSQL & " AND UPPER(rt.reporttype) = '" & UCASE(iReportType) & "'"
     sSQL = sSQL & " AND cr.orgid = "  & iOrgID
     sSQL = sSQL & " AND cr.userid = " & iUserID
  end if

 	set oCustomRpt = Server.CreateObject("ADODB.Recordset")
	 oCustomRpt.Open sSQL, Application("DSN"), 3, 1

  if not oCustomRpt.eof then
     lcl_return = oCustomRpt("customreportid")
  else
     if UCASE(iReportType) <> "ACTIONLINE - DEFAULTS" then
       'Set up customreport for non-ORG/SYSTEM DEFAULTS
        getReportTypeInfo iReportType, lcl_reporttypeid, lcl_reportname

        sSQL = "INSERT INTO egov_customreports ("
        sSQL = sSQL & "orgid, "
        sSQL = sSQL & "userid, "
        sSQL = sSQL & "reporttypeid, "
        sSQL = sSQL & "reportname, "
        sSQL = sSQL & "isuserdefault"
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL &       session("orgid")  & ", "
        sSQL = sSQL &       session("userid") & ", "
        sSQL = sSQL &       lcl_reporttypeid  & ", "
        sSQL = sSQL & "'" & lcl_reportname    & "', "
        sSQL = sSQL & "0"
        sSQL = sSQL & ")"

       	set oInsertCR = Server.CreateObject("ADODB.Recordset")
      	 oInsertCR.Open sSQL, Application("DSN"), 3, 1

       'Retrieve the posting_id that was just inserted
        sSQLid = "SELECT IDENT_CURRENT('egov_customreports') as NewID"
        oInsertCR.Open sSQLid, Application("DSN"), 3, 1
        lcl_identity = oInsertCR.Fields("NewID").value

        lcl_return = lcl_identity

        set oInsertCR = nothing
     end if
  end if

  oCustomRpt.close
  set oCustomRpt = nothing

  getCustomReportID = lcl_return

end function

'------------------------------------------------------------------------------
sub getCustomReportInfo(ByVal iReportType, ByVal iIsDefault, ByVal iOrgID, ByVal iUserID, ByVal iIsUserDefault, _
                        ByRef lcl_customreportid, ByRef lcl_reportypeid, ByRef lcl_reportname, ByRef lcl_isuserdefault)
  lcl_customreportid = ""
  lcl_reporttypeid   = 0
  lcl_reportname     = ""
  lcl_isuserdefault  = False

  sSQL = "SELECT cr.customreportid, cr.reporttypeid, cr.reportname, cr.isuserdefault "
  sSQL = sSQL & " FROM egov_customreports cr, egov_customreports_reporttypes rt "
  sSQL = sSQL & " WHERE cr.reporttypeid = rt.reporttypeid "

 'Search for Default Fields
 'A "FALSE" is passed into THESE calls because we want to skip THIS set and just use the query.
  if iIsDefault then
    'Check for a default record for the org.
     getCustomReportInfo "ACTIONLINE - DEFAULTS", False, session("orgid"), 0, False, lcl_customreportid, lcl_reporttypeid, _
                         lcl_reportname, lcl_isuserdefault

    'If one does NOT exist specifically FOR the org then grab the system default record.
    'The system default record will have: userid = 0, orgid = 0
     if lcl_customreportid = "" then
        getCustomReportInfo "ACTIONLINE - DEFAULTS", False, 0,0, False, lcl_customreportid, lcl_reporttypeid, _
                            lcl_reportname, lcl_isuserdefault
     end if

     sSQL = sSQL & " AND cr.customreportid = " & lcl_customreportid

  else
     sSQL = sSQL & " AND UPPER(rt.reporttype) = '" & UCASE(iReportType) & "'"
     sSQL = sSQL & " AND cr.orgid = "  & iOrgID
     sSQL = sSQL & " AND cr.userid = " & iUserID

    'Check if the user wants to find his/her default search options.
    '*** Currently users can only have a SINGLE "save search options" record.
    '*** In the future users will be able to save multiple search option records, and at that time...
    '*** They will be able to choose which one is their default search options.
     if iIsUserDefault then
        sSQL = sSQL & " AND cr.isuserdefault = 1 "
     end if

  end if

 	set oCustomRpt = Server.CreateObject("ADODB.Recordset")
	 oCustomRpt.Open sSQL, Application("DSN"), 3, 1

  if not oCustomRpt.eof then
     lcl_customreportid = oCustomRpt("customreportid")
     lcl_reporttypeid   = oCustomRpt("reporttypeid")
     lcl_reportname     = oCustomRpt("reportname")
     lcl_isuserdefault  = oCustomRpt("isuserdefault")
  end if

  oCustomRpt.close
  set oCustomRpt = nothing

end sub

'------------------------------------------------------------------------------
function getCustomReportSearchOption(iCustomReportID, iDBColumnName)
  lcl_return = ""

  if iCustomReportID <> "" AND iDBColumnName <> "" then
     sSQL = "SELECT rc.searchvalue "
     sSQL = sSQL & " FROM egov_customreports_reportcolumns rc, egov_customreports_dbcolumns dbc "
     sSQL = sSQL & " WHERE rc.dbcolumnid = dbc.dbcolumnid "
     sSQL = sSQL & " AND rc.customreportid = " & iCustomReportID
     sSQL = sSQL & " AND UPPER(dbc.dbcolumnname) = '" & UCASE(iDBColumnName) & "'"

    	set oSearchOption = Server.CreateObject("ADODB.Recordset")
   	 oSearchOption.Open sSQL, Application("DSN"), 3, 1

     if not oSearchOption.eof then
        lcl_return = oSearchOption("searchvalue")
     end if

     oSearchOption.close
     set oSearchOption = nothing

  end if

  getCustomReportSearchOption = lcl_return

end function

'------------------------------------------------------------------------------
sub saveCustomReportSearchOption(ByVal iCustomReportID, ByVal iDBColumnName, ByVal iSearchValue, ByVal iIsAjaxRoutine, ByRef lcl_success)

  lcl_error_msg = "Required Fields Missing: CustomReportID [" & iCustomReportID & "] - ReportColumn [" & iDBColumnName & "]"

  if iCustomReportID <> "" AND iDBColumnName <> "" AND lcl_success = "Y" then

    'Validate the searchvalue
     if iSearchValue <> "" then
        lcl_searchvalue = formatReportColumn(iCustomReportID, iDBColumnName, iSearchValue, lcl_success)
     else
        lcl_searchvalue = "NULL"
     end if

    'If the value is "Invalid" then send back the error message.  Otherwise, continue with update/insert.
     if lcl_searchvalue = "INVALID VALUE" then
        lcl_success = "N"

        if iIsAjaxRoutine then
           response.write "Invalid Value: " & iDBColumnName
        end if
     elseif lcl_searchvalue = "error" then
        lcl_success = "N"

        if iIsAjaxRoutine then
           response.write lcl_error_msg
        end if
     elseif lcl_searchvalue = "DBCOLUMN DOES NOT EXIST" then
        lcl_success = "N"

        if iIsAjaxRoutine then
           response.write "DBColumnName [" & iDBColumnName & "] does not exist on egov_customreports_dbcolumns for CustomReportID [" & iCustomReportID & "]"
        end if
     else
        if lcl_searchvalue = "DBCOLUMN DOES NOT EXIST" then
           lcl_success = "N"

           if iIsAjaxRoutine then
              response.write "DBColumnName [" & iDBColumnName & "] does not exist on egov_customreports_dbcolumns for CustomReportID [" & iCustomReportID & "]"
           end if
        else
          'Check to see if the "DBColumnName" is a valid column for this report.
           checkForReportColumnOnCustomReport iCustomReportID, iDBColumnName, lcl_dbcolumnid, lcl_dbcolumndatatype, _
                                              lcl_dbcolumndatalength, lcl_return

           if lcl_return = "INSERT" then
              sSQL = "INSERT INTO egov_customreports_reportcolumns (customreportid, dbcolumnid, searchvalue) VALUES ("
              sSQL = sSQL & iCustomReportID & ", "
              sSQL = sSQL & lcl_dbcolumnid  & ", "
              sSQL = sSQL & lcl_searchvalue & ") "
           else
              sSQL = "UPDATE egov_customreports_reportcolumns SET "
              sSQL = sSQL & " searchvalue = " & lcl_searchvalue
              sSQL = sSQL & " WHERE dbcolumnid = (select dbcolumnid "
              sSQL = ssQL &                     " from egov_customreports_dbcolumns "
              sSQL = ssQL &                     " where UPPER(dbcolumnname) = '" & UCASE(iDBColumnName) & "') "
              sSQL = sSQL & " AND customreportid = " & iCustomReportID
           end if

          	set oSaveOpt = Server.CreateObject("ADODB.Recordset")
         	 oSaveOpt.Open sSQL, Application("DSN"), 3, 1

           set oSaveOpt = nothing
           lcl_success = "Y"
        end if
     end if
  else
     if iCustomReportID <> "" AND iDBColumnName <> "" then
        lcl_success = "N"

        if iIsAjaxRoutine then
           response.write lcl_error_msg
		else
			response.write "<!--" & lcl_error_msg & "-->"
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
function formatReportColumn(iCustomReportID, iDBColumnName, iSearchValue, lcl_success)

  lcl_return = "NULL"

  if iCustomReportID <> "" AND iDBColumnName <> "" AND lcl_success = "Y" then

    'Check to see if the "DBColumnName" is a valid column for this report.
     checkForReportColumnOnCustomReport iCustomReportID, iDBColumnName, lcl_dbcolumnid, lcl_dbcolumndatatype, _
                                        lcl_dbcolumndatalength, lcl_return

    'Format the SearchValue
    	'response.write iDBColumnName & " = " & iSearchValue & " - "
     if iSearchValue <> "NULL" AND iSearchValue <> "" AND (lcl_return = "NULL" OR lcl_return = "INSERT" OR lcl_return = "UPDATE") then
	     'response.write lcl_dbcolumndatatype & " - " & lcl_dbcolumndatalength

       'Varchar
        if lcl_dbcolumndatatype = "VARCHAR" then
           if len(iSearchValue) <= lcl_dbcolumndatalength then
              lcl_return = "'" & dbsafe(iSearchValue) & "'"
	   else
	      lcl_return = "'" & left(dbsafe(iSearchValue),lcl_dbcolumndatalength) & "'"
	      'lcl_return = "NULL"
           end if

       'Int
        elseif lcl_dbcolumndatatype = "INT" then
           iSearchValue = replace(iSearchValue,"all",0)

           if dbready_number(iSearchValue) then
              lcl_return = iSearchValue
           else
              lcl_return = "INVALID VALUE"
           end if

       'Date ONLY
        elseif lcl_dbcolumndatatype = "DATE" then
           if len(iSearchValue) <= lcl_dbcolumndatalength then
              if dbready_date(iSearchValue) then
                 lcl_return = "'" & dbsafe(iSearchValue) & "'"
              else
                 lcl_return = "INVALID VALUE"
              end if
           end if

        else
           lcl_return = "INVALID VALUE"
        end if
     end if
  else
     if iCustomReportID <> "" AND iDBColumnName <> "" then
        lcl_return = "error"
     end if
  end if
	'response.write " - " & lcl_return & "<br />"
     'response.flush
  formatReportColumn = lcl_return

end function

'------------------------------------------------------------------------------
sub checkForReportColumnOnCustomReport(ByVal iCustomReportID, ByVal iDBColumnName, ByRef lcl_dbcolumnid, _
                                       ByRef lcl_dbcolumndatatype, ByRef lcl_dbcolumndatalength, ByRef lcl_return)

  if iCustomReportID <> "" AND iDBColumnName <> "" then
    'Get the datatype and length of the ReportColumn
     sSQL = "SELECT dbc.dbcolumnid, dbc.dbcolumndatatype, dbc.dbcolumndatalength "
     sSQL = sSQL & " FROM egov_customreports_dbcolumns dbc, egov_customreports_reportcolumns rc "
     sSQL = sSQL & " WHERE dbc.dbcolumnid = rc.dbcolumnid "
     sSQL = sSQL & " AND UPPER(dbc.dbcolumnname) = '" & iDBColumnName & "' "
     sSQL = sSQL & " AND rc.customreportid = " & iCustomReportID

    	set oColumnSpecs = Server.CreateObject("ADODB.Recordset")
   	 oColumnSpecs.Open sSQL, Application("DSN"), 3, 1

     if not oColumnSpecs.eof then
        lcl_return = "UPDATE"
        lcl_dbcolumnid         = oColumnSpecs("dbcolumnid")
        lcl_dbcolumndatatype   = UCASE(oColumnSpecs("dbcolumndatatype"))
        lcl_dbcolumndatalength = oColumnSpecs("dbcolumndatalength")
     else
       'Check to see if the DBColumnName exists on egov_customreports_dbcolumns.
       'If "yes" then insert the column into egov_customreports_reportcolumns.
       'If "no" then show the error message.
        getDBColumnInfo iDBColumnName, lcl_dbcolumnid, lcl_dbcolumndatatype, lcl_dbcolumndatalength, lcl_return

        lcl_dbcolumnid         = lcl_dbcolumnid
        lcl_dbcolumndatatype   = lcl_dbcolumndatatype
        lcl_dbcolumndatalength = lcl_dbcolumndatalength
        lcl_return             = lcl_return
     end if

     oColumnSpecs.close
     set oColumnSpecs = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub getDBColumnInfo(ByVal iDBColumnName, ByRef lcl_dbcolumnid, ByRef lcl_dbcolumndatatype, ByRef lcl_dbcolumndatalength, ByRef lcl_return)

  sSQL = "SELECT dbcolumnid, dbcolumndatatype, dbcolumndatalength "
  sSQL = sSQL & " FROM egov_customreports_dbcolumns "
  sSQL = sSQL & " WHERE UPPER(dbcolumnname) = '" & UCASE(iDBColumnName) & "'"

  set oDBColumnInfo = Server.CreateObject("ADODB.Recordset")
  oDBColumnInfo.Open sSQL, Application("DSN"), 3, 1

  if not oDBColumnInfo.eof then
     lcl_return             = "INSERT"
     lcl_dbcolumnid         = oDBColumnInfo("dbcolumnid")
     lcl_dbcolumndatatype   = UCASE(oDBColumnInfo("dbcolumndatatype"))
     lcl_dbcolumndatalength = oDBColumnInfo("dbcolumndatalength")
  else
     lcl_return = "DBCOLUMN DOES NOT EXIST"
     lcl_dbcolumnid         = 0
     lcl_dbcolumndatatype   = ""
     lcl_dbcolumndatalength = ""
  end if

  'oDBColumnInfo.close
  set oDBColumnInfo = nothing

end sub

'------------------------------------------------------------------------------
sub getReportTypeInfo(ByVal iReportType, ByRef lcl_reporttypeid, ByRef lcl_reportname)
  lcl_reporttypeid = 0
  lcl_reportname   = ""

  if iReportType <> "" then
     sSQL = "SELECT reporttypeid, reportname "
     sSQL = sSQL & " FROM egov_customreports_reporttypes "
     sSQL = sSQL & " WHERE UPPER(reporttype) = '" & UCASE(iReportType) & "'"

     set oRptType = Server.CreateObject("ADODB.Recordset")
     oRptType.Open sSQL, Application("DSN"), 3, 1

     if not oRptType.eof then
        lcl_reporttypeid = oRptType("reporttypeid")
        lcl_reportname   = oRptType("reportname")
     end if

     oRptType.close
     set oRptType = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub updateCustomReport(ByVal iCustomReportID, ByVal iReportType, ByVal iReportName, ByVal iIsUserDefault)

  if iCustomReportID <> "" AND iReportType <> "" then

    'Get the ReportType data in case we need the "reportname"
     getReportTypeInfo iReportType, lcl_reporttypeid, lcl_reportname

     if iReportName <> "" then
        lcl_reportname = "'" & dbsafe(iReportName)    & "'"
     else
        lcl_reportname = "'" & dbsafe(lcl_reportname) & "'"
     end if

     if UCASE(iIsUserDefault) = "ON" then
        lcl_isUserDefault = 1
     else
        lcl_isUserDefault = 0
     end if

     sSQL = "UPDATE egov_customreports SET "
     sSQL = sSQL & " reportname = "    & lcl_reportname    & ", "
     sSQL = sSQL & " isuserdefault = " & lcl_isUserDefault
     sSQL = sSQL & " WHERE customreportid = " & iCustomReportID

     set oCRUpdate = Server.CreateObject("ADODB.Recordset")
     oCRUpdate.Open sSQL, Application("DSN"), 3, 1

     set oCRUpdate = nothing
  end if

end sub
%>
