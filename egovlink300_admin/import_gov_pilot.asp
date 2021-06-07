<%
response.buffer = false
response.write(ScriptEngine & "<br />")
response.write(ScriptEngineBuildVersion & "<br />")
response.write(ScriptEngineMajorVersion & "<br />")
response.write(ScriptEngineMinorVersion)
response.end

'sSQL = "SELECT TOP 1 Ref_No " _
sSQL = "SELECT Ref_No " _
      & " ,Date_Entered " _
      & " ,Property_Address " _
      & " ,Concern_Type " _
      & " ,Description_of_Concern " _
      & " ,Inspector_Assigned " _
      & " ,Complainant_Address as useraddress " _
      & " ,Complainant_Email as useremail " _
      & " ,Complainant_First_Name as userfname " _
      & " ,Complainant_Last_Name as userlname " _
      & " ,Complainant_Phone as userphone " _
      & " ,Complainant_Zip as userzip " _
      & " ,Concern_Type_Other " _
      & " ,Date_Closed " _
      & " ,Date_Escalation " _
      & " ,Date_Verified " _
      & " ,Defendant_Address " _
      & " ,Defendant_CSZ " _
      & " ,Defendant_Name " _
      & " ,Department_Assigned " _
      & " ,Source_Type " _
      & " ,Status " _
      & " ,u.UserID as adminid " _
  & " FROM harlingen_govpilot_import " _
  & " LEFT JOIN Users u ON u.OrgID = 209 AND ( u.FirstName + ' ' + u.LastName = inspector_assigned or u.FirstName + ' J. ' + u.LastName = inspector_assigned)"
  'sSQL = sSQL & "WHERE ref_no = '2018-01-2919'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1

Do WHile Not oRs.EOF
 	sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid = 0"


	'Insert User Info (userfname, userlname, useremail, userhomephone, useraddress, userzip)
	sSQL = "INSERT INTO egov_users (userfname, userlname, useremail, userhomephone, useraddress, userzip) VALUES('" & oRs("userfname") & "','" & oRs("userlname") & "','" & oRs("useremail") & "','" & oRs("userphone") & "','" & oRs("useraddress") & "','" & oRs("userzip") & "')"
    	Set oCmd = Server.CreateObject("ADODB.Connection")
    	oCmd.Open Application("DSN")
    	oCmd.Execute(sSQL)
    	
	'Get Userid FROM insert
    	sSQL2 = "SELECT @@IDENTITY AS NewID"
    	set rs = oCmd.Execute(sSQL2)
    	userid = rs.Fields("NewID").value
    	set rs = Nothing
    	oCmd.Close
    	Set oCmd = Nothing

	adminid = oRs("adminid")
	if isnull(oRs("adminid")) then adminid = "NULL"

	sSQL = "INSERT INTO egov_actionline_requests (orgid,category_id,category_title,userid, submit_date,status,assignedemployeeid,comment) VALUES(209,18675,'GovPilot Import'," & userid & ",'" & oRs("Date_Entered") & "','RESOLVED'," & adminid & ",'"
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Ref No.</b><br>" & vbcrlf & oRs("ref_no") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Property Address</b><br>" & vbcrlf & oRs("Property_Address") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Concern Type</b><br>" & vbcrlf & oRs("Concern_Type") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Description of Concern</b><br>" & vbcrlf & oRs("Description_of_Concern") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Concern Type Other</b><br>" & vbcrlf & oRs("Concern_Type_Other") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Date Closed</b><br>" & vbcrlf & oRs("Date_Closed") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Date Escalation</b><br>" & vbcrlf & oRs("Date_Escalation") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Date Verified</b><br>" & vbcrlf & oRs("Date_Verified") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Defendant Address</b><br>" & vbcrlf & oRs("Defendant_Address") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Defendant CSZ</b><br>" & vbcrlf & oRs("Defendant_CSZ") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Defendant Name</b><br>" & vbcrlf & oRs("Defendant_Name") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Department Assigned</b><br>" & vbcrlf & oRs("Department_Assigned") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Source Type</b><br>" & vbcrlf & oRs("Source_Type") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "<p>" & vbcrlf & "<b>Status</b><br>" & vbcrlf & oRs("Status") & vbcrlf & "</p>" & vbcrlf
	sSQL = sSQL & "')"
	'response.write sSQL & "<br />" & vbcrlf
    	Set oCmd = Server.CreateObject("ADODB.Connection")
    	oCmd.Open Application("DSN")
    	oCmd.Execute(sSQL)
    	oCmd.Close
    	Set oCmd = Nothing

	oRs.MoveNext
loop

%>
