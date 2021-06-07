<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_picker.asp
' AUTHOR: ?????
' CREATED: ?????
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Stolen from list_picker.asp
'
' MODIFICATION HISTORY
' 1.0   5/1/2006   Steve Loar - Created this version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPostCount
iPostCount = 0

If request("postcount") <> "" Then
	iPostCount = clng(request("postcount")) + 1
End If 
%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->

<html>
	<head>
	  <title><%=langBSCommittees%></title>
	  <link rel="stylesheet" type="text/css" href="../global.css" />
	  <link rel="stylesheet" type="text/css" href="classes.css" />

<script language="Javascript">
<!--

	function openWin1(url, name) 
	{
		popupWin = window.open(url, name,"resizable,width=800,height=600");
	}

	function pause(numberMillis) {
        var now = new Date();
        var exitTime = now.getTime() + numberMillis;
        while (true) {
            now = new Date();
            if (now.getTime() > exitTime)
                return;
        }
    }

	function reloadparent( sURL, iPostCount )
	{
		if (iPostCount > 0)
		{
			opener.location.reload(true);
			pause(600);
			opener.location.href = sURL;
		}
	}

	function CloseYourself( sUrl )
	{
		opener.location.href=sUrl;
		self.close();
	}

//-->
</script>

	  
	</head>
	<body onUnload="javascript: opener.location.href='edit_class.asp?classid=<%=request("classid")%>#children');" onLoad="javasript:reloadparent('edit_class.asp?classid=<%=request("classid")%>#children',<%=iPostCount%>);" bgcolor="#c9def0">
		<form name-"ClassForm" method="post" action="">
		<input type="hidden" name="classid" value=<%=request("classid")%> />

		<h3>Add To: <%=GetClassName( request("classid") ) %></h3>
<% 
		ShowAvailableClasses 
%>
					&nbsp;<a href='javascript:document.c1.submit();'><img src='../images/ieforward.gif' align='absmiddle' border="0"></a>
					<input type="button" value="Close" onclick="CloseYourself('edit_class.asp?classid=<%=request("classid")%>#children');" />
		</form>
	</body>
</html>

<!--#Include file="class_global_functions.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub  ShowClassInstructorList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowAvailableClasses( )
	Dim sSql, oList

	sSql = "Select classid, classname from egov_class "
	sSql = sSql & " where classtypeid = 3 and orgid = " & Session("OrgID") & " Order By classname"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select border=""0"" name=""classid"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("classid") & """>" & oList("classname") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowClassWaiverList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowClassWaiverList( iClassid, sListtype )
	Dim sSql, oList

	sSql = "Select W.waiverid, W.waivername from egov_class_waivers W, egov_class_to_waivers C "
	sSql = sSql & " where C.waiverid = W.waiverid "
	sSql = sSql & " and C.classid = " & iClassId & " and W.orgid = " & Session("OrgID") & " Order By W.waivername"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	response.write vbcrlf & "<tr><td><strong>Assigned Waivers</strong></td></tr>"
	response.write vbcrlf & "<form name=""c1"" method=""post"" action=""class_removewaiver.asp"">"
	response.write vbcrlf & "<input type=""hidden"" name=""classid"" value=""" & iClassId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""listtype"" value=""" & sListtype & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""postcount"" value=""" & iPostCount & """ />"
	response.write vbcrlf & "<tr><td>"
	response.write vbcrlf & "<select class=""instructorlist"" size=""15"" border=""0"" name=""waiverlist"" multiple=""multiple"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("waiverid") & """>" & oList("waivername") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "</form>"
	response.write vbcrlf & "</table>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowClassCategoryList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowClassCategoryList( iClassid, sListtype )
	Dim sSql, oList

	sSql = "Select W.categoryid, W.categorytitle from egov_class_categories W, egov_class_category_to_class C "
	sSql = sSql & " where C.categoryid = W.categoryid "
	sSql = sSql & " and C.classid = " & iClassId & " and W.orgid = " & Session("OrgID") & " Order By W.categorytitle"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	response.write vbcrlf & "<tr><td><strong>Assigned Categories</strong></td></tr>"
	response.write vbcrlf & "<form name=""c1"" method=""post"" action=""class_removecategory.asp"">"
	response.write vbcrlf & "<input type=""hidden"" name=""classid"" value=""" & iClassId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""listtype"" value=""" & sListtype & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""postcount"" value=""" & iPostCount & """ />"
	response.write vbcrlf & "<tr><td>"
	response.write vbcrlf & "<select class=""instructorlist"" size=""15"" border=""0"" name=""categorylist"" multiple=""multiple"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("categoryid") & """>" & oList("categorytitle") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "</form>"
	response.write vbcrlf & "</table>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowRemainingInstructorList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowRemainingInstructorList( iClassid, sListtype )
	Dim sSql, oList

	sSql = "Select I.instructorid, I.firstname, I.lastname from egov_class_instructor I "
	sSql = sSql & " where I.instructorid not in ( select instructorid from egov_class_to_instructor where classid = " & iClassid & " )"
	sSql = sSql & " and I.orgid = " & Session("OrgID") & " Order By I.lastname, I.firstname"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	response.write vbcrlf & "<tr><td><strong>Available Instructors</strong></td></tr>"
	response.write vbcrlf & "<form name=""r1"" method=""post"" action=""class_addinstructor.asp"">"
	response.write vbcrlf & "<input type=""hidden"" name=""classid"" value=""" & iClassId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""listtype"" value=""" & sListtype & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""postcount"" value=""" & iPostCount & """ />"
	response.write vbcrlf & "<tr><td>"
	response.write vbcrlf & "<select class=""instructorlist"" size=""15"" border=""0"" name=""instructorlist"" multiple=""multiple"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("instructorid") & """>" & oList("lastname") & ", " & oList("firstname") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "</form>"
	response.write vbcrlf & "</table>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowRemainingInstructorList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowRemainingWaiverList( iClassid, sListtype )
	Dim sSql, oList

	sSql = "Select I.waiverid, I.waivername from egov_class_waivers I "
	sSql = sSql & " where I.waiverid not in ( select waiverid from egov_class_to_waivers where classid = " & iClassid & " )"
	sSql = sSql & " and I.orgid = " & Session("OrgID") & " Order By I.waivername"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	response.write vbcrlf & "<tr><td><strong>Available Waivers</strong></td></tr>"
	response.write vbcrlf & "<form name=""r1"" method=""post"" action=""class_addwaiver.asp"">"
	response.write vbcrlf & "<input type=""hidden"" name=""classid"" value=""" & iClassId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""listtype"" value=""" & sListtype & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""postcount"" value=""" & iPostCount & """ />"
	response.write vbcrlf & "<tr><td>"
	response.write vbcrlf & "<select class=""instructorlist"" size=""15"" border=""0"" name=""waiverlist"" multiple=""multiple"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("waiverid") & """>" & oList("waivername") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "</form>"
	response.write vbcrlf & "</table>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowRemainingCategoryList( iClassid, sListtype )
'--------------------------------------------------------------------------------------------------
Sub ShowRemainingCategoryList( iClassid, sListtype )
	Dim sSql, oList

	sSql = "Select I.categoryid, I.categorytitle from egov_class_categories I "
	sSql = sSql & " where I.categoryid not in ( select categoryid from egov_class_category_to_class where classid = " & iClassid & " )"
	sSql = sSql & " and I.orgid = " & Session("OrgID") & " and isroot = 0 Order By I.categorytitle"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	response.write vbcrlf & "<tr><td><strong>Available Categories</strong></td></tr>"
	response.write vbcrlf & "<form name=""r1"" method=""post"" action=""class_addcategory.asp"">"
	response.write vbcrlf & "<input type=""hidden"" name=""classid"" value=""" & iClassId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""listtype"" value=""" & sListtype & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""postcount"" value=""" & iPostCount & """ />"
	response.write vbcrlf & "<tr><td>"
	response.write vbcrlf & "<select class=""instructorlist"" size=""15"" border=""0"" name=""categorylist"" multiple=""multiple"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("categoryid") & """>" & oList("categorytitle") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "</form>"
	response.write vbcrlf & "</table>"

	oList.close
	Set oList = Nothing

End Sub 




%>

