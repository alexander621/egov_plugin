<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not UserHasPermission(session("userid"),"form creator") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

%>
<html>
<head>
 	<title>E-Gov Administration Console {Forms Management}</title>

	 <link type="text/css" rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	 <link type="text/css" rel="stylesheet" href="../global.css">

<style type="text/css">
.fieldset
{
   margin: 10px 0px;
   border-radius: 6px;
}

.fieldset legend
{
   padding: 4px 8px;
   border: 1pt solid #808080;
   border-radius: 6px;
   font-size: 1.25em;
   color: #800000;
}

#formsCreatorPageHeader
{
   font-size: 1.25em;
   font-weight: bold;
}

#screenMsg
{
   text-align: right;
   font-size: 1.125em;
   font-weight: bold;
   color: #ff0000;
}

#scFormName
{
   width: 83%;
}

#buttonSearch,
#buttonCreate
{
   cursor: pointer;
}

#buttonCreate
{
   margin: 10px 0px;
}

.formLabel
{
   white-space: nowrap;
}

#noFormsExist
{
   margin: 10px 0px;
   text-align: center;
   font-size: 1.25em;
   font-weight: bold;
   color: #ff0000;
}

#formsListTable
{
   width: 100%;
   background-color: #eeeeee;
}

#formsListTable th
{
   padding: 2px;
}

</style>

  <script src="../scripts/ajaxLib.js"></script>
  <script src="../scripts/modules.js"></script>
  <script src="../scripts/jquery-1.9.1.min.js"></script>

<script>
<!--
//-->
</script>
</head>
<body>
	<% ShowHeader sLevel %>
	<!-- #include file="../menu/menu.asp" //-->
<div id="content">
	<div id="centercontent">
		<div id="formsCreatorPageHeader">E-Gov Action Line Forms Market</div>
		<div id="screenMsg">&nbsp;</div>

		<%
 		'BEGIN: Forms List -----------------------------------------------------------
		%>
		<div>
			<input type="button" name="create" id="buttonCreate" onclick="window.location='copy_form.asp?task=NEW&iformid=57&iorgid=<%=session("orgid")%>';" value="Create a New Form" />

			<% subListForms %>

		</div>
		<%
 		'END: Forms List -------------------------------------------------------------
 		%>

	</div>
</div>
<!--#Include file="../admin_footer.asp"-->
</body>
</html>

<%
'------------------------------------------------------------------------------
sub subListForms()

  sSQL = "SELECT * "
  sSQL = sSQL & " FROM actionline_form_market "
  sSQL = sSQL & " ORDER BY formname "

 	set oFormList = Server.CreateObject("ADODB.Recordset")
 	oFormList.Open sSQL, Application("DSN"), 3, 1
	
 	if not oFormList.eof then %>
		<style> .tablelist button {margin-top:5px;} .tablelist button:first-child {margin-top:0} .tablelist .rowpad {padding-top:5px;padding-bottom:5px;}</style>
		<script>
		function openInNewTab(url) {
  			var win = window.open(url, '_blank');
  			win.focus();
		}
		</script>
	<%
    	response.write "<table id=""formsListTable"" border=""0"" cellspacing=""0"" class=""tablelist"">" & vbcrlf
	 	  response.write "  <tr>" & vbcrlf
     		  response.write "      <th align=""left"">Form Name</th>" & vbcrlf
		  response.write "	<th>Forms</th>" & vbcrlf
     		  response.write "  </tr>" & vbcrlf

     sRowClass = "formrowW"

   		do while not oFormList.eof
        sRowClass           = changeBGColor(sRowClass,"formrowW","formrowG")
     			sType               = "STANDARD"
        sEnabledColor       = "#ff0000"
        sEnabledLabel       = "OFF"
        sDisplayOnListColor = "#ff0000"

       'Display the row
   	  		response.write "  <tr class=""" & sRowClass & """>" & vbcrlf
   	  		response.write "      <td class=""formlist rowpad"">&nbsp;&nbsp;<nobr>" & oFormList("formname") & "</nobr></td>" & vbcrlf
   	  		response.write "      <td class=""formlist rowpad"" align=""center"">" 
						arrForms = split(oFormList("formids"),"|")
						x = 0
						for each id in arrForms
							x = x+1
							%>
							<button onclick="openInNewTab('formmarketview.asp?iformid=<%=id%>');">View Form<% if UBOUND(arrForms) > 0 then response.write " Version " & x %></button><br />
						<%next
			response.write "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

      		oFormList.movenext
     loop

    	response.write "</table>" & vbcrlf
 	else
     response.write "<div id=""noFormsExist"">*** No forms exist ***</div>" & vbcrlf
  end if

 	set oFormList = nothing

end sub

%>
