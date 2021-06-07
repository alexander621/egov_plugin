<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 Dim iLetterID, sLetterBody, sLetterName
 iLetterID = request("FLid")

'Get form letter information
 subGetLetterInformation(iLetterID)
%>
<html>
<head>
  <title>E-Gov Administration Consule {Form Letter Preview}</title>

  <link rel="stylesheet" type="text/css" href="../global.css" />

</head>
<body>
<table cellpadding="10" border="0" width="550">
  <tr>
      <td valign="top">
          <p align="center"><strong><%=sLetterName%> (Template)</strong></p>
         	<p><%=sLetterBody%></p>
      </td>
  </tr>
</table>

<!--include file="bottom_include.asp"-->

</body>
</html>
<%
'------------------------------------------------------------------------------
sub subGetLetterInformation( intLetter )
  
  sSQL = "SELECT * FROM FormLetters WHERE FLid= " & intLetter

	 set oLetter = Server.CreateObject("ADODB.Recordset")
	 oLetter.Open sSQL, Application("DSN"), 3, 1
  
  if not oLetter.eof then
   		sLetterName = oLetter("FLtitle")
   		sLetterBody = oLetter("FLbody")
		   sLetterBody = replace( oLetter("FLbody"), "'", "''" )
   		'sLetterBody = replace( sLetterBody, chr(13), "<br>" ) 

    'Check to see if message contains HTML.
     if oLetter("containsHTML") then
        lcl_containsHTML = "Y"
     else
        lcl_containsHTML = "N"
     end if

    'Check for custom html.  If any HTML tags exist in the email body the admin is sending then leave our HTML off of the email.
     lcl_customhtml = 0

     if instr(UCASE(sLetterBody),"<HTML") > 0 then
        lcl_customhtml = lcl_customhtml + 1
     end if

     if instr(UCASE(sLetterBody),"<HEAD") > 0 then
        lcl_customhtml = lcl_customhtml + 1
     end if

     if instr(UCASE(sLetterBody),"<BODY") > 0 then
        lcl_customhtml = lcl_customhtml + 1
     end if

     if sLetterBody <> "" then
       'If the message IS in HTML format already there is no need to set the carriage returns "chr(10)" to "<br />" tags.
       'If the message is NOT in HTML format then we want to make sure that the carriage returns are picked up 
        if lcl_containsHTML = "Y" then
           lcl_replace_linebreaks = ""
        else
           lcl_replace_linebreaks = "<br />"
        end if

        sLetterBody = replace(sLetterBody,vbcrlf,lcl_replace_linebreaks & vbcrlf)
     end if

   		oLetter.Close
  end if

  set oLetter = nothing
	  
end sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  JSsafe = strDB
End Function
%>