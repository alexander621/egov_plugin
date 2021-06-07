<!-- #include file="loadfolder_home.inc" //-->


<%
' GET CITY DOCUMENT LOCATION
sLocationName =  GetVirtualDirectyName()
response.write "<!--" & sLocationName & "-->"

' GENERATE LIST OF DOCUMENTS BASED ON INSTITUTION
response.write LoadFolder("/public_documents300/" & sLocationName,"/public_documents300/custom/pub/" & sLocationName & "/published_documents"  )




'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = strURL(1) 

	GetVirtualDirectyName = trim(replace(sReturnValue,"/",""))

End Function%>


 
  
