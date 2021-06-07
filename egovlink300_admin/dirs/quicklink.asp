<br />
<% Call DrawQuicklinks("",1) %>

<form method="post" name="SearchUser"  action="search_action.asp">
  <div style="padding-bottom:3px;"><%=langSearch & " " %> Departments:</div>
  <table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td><input type="text" size="18" name="keywords" id="keywords" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" /></td>
    </tr>
    <tr align="right">
        <td>
            <a href="" onclick="return CheckKeyWords();"><img src="<%=RootPath%>images/go.gif" border="0" /><%=langGo%></a>
        </td>
    </tr>
 </table>
</form>

<script language="javascript">
<!--
	function CheckKeyWords()
	{
		if (RTrim(document.SearchUser.keywords.value) == "")
		{
			alert("No Keywords");
			document.SearchUser.keywords.focus();
		 return false;				
		}					
		return true;
	}

	function isWhitespace (s)
	{   var i;
		for (i = 0; i < s.length; i++)
		{   
			// Check that current character isn't whitespace.
			var c = s.charAt(i);

			if (whitespace.indexOf(c) == -1) return false;
		}
		return true;
	}

	function RTrim(str)
	/***
			PURPOSE: Remove trailing blanks from our string.
			IN: str - the string we want to RTrim

			RETVAL: An RTrimmed string!
	***/
	{
			// We don't want to trip JUST spaces, but also tabs,
			// line feeds, etc.  Add anything else you want to
			// "trim" here in Whitespace
			var whitespace = new String(" \t\n\r");

			var s = new String(str);

			if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
				// We have a string with trailing blank(s)...

				var i = s.length - 1;       // Get length of string

				// Iterate from the far right of string until we
				// don't have any more whitespace...
				while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
					i--;


				// Get the substring from the front of the string to
				// where the last non-whitespace character is...
				s = s.substring(0, i+1);
			}

			return s;
	}
//-->
</script>