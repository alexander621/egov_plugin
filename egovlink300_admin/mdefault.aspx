<%@ Page Language="C#" %>

<!DOCTYPE html>

<script runat="server">
	string OrgID = common.getOrgId( );
	
</script>

<html lang="en">
<head runat="server">
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<meta name="viewport" content="user-scalable=no, width=device-width" />
    <title>Community City</title>
    
    <link rel="stylesheet" href="styles/jquery.mobile-1.0a2.min.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.4.4.min.js"></script>
	<script type="text/javascript" src="https://code.jquery.com/mobile/1.0a2/jquery.mobile-1.0a2.min.js"></script>
	
</head>
<body>
 
    <div data-role="page" data-theme="a" id="jqm-home">
	<div data-role="header" data-theme="a">
		<h1>Community City (<%=OrgID %></h1>
	</div><!-- /header -->
    
    </div>
    
</body>
</html>
