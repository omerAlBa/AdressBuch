<!DOCTYPE html>
<html>
<head>
	<title>Page number</title>
</head>
<body>
	<script type="text/javascript">
		var Pointer;
	</script>
	<%
		Set v_Page = Server.CreateObject("ADODB.Connection")
		v_Page.Open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract" 

		sSql = "Select Name From AdressenAP"

		set adoRs = v_Page.Execute(sSql)
		
		'GetRows Retrieves multiple records of a Recordset object into an array.
		PageData = adoRs.GetRows()

		For i = 0 to ubound(PageData, 2)

		Next
		 Response.write(i)
		adoRs.Close
		v_Page.Close
		
		set adoRs = Nothing
		Set v_Page = Nothing
	%>
	<form method="get" action="../NeueHompage/index.asp">
		<input type="text" value="<%Response.write(i)%>" name="i">
	</form>
	
		 	
</body>
</html>
