<!DOCTYPE html>
<html>
<body>
<%
	Set o_cnn = Server.CreateObject("ADODB.Connection")
	o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"%>

<%
	set rs=Server.CreateObject("ADODB.recordset")
	rs.Open "DELETE FROM AdressenAP WHERE APID='" & Request.Querystring("APID") & "'", o_cnn
	
	o_cnn.close	
	Response.Redirect("../index.asp")
%>
</body>
</html>