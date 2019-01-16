<!DOCTYPE html>
<html>
<body>

	<%
		Set o_cnn = Server.CreateObject("ADODB.Connection")
		o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		sql = "INSERT INTO AdressenAP (Titel, Name, Vorname) " &_
			  "VALUES ('" & Request("Gname") & "','" & Request("Nname") & "','" & Request("Vname") & "')"

		Response.Write sql

		o_cnn.Execute sql
		o_cnn.close

		Response.Redirect("../index.asp")
	%>

</body>
</html>