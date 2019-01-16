<!DOCTYPE html>
<html>
<link rel="stylesheet" href="font-awesome/css/font-awesome.min.css">
	<style type="text/css">
	#Block
		{
			background: lightblue;
			width: 280px;
			height: auto;
			padding-left: 20px;
			color: white;
			position: fixed;
			left: 500px;
			top: 200px;
			padding-bottom: 10px;
			padding-top: 10px;
			border-top-left-radius: 45px;
			text-align: left;
			font-size: x-large;
		}
		#searchPicture
		{
			width: 80%;
			height: 60%;
		}

	</style>
<body>
	<div id="Block">
	<%SET o_cnn = Server.CreateObject("ADODB.Connection")
		  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		  SET rs=Server.CreateObject("ADODB.recordset")
		  rs.Open "SELECT Name, Vorname FROM AdressenAP WHERE Name='" & Request("name") & "'",o_cnn

		  do until rs.EOF
		  	Response.Write(rs.Fields("Name").Value & " " & rs.Fields("Vorname").value)

		  	rs.MoveNext
		  Loop
		  	o_cnn.close
		%>
	</div>
	<image id="searchPicture" src="https://image.flaticon.com/icons/svg/148/148834.svg"></image>
</body>
</html>
