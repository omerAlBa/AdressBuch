<!DOCTYPE html>
<html>
<link rel="stylesheet" href="font-awesome/css/font-awesome.min.css">
<style type="text/css">
	#infoPicture
	{
		width: 50px;
		height: 50px;
		position: fixed;
		left: 800px;
		top: 150px;
	}
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
		#header
		{
			position: fixed;
			left:520px;
			top: 175px;
			color: blue;
			font-size: x-large;
		}
		.AddUserPic
		{
			width: 30px;
			height: 30px;
			margin-left: 8.5px;
		}
		body
		{
			background-image: url("https://image.flaticon.com/icons/svg/1343/1343912.svg")
			
		}

</style>
<div id=header>Info</div>
<body><div id="Block">
	<%
	SET o_cnn = Server.CreateObject("ADODB.Connection")
		o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		SET rs=Server.CreateObject("ADODB.recordset")
		rs.Open "SELECT Geschlecht, Name, Vorname, APID FROM AdressenAP WHERE APID='" & Request.Querystring("APID") & "'",o_cnn
		
		if Request("Geschlecht")=0 then
			Response.Write("Herr" & " ")
		end if 
		if Request("Geschlecht")=255 then
			Response.Write("Frau" & " ")
		end if


		Response.Write(rs.Fields("Name").value & " " & " " & rs.Fields("Vorname"))

		Response.Write("<a class=""verlinkung""href=""../Update/Changer.asp?APID=" & rs.Fields("APID") & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/126/126794.svg""></a>")


	%></div>
</body>
</html>
<image id="infoPicture" src="https://image.flaticon.com/icons/svg/1076/1076337.svg"></image>




