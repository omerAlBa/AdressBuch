<!DOCTYPE html>
<html>
<head>

	<link rel="stylesheet" href="Design.css">
	<title>AdressBuch</title>

</head>

<body>
	<h1>AdressBuch</h1>
	<div id="Block">
		
		<%	
			Set v_Page = Server.CreateObject("ADODB.Connection")
			v_Page.Open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract" 

			sSql = "Select Name, Vorname, Geschlecht, APID From AdressenAP"

			set adoRs = v_Page.Execute(sSql)

			'GetRows Retrieves multiple records of a Recordset object into an array.
			PageData = adoRs.GetRows()
			
			For v = 0 to ubound(PageData, 2)
			Next

			anzahl = 10

			seite = Request("Pointer1")
			start = (seite * anzahl) - anzahl
			start = seite * anzahl
			ende  = start + anzahl -1 
			
			if ende>v then
				ende=v-1
			end if

			For i = start to ende

				if PageData(2,i)=0 then
					Response.Write "Herr" & " "
				end if
				if PageData(2,i)=255 then
					Response.Write "Frau" & " "
				end if

				Response.Write  PageData(0,i) & " " & PageData (1,i)
				'Add Links'
				Response.Write("<a class=""verlinkung"" onclick=""myFunction();"" href=""DeleteFile/DeleteUser.asp?APID=" & PageData(3,i) & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/126/126831.svg""></a>")
				Response.Write("<a class=""verlinkung""href=""InfoFile/info.asp?APID=" & PageData(3,i) & """><image  class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/1076/1076337.svg""></a>") 
				Response.Write("<a class=""verlinkung""href=""Update/Changer.asp?APID=" & PageData(3,i) & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/126/126794.svg""></a>")
			  	Response.Write("<br>")
				  
				Next

				adoRs.Close
				v_Page.Close
			%>	
		</div>
			<%
				set adoRs = Nothing
				Set v_Page = Nothing
			%>
		<!------------------Seiten Zahleingabe------------------------->
	<div id="untereKasten">
		<div id="SeitenZahlEingabe" class="AnzeigeDerSeite">


			<a class="recht" href="index.asp?Pointer1=<%Response.Write(Request("Pointer1"))-1%>"><image class="recht" src="https://image.flaticon.com/icons/svg/1286/1286897.svg"></a>
			<a id="recht1" href="index.asp?Pointer1=0"><image class="recht1" src="https://image.flaticon.com/icons/svg/1427/1427080.svg"></image></a>

			<script src="https://code.jquery.com/jquery-latest.js"></script>
			<form method="get" action="index.asp">
				<input min="0" max="<%Response.Write V%>" type="number" value="<%Response.Write(Request("Pointer1")+1)%>" name="Pointer1">
				<script type="text/javascript">
					if ( $("#SeitenAusgeber").val<0) {
						$("#SeitenAusgeber").val=0
					}
				</script>
			</form>

			<a class="links" href="index.asp?Pointer1=<%Response.Write(Request("Pointer1"))+1%>"><image class="links" src="https://image.flaticon.com/icons/svg/1286/1286867.svg"></a>
			<a class="links1" href="index.asp?Pointer1=<%Response.Write(round(v/10))%>"><image class="links1" src="https://image.flaticon.com/icons/svg/1427/1427051.svg">
	
	</div></div>
	<!-----------------------------Close Paging/Homepage-------------------------------------------->
	<div id="neuerEintrag">
		<a href="AddUserFile/index.html" id="ADDBTN"><image src="https://image.flaticon.com/icons/svg/104/104779.svg"></image></a>
		<div id="neuerEintrag1">Neuer Eintrag?</div>
	</div>

	<div id="SucheDIV">
		<form method="get" action="SearchFile/ISearch.asp">		
			<input id="Sucher" type="text" placeholder="Name..." name="name">
		</form>
	</div>
		
	<div class="zeigAn"> <!--Show Number of Pages-->
		<%Response.Write ("Seite" & " " & (Request("Pointer1")+1) & " " & "von" & " " & round(v/10)+1)%>
	</div>
	<!---Function------>
	<script>
		if (<%Response.Write(start)%>==0){
			$(".recht").hide();
			$(".recht1").hide();	 
		}
		if (<%Response.Write(Request("Pointer1"))%>==(<%Response.Write(round(v/10))%>)){
			$(".links").hide();
			$(".links1").hide();
		}

		function myFunction() {
		var ConfirmEntscheid = confirm("Are you sure?");
			if (ConfirmEntscheid == false)
			{
				event.preventDefault();
				alert("Execution stoped");
			}
			if (ConfirmEntscheid == true) 
			{
				alert("Delete was successful")
			}
		}
	</script>
	<!-----------XML---------->
	<a id="Benzema" href="downloadXML.asp">XML</a><br>
	<a href="downloadJSON.asp">JSON</a>
	
	<!---<form action="Update/Upload2sql.asp" method="post" enctype="multi/form-data">
		<label>W&aumlhlen Sie die hochzuladene Datein von Ihrem aus:
		<input name="datei" type="file" multiple>
		<input type="submit" name="text/xml">
		</label>
	</form>--->
	<form method="post" enctype="multipart/form-data" action="Update/Upload2sql.asp" id="excelUpload">
	<input type=file size=50 name="FILE1" id="FILE1"> <br/> 
	<input type="submit" name="text/xml">
	</form>
</body>
</html>