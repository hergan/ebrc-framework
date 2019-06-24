<!-- #INCLUDE FILE="assets/inc/adovbs.asp" -->
<!-- #INCLUDE FILE="assets/inc/defaults.asp" -->
<!-- #INCLUDE FILE="assets/inc/common.asp" -->



<%
' Vars Required for Form - Change Values here to change the form.

Dim	GTM
	'TODO: Change GTM ID
	GTM = "GTM-"

Dim sURI, sSeminar_Date, rsSeminars
Dim sSQL, sOutput, sSeminar_ID, sDateTime, sCity, sLocation, sAddress, sCityStateZip, sPhone, sIdValue

Dim x, sValue


Dim cust_id, firstname, middlename, lastname
Dim address, city, state, zipcode
Dim phone, email
Dim dob_mm, dob_dd, dob_yyyy, gender, primarycare, seminar
Dim k1_dob_mm, k1_dob_dd, k1_dob_yyyy, k1_gender
Dim k2_dob_mm, k2_dob_dd, k2_dob_yyyy, k2_gender
Dim k3_dob_mm, k3_dob_dd, k3_dob_yyyy, k3_gender
Dim k4_dob_mm, k4_dob_dd, k4_dob_yyyy, k4_gender

dim daysInMonth, k1_daysInMonth, k2_daysInMonth, k3_daysInMonth, k4_daysInMonth

Dim healthwellness, sportsperf, kidshlthenews, balance

dim intFields, arrFields(500, 2), sItem
Dim itemsRequested(99)


'hidden fields
dim sClientCode, sMarketCode, sProcCode, sformcode, sebrc_track_id, suserid, susertype, slookupfile
dim sPageTitle, sClientValidation_JavaCode, sBaseFont, sLogoImg, sBkImage, sBkColor

dim strMissing, strErrorMsg

	sClientCode = "0846"
	sMarketCode = "GB"
	sProcCode = "NRP"
	sFormCode = "190501A"
	sebrc_track_id = session("ebrc_id")
	suserid = "eBRC"
	susertype = "O"
	slookupfile = ""
	sPageTitle = "St. Joseph Medical Center"
	sClientValidation_JavaCode = ""
	sBaseFont = "Abadi MT Condensed"
	sLogoImg = "images/blank.gif"
	sBkImage = ""
	sBkColor = "#FFFFFF"
	
	fVerify = false
	
	' Test DB and report down if can not connect.
	DBTest sConn, sDBLogin, sDBPassword, "assets/inc/DBDown.asp"

	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open sConn, sDBLogin, sDBPassword
	if oconn is nothing then
		strResponse = strResponse & "No connection made"
	else
	end if
	
	'pull the items available info from the database and NOT hard code it!!!!
	sSQL = "select * FROM [Prod_System].[dbo].[form_definition] "
	sSQL = sSQL & " where client_num = '" & sClientcode & "' "
	sSQL = sSQL & " and formcode = '" & sformcode & "' "
	sSQL = sSQL & " and fieldname like ('item_%') "
	sSQL = sSQL & " and fieldtype = 'C'  "
	sSQL = sSQL & " order by row_no, col_no "
	'response.write sSQL & "<br />"
	'response.end

	Set rsItems = Server.CreateObject("ADODB.RecordSet")
	rsItems.Open sSQL, oconn    
	if not rsItems.eof then
		'
	else
		'
	end if
	'

	'pull the referral options available info from the database and NOT hard code it!!!!
	sSQL = "select * FROM [Prod_System].[dbo].[form_definition] "
	sSQL = sSQL & " where client_num = '" & sClientcode & "' "
	sSQL = sSQL & " and formcode = '" & sFormCode & "' "
	sSQL = sSQL & " and fieldname like ('ref_%') "
	sSQL = sSQL & " and fieldtype = 'C'  "
	sSQL = sSQL & " order by row_no, col_no "
	'Response.Write sSQL & "<br />"

	Set rsReferrals = Server.CreateObject("ADODB.RecordSet")
	rsReferrals.Open sSQL, oconn    
	if not rsReferrals.eof then
		'
	else
		'
	end if
	'

	sURI = Request.ServerVariables("SERVER_NAME")
	'get any seminar data that is available
	sSQL = "SELECT * "
	sSQL = sSQL & "FROM tblNRD_Seminars "
	sSQL = sSQL & "WHERE Active = 1 AND "
	sSQL = sSQL & " Client_Num = '" & sClientCode & "' AND "
	sSQL = sSQL & " FormCode = '" & sFormCode  & "' AND "
	sSQL = sSQL & " (CONVERT(Char(10), GETDATE(), 20) >= CONVERT(Char(10), Start_Date, 20)) AND "
	sSQL = sSQL & " (CONVERT(Char(10), GETDATE(), 20) <= CONVERT(Char(10), End_Date, 20)) "
	sSQL = sSQL & "ORDER BY End_Date "
	'RW(sSQL)
	Set rsSeminars = Server.CreateObject("ADODB.RecordSet")
	rsSeminars.Open sSQL, oconn    


	'
	'init days in month to 31 for now
	daysInMonth = 31
	k1_daysInMonth = 31
	k2_daysInMonth = 31
	k3_daysInMonth = 31
	k4_daysInMonth = 31
	
	strMissing = ""
	strErrorMsg = ""

	seminar = ""
	
	if request("formcode") = sFormCode then
		'this is probably a postback, so need to verify data
		'response.write "postback<br>"
		seminar = request("seminar")
		if seminar <> "" then
			seminar = replace(seminar, "<span style='display:none;'>", "")
			seminar = replace(seminar, "</span>", "")
		end if
		'response.write "Seminar = " & seminar & "<br>"
		
		cust_id = request("cust_id")
		firstname = request("firstname")
		middlename = request("middlename")
		lastname = request("lastname")
		address = request("address")
		city = request("city")
		state = request("state")
		zipcode = request("zipcode")
		phone = request("q_02")
		email = request("q_03")
		dob_mm = request("q_04_1")
		dob_dd = request("q_04_2")
		dob_yyyy = request("q_04_3")
		if dob_mm <> "" and dob_yyyy <> "" then
			'calc number of days for this birth month
			select case cint(dob_mm)
				case 1, 3, 5, 7, 8, 10, 12:
					daysInMonth = 31
				case 4, 6, 9, 11:
					daysInMonth = 30
				case 2:
					'default to 28, but then check for leap year
					daysInMonth = 28
					if cint(dob_yyyy) mod 4 = 0 then
						daysInMonth = 29
					end if
				case else:
					daysInMoth = 0
			end select
		end if
		gender = request("gender")

		k1_dob_mm = request("kid_dob_mon_1")
		k1_dob_dd = request("kid_dob_day_1")
		k1_dob_yyyy = request("kid_dob_yr_1")
		if k1_dob_mm <> "" and k1_dob_yyyy <> "" then
			'calc number of days for this birth month
			select case cint(k1_dob_mm)
				case 1, 3, 5, 7, 8, 10, 12:
					k1_daysInMonth = 31
				case 4, 6, 9, 11:
					k1_daysInMonth = 30
				case 2:
					'default to 28, but then check for leap year
					k1_daysInMonth = 28
					if cint(k1_dob_yyyy) mod 4 = 0 then
						k1_daysInMonth = 29
					end if
				case else:
					k1_daysInMoth = 0
			end select
		end if
		k1_gender = request("kid_gender_1")

		k2_dob_mm = request("kid_dob_mon_2")
		k2_dob_dd = request("kid_dob_day_2")
		k2_dob_yyyy = request("kid_dob_yr_2")
		if k2_dob_mm <> "" and k2_dob_yyyy <> "" then
			'calc number of days for this birth month
			select case cint(k2_dob_mm)
				case 1, 3, 5, 7, 8, 10, 12:
					k2_daysInMonth = 31
				case 4, 6, 9, 11:
					k2_daysInMonth = 30
				case 2:
					'default to 28, but then check for leap year
					k2_daysInMonth = 28
					if cint(k2_dob_yyyy) mod 4 = 0 then
						k2_daysInMonth = 29
					end if
				case else:
					k2_daysInMoth = 0
			end select
		end if
		k2_gender = request("kid_gender_2")
		
		k3_dob_mm = request("kid_dob_mon_3")
		k3_dob_dd = request("kid_dob_day_3")
		k3_dob_yyyy = request("kid_dob_yr_3")
		if k3_dob_mm <> "" and k3_dob_yyyy <> "" then
			'calc number of days for this birth month
			select case cint(k3_dob_mm)
				case 1, 3, 5, 7, 8, 10, 12:
					k3_daysInMonth = 31
				case 4, 6, 9, 11:
					k3_daysInMonth = 30
				case 2:
					'default to 28, but then check for leap year
					k3_daysInMonth = 28
					if cint(k3_dob_yyyy) mod 4 = 0 then
						k3_daysInMonth = 29
					end if
				case else:
					k3_daysInMoth = 0
			end select
		end if
		k3_gender = request("kid_gender_3")

		k4_dob_mm = request("kid_dob_mon_4")
		k4_dob_dd = request("kid_dob_day_4")
		k4_dob_yyyy = request("kid_dob_yr_4")
		if k4_dob_mm <> "" and k4_dob_yyyy <> "" then
			'calc number of days for this birth month
			select case cint(k4_dob_mm)
				case 1, 3, 5, 7, 8, 10, 12:
					k4_daysInMonth = 31
				case 4, 6, 9, 11:
					k4_daysInMonth = 30
				case 2:
					'default to 28, but then check for leap year
					k4_daysInMonth = 28
					if cint(k4_dob_yyyy) mod 4 = 0 then
						k4_daysInMonth = 29
					end if
				case else:
					k4_daysInMoth = 0
			end select
		end if
		k4_gender = request("kid_gender_4")

		primarycare = request("q_07_1")

		balance = request("balance")
		
		healthwellness = request("healthwellness")
		sportsperf = request("sportsperf")
		kidshlthenews = request("kidshlthenews")
		
		for x = 1 to 99
			itemsRequested(x) = request("item_" & right("00" & x, 2))
		next

		sClientCode = request("clientcode")
		sMarketCode = request("marketcode")
		sProcCode = request("proccode")
		sformcode = request("formcode")
		sebrc_track_id = request("ebrc_track_id")
		suserid = request("userid")
		susertype = request("usertype")
		slookupfile = request("lookupfile")
		sPageTitle = request("PageTitle")
		sClientValidation_JavaCode = request("ClientValidation_JavaCode")
		sBaseFont = request("BaseFont")
		sLogoImg = request("LogoImg")
		sBkImage = request("BkImage")
		sBkColor = request("BkColor")
		'
		'all of the fields have been captured, check for validity
		'First Name, Last Name, City, State and ZipCode are required
		if trim(firstname) = "" then strMissing = strMissing & "firstname,"
		if trim(lastname) = "" then strMissing = strMissing & "lastname,"
		if trim(address) = ""  then strMissing = strMissing & "address,"
		if trim(city) = "" then strMissing = strMissing & "city,"
		if trim(state) = "" then strMissing = strMissing & "state,"
		if trim(zipcode) = "" then strMissing = strMissing & "zipcode,"
		'
		'if any subscriptions are requested, then a valid email address is required
		if healthwellness <> "" or sportsperf <> "" or kidshlthenews <> "" then
			if trim(email) = "" then strMissing = strMissing & "email,"
		end if
		'
		'if any of the dob parts are provided then all of the parts must be provided
		if dob_mm <> "" or dob_dd <> "" or dob_yyyy <> "" then
			if dob_mm = "" then strMissing = strMissing & "a_dob_mm,"
			if dob_dd = "" then strMissing = strMissing & "a_dob_dd,"
			if dob_yyyy = "" then strMissing = strMissing & "a_dob_yyyy,"
		end if
		
		'if strMissing = "" then basic data is valid, verify the data amd build the temp form
		if strMissing = "" then
			'clear the old data field array
			intFields = 0
			for x = 1 to 500
				arrFields(x, 1) = ""
				arrFields(x, 2) = ""
			next
			'
			'now get all of the fields and their current value for later use
			For Each sItem In Request.Form
				intFields = intFields + 1
				arrFields(intFields, 1) = sItem
				arrFields(intFields, 2) = Request.Form(sItem)
			Next
			
			fVerify = true
		else
			'else let the user know something is wrong
			strErrorMsg = "Please correct the highlighted field(s).<br>"
		end if
		
	end if
	
function GetNextSeminar
	sSeminar_ID = rsSeminars("Seminar_ID")
	sSeminar_Date = rsSeminars("Seminar_Date")
	sDateTime = rsSeminars("Field_DateTime")
	sCity = rsSeminars("Field_City")
	sLocation = rsSeminars("Field_Location")
	sAddress = rsSeminars("Field_Address")
	sCityStateZip = rsSeminars("Field_CityStateZip")
	sPhone = rsSeminars("Field_Phone")
End function	
%>
<!DOCTYPE html>
<html lang="en">
	<head>
		<meta charset="utf-8"/>
		<!-- Main SEO Description -->
		<title><%=sPageTitle%></title>
		<meta name="description" content="Generic Boilerplate" />
		<meta name="keywords" content="" />
		<!-- Mobile Meta Tags-->
		<meta name="HandheldFriendly" content="true" />
		<meta name="MobileOptimized" content="320" />
		<meta name="viewport" content="width=device-width, initial-scale=1" />
		<!-- Author -->
		<meta name="author" content="Joe Clark <jclark@cmpkc.com>" />
		<meta name="creation-date" content="04/17/2018" />
		<meta name="company" content="Creative Marketing Programs" />
		<meta name="copyright" content="2018 - Creative Marketing Programs, All Rights Reserved." />
		<!-- Robots -->
		<meta name="robots" content="none" />
		<meta name="robots" content="noindex, nofollow" />
		<meta name="googlebot" content="nofollow" />
		<!-- Style Sheets -->
		<link rel="stylesheet" type="text/css" href="assets/css/bootstrap.min.css">
		<link rel="stylesheet" type="text/css" href="assets/css/style.css">		
		<!-- HTML5 shim, for IE6-9 support of HTML elements -->
		<!--[if lt IE 10]><script src="assets/js/html5-shiv.js" type="text/javascript" charset="utf-8"></script><![end if]-->

		<!--Bootstrap jQuery Support -->
		<script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>

		<script type="text/javascript">
			$(document).ready(function()
			{
				//only allow numerics in zipcode field
				$("#zipcode").keydown(function(event) {
					// Allow: backspace, delete, tab, escape, enter and .
					if ( $.inArray(event.keyCode,[46,8,9,27,13,190]) !== -1 ||
						 // Allow: Ctrl+A
						(event.keyCode == 65 && event.ctrlKey === true) || 
						 // Allow: home, end, left, right
						(event.keyCode >= 35 && event.keyCode <= 39)) {
							 // let it happen, don't do anything
							 return;
					}
					else {
						// Ensure that it is a number and stop the keypress
						if (event.shiftKey || (event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105 )) {
							event.preventDefault(); 
						}   
					}
				});
						
				///////////////
				// FUNCTIONS //
				////////////////////////////////////////////////////////////////////////////////////////////////////////
				/* Calculates days in a month for a given year */
				function daysInMonth(month, year)
				{
					return new Date(year, month, 0).getDate();
				}
				/* Used for setting the days when the month changes see Events below. */
				function setDays(month, day, year)
				{
					// Default to 2012 Leap Year for February.  Currently hard coded.
					var mm = $(month).val();
					var yyyy = (typeof(year)==='undefined'?'2012':$(year).val());
					// Clear Combo box before load and add a blank to the beginning
					$(day).empty().append('<option value=""></option>');
					// Force to a leap year so that Feb always has 29 days
					var Monthdays = daysInMonth(mm,yyyy);
					for(i=1; i<Monthdays+1; i++)
					{
						var strDays = (i<10?"0"+ i.toString(): i.toString());
						$(day).append('<option value="'+strDays+'">'+strDays+'</option>');
					}
				}

				//////////
				// MISC //
				////////////////////////////////////////////////////////////////////////////////////////////////////////
				
				/* Get the window width */
				//alert($(window).width());

				////////////////////
				// SETUP CONTROLS //
				////////////////////////////////////////////////////////////////////////////////////////////////////////

				/* Setup DOB combo boxes */
				var mm   = new Array('#dob-mm',   '#child1-dob-mm',   '#child2-dob-mm',   '#child3-dob-mm',   '#child4-dob-mm');
				var dd   = new Array('#dob-dd',   '#child1-dob-dd',   '#child2-dob-dd',   '#child3-dob-dd',   '#child4-dob-dd');
				var yyyy = new Array('#dob-yyyy', '#child1-dob-yyyy', '#child2-dob-yyyy', '#child3-dob-yyyy', '#child4-dob-yyyy');
				var today = new Date();
				var today_mm = today.getMonth()+1; //January is 0!
				var today_yyyy = today.getFullYear();
				var i;
				
				/* Setup Days of Month in combo boxes below. */
				var Monthdays = daysInMonth(today_mm,today_yyyy);
				var dd_index;
				for (dd_index = 0; dd_index < dd.length; ++dd_index)
				{
				//	for(i=1; i<Monthdays+1; i++)
				//	{
				//		var strDays = (i<10?"0"+ i.toString(): i.toString());
				//		$(dd[dd_index]).append('<option value="'+strDays+'">'+strDays+'</option>');
				//	}
				}

			});

			function checkForm() {
				//alert("Checking data");
				var ok = true;
				//check for required fields

				if ($("#firstname").val() == "") {
					$("#firstname").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#firstname").removeClass("alert-danger");
				}
				if ($("#lastname").val() == "") {
					$("#lastname").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#lastname").removeClass("alert-danger");
				}

				if ($("#address").val() == "") {
					$("#address").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#address").removeClass("alert-danger");
				}
				if ($("#city").val() == "") {
					$("#city").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#city").removeClass("alert-danger");
				}
				if ($("#state").val() == "") {
					$("#state").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#state").removeClass("alert-danger");
				}
				if ($("#zipcode").val() == "" | $("#zipcode").val().length != 5) {
					$("#zipcode").addClass("alert-danger");
					ok = false;
				}
				else
				{
					$("#zipcode").removeClass("alert-danger");
				}

				//if email is provided, it must be valid
				//if ($("#email").val() != "") {
					//changed 4/22/18 Email IS required!!
					if (!isEmailValid($("#email").val())) {
						$("#email").addClass("alert-danger");
						ok = false;
					}
					else {
						$("#email").removeClass("alert-danger");
					}
				//}

				//if phone is provided, it must be 10 digits
				if ($("#phone").val() != "") {
					var phoneno = /^\(?([0-9]{3})\)?[-. ]?([0-9]{3})[-. ]?([0-9]{4})$/;
					var ph = $("#phone").val();
					if (ph.match(phoneno)) {
						$("#phone").removeClass("alert-danger");
					} else {
						$("#phone").addClass("alert-danger");
						ok = false;
					}
				}

				//if any birthdate part is entered, then all are required.
				if ($("#dob-mm").val() != "" | $("#dob-dd").val() != "" | $("#dob-yyyy").val() != ""){
					if ($("#dob-mm").val() == "") {
						$("#dob-mm").addClass("alert-danger");
						ok = false;
					}
					else
					{
						$("#dob-mm").removeClass("alert-danger");
					}

					if ($("#dob-dd").val() == "") {
						$("#dob-dd").addClass("alert-danger");
						ok = false;
					}
					else
					{
						$("#dob-dd").removeClass("alert-danger");
					}

					if ($("#dob-yyyy").val() == "") {
						$("#dob-yyyy").addClass("alert-danger");
						ok = false;
					}
					else
					{
						$("#dob-yyyy").removeClass("alert-danger");
					}
				}

				if (ok == false) {
					alert("Please correct the highlighted fields.");
					return false;
				} else {
					$("#ebrc-form").submit();
					//return true;
				}
			}

			function isEmailValid(email) {
				var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
				return re.test(email);
			}
		</script>

	</head>

	<body class="bg-light">
		<!-- Google Tag Manager -->
		<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
			new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
			j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
			'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
			})(window,document,'script','dataLayer','<%=gTM%>');</script> 
		<!-- End Google Tag Manager -->			
		<header>
			<div class="container">
				<div class="row">
					<div class="col-12 py-4 mx-auto text-center">
						<img class="img-fluid" src="./assets/images/st_joseph_logo.png" 		alt="<%=sPageTitle%>">
					</div>
				</div>
			</div>
		</header>
		<%if fVerify = false then%>
			<div id="fadeIn"></div>
			<main class="container" role="main">
				<section class="jumbotron mx-3 mb-3">
					<div class="row justify-content-center">
						<img class="ffak-img mb-md-3" src="./assets/images/ffak_graphic.png" alt="<%=sPageTitle%>">
						<div class="col-md-9">
							<h4 class="mb-0 text-center text-md-left">We would like to send you a <strong>FREE FIRST-AID KIT</strong> with our thanks for filling out this form!</h4>
						</div>
						
					</div>
				</section>
				<p class="text-center font-italic text-muted"><small>Your privacy is important to us. The information we gather here will only be used by <%=sPageTitle%> and will not be shared in any way with a third party.</small></p>
				<!--[if lte IE 8]>
					<div id="upgrade" class="old_IE">Like a growing number across the Web, this site no longer supports old versions of Internet Explorer and your viewing experience may be affected.<br>
					We recommend that you download a newer browser.</div>
					<br><br>
				<![endif]-->
				<noscript>
					<div id="overlay" class="overlay">
						<div id="noscript" class="noscript" >
							For full functionality of this page, it is necessary to enable JavaScript. 
							Here are the instructions to enable JavaScript in your <a href="http://www.enable-javascript.com" target="_blank"> 
							web browser</a>.
							<br><br>
							After you enable javascript, click <a href="#">here</a> to refresh this page.
						</div>
					</div>
					<br><br>
				</noscript>
				<form id="ebrc-form" name="ebrc-form" method="post" action="<%=sAction%>">
					<!-- Hidden Data -->
					<input type="hidden" id="clientcode"                name="clientcode"                value="<%=sClientCode%>" />
					<input type="hidden" id="MarketCode"                name="marketcode"                value="<%=sMarketCode%>" />
					<input type="hidden" id="proccode"                  name="proccode"                  value="<%=sProcCode%>" />
					<input type="hidden" id="formcode"                  name="formcode"                  value="<%=sFormCode%>" />
					<input type="hidden" id="ebrc_track_id"             name="ebrc_track_id"             value="<%=sebrc_track_id%>" />
					<input type="hidden" id="userid"                    name="userid"                    value="<%=suserid%>" />
					<input type="hidden" id="usertype"                  name="usertype"                  value="<%=susertype%>" />
					<input type="hidden" id="lookupfile"                name="lookupfile"                value="<%=slookupfile%>" />
					<input type="hidden" id="pagetitle"                 name="PageTitle"                 value="<%=spagetitle%>" />
					<input type="hidden" id="clientvalidation_javacode" name="ClientValidation_JavaCode" value="<%=sClientValidation_JavaCode%>" />
					<input type="hidden" id="basefont"                  name="BaseFont"                  value="<%=sBaseFont%>" />
					<input type="hidden" id="logoimg"                   name="LogoImg"                   value="<%=sLogoImg%>" />
					<input type="hidden" id="bkimage"                   name="BkImage"                   value="<%=sBkImage%>" />
					<input type="hidden" id="bkcolor"                   name="BkColor"                   value="<%=sBkColor%>" />
					<!-- Start of Page -->
					<%if strErrorMsg <> "" then%>
						<div id="error" class="error">
							<%=strErrorMsg%>
						</div>
					<%end if%>
					<div id="divUser_Info" >
						<div id="divPersonal" >
							<!-- START USER CONTENT -->
							<div class="form-row">
								<div class="form-group col-md-4">
									<label id="cust_id-label" for="cust_id">Gift Code:</label>
									<input type="text" id="cust_id" name="cust_id" value="<%=cust_id%>" maxlength="10" class="form-control" />
									<p class="small"> (Gift Code located in the address block of your postcard)</p>
								</div>								
							</div>
							<div class="form-row">
								<div class="form-group col-md-5">
									<label id="firstname-label" for="firstname"><span class="required">*</span>First Name:</label>
									<input type="text" id="firstname" name="firstname" value="<%=firstname%>" maxlength="25" class="form-control <%if instr(strMissing, "firstname,") then%> missing <%end if%>" data-required="true" data-notblank="true" />
								</div>
								<div class="form-group col-md-2">
									<label for="middlename">Middle Initial:</label>
									<input type="text" id="middlename" name="middlename" class="form-control" value="<%=middlename%>" maxlength="1">
								</div>
								<div class="form-group col-md-5">
									<label id="lastname-label" for="lastname"><span class="required">*</span>Last Name:</label>
									<input type="text" id="lastname" name="lastname" value="<%=lastname%>" maxlength="25" class="form-control <%if instr(strMissing, "lastname,") then%> missing <%end if%>" data-required="true" data-notblank="true" />
								</div>
							</div>
							<div class="form-row">
								<div class="form-group col-12">
									<label id="address-label" for="address"><span class="required">*</span>Address:</label>
									<input type="text" id="address" value="<%=address%>" maxlength="80" size="108" class="form-control <%if instr(strMissing, "address,") then%> missing <%end if%>" name="address" data-required="true" data-notblank="true" />
								</div>
							</div>
							<div class="form-row">
								<div class="form-group col-md-6">
									<label id="city-label" for="city"><span class="required">*</span>City:</label>
									<input type="text" id="city" value="<%=city%>" maxlength="25" size="58" class="form-control <%if instr(strMissing, "city,") then%> missing <%end if%>" name="city" data-required="true" data-notblank="true" />
								</div>
								<div class="form-group col-md-2">
									<label id="state-label" for="state"><span class="required">*</span>State:</label>
									<select id="state" name="state" class="form-control"> <%if instr(strMissing, "state,") then%> class="missing" <%end if%>" data-required="true" data-notblank="true">
										<option value="" <%if state = "" then%> selected="selected"<%end if%>></option>
										<option value="AK" <%if state = "AK" then%> selected="selected"<%end if%>>AK</option><option value="AL" <%if state = "AL" then%> selected="selected"<%end if%>>AL</option>
										<option value="AR" <%if state = "AR" then%> selected="selected"<%end if%>>AR</option><option value="AZ" <%if state = "AZ" then%> selected="selected"<%end if%>>AZ</option>
										<option value="CA" <%if state = "CA" then%> selected="selected"<%end if%>>CA</option><option value="CO" <%if state = "CO" then%> selected="selected"<%end if%>>CO</option>
										<option value="CT" <%if state = "CT" then%> selected="selected"<%end if%>>CT</option><option value="DC" <%if state = "DC" then%> selected="selected"<%end if%>>DC</option>
										<option value="DE" <%if state = "DE" then%> selected="selected"<%end if%>>DE</option><option value="FL" <%if state = "FL" then%> selected="selected"<%end if%>>FL</option>
										<option value="GA" <%if state = "GA" then%> selected="selected"<%end if%>>GA</option><option value="HI" <%if state = "HI" then%> selected="selected"<%end if%>>HI</option>
										<option value="IA" <%if state = "IA" then%> selected="selected"<%end if%>>IA</option><option value="ID" <%if state = "ID" then%> selected="selected"<%end if%>>ID</option>
										<option value="IL" <%if state = "IL" then%> selected="selected"<%end if%>>IL</option><option value="IN" <%if state = "IN" then%> selected="selected"<%end if%>>IN</option>
										<option value="KS" <%if state = "KS" then%> selected="selected"<%end if%>>KS</option><option value="KY" <%if state = "KY" then%> selected="selected"<%end if%>>KY</option>
										<option value="LA" <%if state = "LA" then%> selected="selected"<%end if%>>LA</option><option value="MA" <%if state = "MA" then%> selected="selected"<%end if%>>MA</option>
										<option value="MD" <%if state = "MD" then%> selected="selected"<%end if%>>MD</option><option value="ME" <%if state = "ME" then%> selected="selected"<%end if%>>ME</option>
										<option value="MI" <%if state = "MI" then%> selected="selected"<%end if%>>MI</option><option value="MN" <%if state = "MN" then%> selected="selected"<%end if%>>MN</option>
										<option value="MO" <%if state = "MO" then%> selected="selected"<%end if%>>MO</option><option value="MS" <%if state = "MS" then%> selected="selected"<%end if%>>MS</option>
										<option value="MT" <%if state = "MT" then%> selected="selected"<%end if%>>MT</option><option value="NC" <%if state = "NC" then%> selected="selected"<%end if%>>NC</option>
										<option value="ND" <%if state = "ND" then%> selected="selected"<%end if%>>ND</option><option value="NE" <%if state = "NE" then%> selected="selected"<%end if%>>NE</option>
										<option value="NH" <%if state = "NH" then%> selected="selected"<%end if%>>NH</option><option value="NJ" <%if state = "NJ" then%> selected="selected"<%end if%>>NJ</option>
										<option value="NM" <%if state = "NM" then%> selected="selected"<%end if%>>NM</option><option value="NV" <%if state = "NV" then%> selected="selected"<%end if%>>NV</option>
										<option value="NY" <%if state = "NY" then%> selected="selected"<%end if%>>NY</option><option value="OH" <%if state = "OH" then%> selected="selected"<%end if%>>OH</option>
										<option value="OK" <%if state = "OK" then%> selected="selected"<%end if%>>OK</option><option value="OR" <%if state = "OR" then%> selected="selected"<%end if%>>OR</option>
										<option value="PA" <%if state = "PA" then%> selected="selected"<%end if%>>PA</option><option value="RI" <%if state = "RI" then%> selected="selected"<%end if%>>RI</option>
										<option value="SC" <%if state = "SC" then%> selected="selected"<%end if%>>SC</option><option value="SD" <%if state = "SD" then%> selected="selected"<%end if%>>SD</option>
										<option value="TN" <%if state = "TN" then%> selected="selected"<%end if%>>TN</option><option value="TX" <%if state = "TX" then%> selected="selected"<%end if%>>TX</option>
										<option value="UT" <%if state = "UT" then%> selected="selected"<%end if%>>UT</option><option value="VA" <%if state = "VA" then%> selected="selected"<%end if%>>VA</option>
										<option value="VT" <%if state = "VT" then%> selected="selected"<%end if%>>VT</option><option value="WA" <%if state = "WA" then%> selected="selected"<%end if%>>WA</option>
										<option value="WI" <%if state = "WI" then%> selected="selected"<%end if%>>WI</option><option value="WV" <%if state = "WV" then%> selected="selected"<%end if%>>WV</option>
										<option value="WY" <%if state = "WY" then%> selected="selected"<%end if%>>WY</option>
									</select>
								</div>
								<div class="form-group col-md-4">
									<label id="zipcode-label" for="zipcode"><span class="required">*</span>Zip:</label>
									<input type="text" id="zipcode" name="zipcode" value="<%=zipcode%>" maxlength="5" class="form-control <%if instr(strMissing, "zipcode,") then%> missing <%end if%>" data-required="true" data-type="digits" data-rangelength="[5,5]" data-trigger="keyup change" data-error-message="Zipcode must be 5 numbers." />		
								</div>
							</div>
							<div class="form-row">
								<div class="form-group col-lg-2">
									<label id="phone-label" for="phone">Telephone:</label>
									<input type="tel" id="phone" value="<%=phone%>" maxlength="12" class="form-control" name="q_02" placeholder="123-456-7890"/>
								</div>
								<div class="form-group col-lg-4">
									<label id="email-label" for="email"><span class="required">*</span>Email:</label>
									<input type="text" id="email" name="q_03" value="<%=email%>" maxlength="50" placeholder="example@domain.com" class="form-control <%if instr(strMissing, "email,") then%> missing <%end if%>" onblur="cmp_email_required();" />
								</div>
							</div>
							<hr>
							<div class="form-inline">
								<label class="mb-2" id="dob-label">Date of Birth:</label>
								<select id="dob-mm" name="q_04_1" class="form-control ml-lg-2 mr-sm-2 mb-2" <%if instr(strMissing, "dob_mm,") then%>class="missing"<%end if%>>
									<option value="" <%if dob_mm = "" then%> selected="selected"<%end if%>></option>
									<%for x = 1 to 12%>
									<option value="<%=x%>" <%if dob_mm = trim(x) then%> selected="selected"<%end if%>><%=monthname(x, true)%></option>
									<%next%>
								</select>
								<select id="dob-dd" name="q_04_2" class="form-control m mr-sm-2 mb-2" <%if instr(strMissing, "dob_dd,") then%>class="missing"<%end if%>>
									<option value="" <%if dob_dd = "" then%> selected="selected"<%end if%>></option>
									<%for x = 1 to daysInMonth%>
									<option value="<%=right("00" & x, 2)%>" <%if trim(dob_dd) = right("00" & x, 2) then%> selected="selected"<%end if%>><%=x%></option>
									<%next%>
								</select>
								<select id="dob-yyyy" name="q_04_3" class="form-control m mr-sm-2 mb-2" <%if instr(strMissing, "dob_yyyy,") then%>class="missing"<%end if%>>
									<option value=""  <%if dob_yyyy = "" then%> selected="selected"<%end if%>></option>
									<%for x = year(now) - 100 to year(now)%>
									<option value="<%=x%>" <%if dob_yyyy = trim(x) then%> selected="selected"<%end if%>><%=x%></option>
									<%next%>
								</select>
								<div class="form-group">
									<label id="gender-label" class="mb-2 ml-sm-4 mr-sm-2">Gender:</label>
									<div class="form-check form-check-inline">
										<label id="gender-male-label" class="form-check-label mb-2 mr-sm-2" for="gender_male">
											<input type="radio" class="form-check-input mb-2 ml-sm-2 mr-sm-2" id="gender_male" name="gender" value="Male" <%if gender = "Male" then%> checked="checked"<%end if%> />
											Male
										</label>
									</div>
									<div class="form-check form-check-inline">
										<label id="gender-female-label" class="form-check-label mb-2 mr-sm-2" for="gender_female">
											<input type="radio" class="form-check-input mb-2 mr-sm-2" id="gender_female" name="gender" value="Female"  <%if gender = "Female" then%> checked="checked"<%end if%> />
											Female
										</label>
									</div>
								</div>
							</div>
							<div class="form-row">
								<span id="pcp-label">Do you have a primary care doctor?</span>
								<br class="d-none d-xs-md">
								<div class="form-check form-check-inline mx-3">
									<label id="pcp-yes-label" class="form-check-label" for="pcp-yes">
										<input type="radio" id="pcp-yes" class="form-check-input" name="q_07_1" <%if primarycare = "Yes" then%>checked="checked"<%end if%> value="Yes" />
										Yes
									</label>
								</div>
								<div class="form-check form-check-inline mx-3">
									<label id="pcp-no-label" class="form-check-label" for="pcp-no">
										<input type="radio" id="pcp-no" class="form-check-input" name="q_07_1" <%if primarycare = "No" then%>checked="checked"<%end if%> value="No" />
										No
									</label>
								</div>
								<br>
								<div class="col-12">
									<small>For help finding a primary care physician or specialist, call 816-429-3714.</small>
								</div>
							</div>
							<hr>
							<div class="form-group row item-list-row">
								<div class="col-12">
									<div class="mb-2 info-item-heading">I’m interested in information about:</div>
									<div class="row">
										<%if not rsItems.eof then%>
											<%dim item_id, item_num, item_text%>
											<%while not rsItems.eof%>
												<%item_id = trim(rsItems("fieldname"))%>
												<%item_num = right(item_id, 2)%>
												<%item_text = rsItems("fieldlabel")%>
													<div class="col-lg-6 mb-2 item-list-container">
														<div class="form-check ">
															<label class="form-check-label" for="<%=item_id%>">
																<input type="checkbox" class="form-check-input" id="<%=item_id%>" name="<%=item_id%>" value="<%=item_text%>"  <%if itemsRequested(item_num) <> "" then%>checked="checked"<%end if%> />
																<%=item_text%>
															</label>
														</div>
													</div>
												<%rsItems.movenext%>
											<%wend%>
										<%end if%>
									</div>	
								</div>
							</div>

						</div>
						<!-- END USER CONTENT -->
						<div class="row justify-content-center">
							<div class="col-lg-3">
							<input type="submit" id="submit" name="submit" value="Submit" class="btn-block btn brandBtn mr-md-2 my-3" onclick="return checkForm();"/>
							</div>
							<div class="">
								<input type="Reset" id="reset" name="Reset" value="Clear" class="btn-block btn btn-secondary my-3" onclick="clearHighLights();">
							</div>
						</div>	
					</div>  <!-- user-info -->
				</form>

				<%else%>
					
					<!-- confirmation page  -->
					<div id="ver_page" class="container">
						<div id="ver_data" class="row">
							<div class="col-12 text-center pb-2">							
								<h3>Thank you for your response. 
									<br>
									<small>Please verify the information you've provided.</small>
								</h3>
							</div>
						</div>
						<div class="row mx-auto ver-info">
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">First Name:</p>
									<p><%=firstname%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">Last Name:</p>
									<p><%=lastname%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">Address:</p>
									<p><%=address%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">City:</p>
									<p><%=city%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">State:</p>
									<p><%=state%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">Zip Code:</p>
									<p><%=zipcode%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">Telephone:</p>
									<p><%=phone%></p>
								</div>
							</div>
							<div class="col-12">
								<div class="row">
									<p class="font-weight-bold mr-2">Email:</p>
									<p><%=email%></p>
								</div>
							</div>							
						</div>

						<form id="eBRCVer" name="eBRCVer" method="post" action="../../../transadd.asp">
							<%for x = 1 to intFields%>
								<input type="hidden" id="<%=arrFields(x, 1)%>" name="<%=arrFields(x, 1)%>" value="<%=arrFields(x, 2)%>">
							<%next%>

							<br>
							<div class="row justify-content-center">
								<div class="col-lg-3">
								<input type="submit" id="submit1" name="submit1" value="Submit" class="btn-block btn brandBtn mr-md-2 my-3" onclick="return checkForm();"/>
								</div>
								<div class="">
									<input type="Reset" ID="back" name="back" onClick="MoveBack(); return false;" value="Back" class="btn-block btn btn-secondary my-3">
								</div>
							</div>								
<!-- 							<div id="verSubmit">
								<div style="width:100%; text-align: center;">
									<input type="submit" id="submit1" name="submit1" value="Submit" />
									&nbsp;&nbsp;&nbsp;
									<input type="Reset" ID="back" name="back" onClick="MoveBack(); return false;" value="Back">
								</div>
							</div> -->

							<div id="divSubmitted" style="display:none;">
								<span style="text-align:center;font-size:12pt;font-weight:bold;color:blue;">
									Your information has been submitted.<br />
									You will now be redirected to the COMPANY NAME website.
								</span><br />
							</div>
						</form>

						<script type="text/javascript">
							function safeSubmit() {
								$("#verSubmit").hide();
								$("#divSubmitted").show();

								$("#eBRCVer").submit();
								return true;
							}

							function MoveBack() {
								history.back();
							}
						</script>
					</div>

				<%end if%>
			</div>
		</main><!-- main content container -->
		<footer class="py-5 justify-content-center position-sticky">
				<div class="col-lg-2 footer-comp-info text-center">
					<img src="./assets/images/st_joseph_logo_white.png" alt="<%=sPageTitle%>" class="img-fluid mb-2">
					<div>1000 Carondelet Drive</div>
					<div>Kansas City, MO 64114</div>
					<div>816-429-3714</div>
				</div>
		</footer>
		<script id="__bs_script__">//<![CDATA[
			document.write("<script async src='http://HOST:3000/browser-sync/browser-sync-client.js?v=2.26.7'><\/script>".replace("HOST", location.hostname));
		//]]></script>
	</body>
</html>