<!DOCTYPE html>
<html>
<head>
	<title><%= Language.ProjectPhrase("BodyTitle") %></title>
<% If gsExport = "" Or gsExport = "print" Then %>
<link rel="stylesheet" type="text/css" href="bootstrap/css/bootstrap.css">
<% End If %>
<% If gsExport = "" Then %>
<link rel="stylesheet" type="text/css" href="css/jquery.fileupload-ui.css">
<% End If %>
<% If gsExport = "" Or gsExport = "print" Then %>
<link rel="stylesheet" type="text/css" href="<%= EW_PROJECT_STYLESHEET_FILENAME %>">
<% If ew_IsMobile() Then %>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" type="text/css" href="css/ewmobile.css">
<% End If %>
<% If gsExport = "print" And Request.QueryString("pdf") = "1" And EW_PDF_STYLESHEET_FILENAME <> "" Then ' ??? %>
<link rel="stylesheet" type="text/css" href="<%= EW_PDF_STYLESHEET_FILENAME %>">
<% End If %>
<script type="text/javascript" src="<%= ew_jQueryFile("jquery-%v.min.js") %>"></script>
<% If ew_IsMobile() Then %>
<link rel="stylesheet" type="text/css" href="<%= ew_jQueryFile("jquery.mobile-%v.min.css") %>">
<script type="text/javascript">
jQuery(document).bind("mobileinit", function() {
	jQuery.mobile.ajaxEnabled = false;
	jQuery.mobile.ignoreContentEnabled = true;
});
</script>
<script type="text/javascript" src="<%= ew_jQueryFile("jquery.mobile-%v.min.js") %>"></script>
<% End If %>
<% End If %>
<% If gsExport = "" Then %>
<script type="text/javascript" src="bootstrap/js/bootstrap.min.js"></script>
<script type="text/javascript" src="jqueryfileupload/jquery.ui.widget.js"></script>
<script type="text/javascript" src="jqueryfileupload/jqueryfileupload.min.js"></script>
<script type="text/javascript">
var EW_LANGUAGE_ID = "<%= gsLanguage %>";
var EW_DATE_SEPARATOR = "/" || "/"; // Default date separator
var EW_DECIMAL_POINT = "<%= EW_DECIMAL_POINT %>";
var EW_THOUSANDS_SEP = "<%= EW_THOUSANDS_SEP %>";
var EW_MAX_FILE_SIZE = <%= EW_MAX_FILE_SIZE %>; // Upload max file size
var EW_UPLOAD_ALLOWED_FILE_EXT = "<%= EW_UPLOAD_ALLOWED_FILE_EXT %>"; // Allowed upload file extension
var EW_FIELD_SEP = ", "; // Default field separator
// Ajax settings
var EW_LOOKUP_FILE_NAME = "pom_ewlookup11.asp"; // Lookup file name
var EW_AUTO_SUGGEST_MAX_ENTRIES = <%= EW_AUTO_SUGGEST_MAX_ENTRIES %>; // Auto-Suggest max entries
// Common JavaScript messages
var EW_DISABLE_BUTTON_ON_SUBMIT = true;
var EW_IMAGE_FOLDER = "images/"; // Image folder
var EW_UPLOAD_URL = "<%= EW_UPLOAD_URL %>"; // Upload url
var EW_UPLOAD_THUMBNAIL_WIDTH = <%= EW_UPLOAD_THUMBNAIL_WIDTH %>; // Upload thumbnail width
var EW_UPLOAD_THUMBNAIL_HEIGHT = <%= EW_UPLOAD_THUMBNAIL_HEIGHT %>; // Upload thumbnail height
var EW_USE_JAVASCRIPT_MESSAGE = false;
<% If ew_IsMobile() Then %>
var EW_IS_MOBILE = true;
<% Else %>
var EW_IS_MOBILE = false;
<% End If %>
<% If EW_MOBILE_REFLOW Then %>
var EW_MOBILE_REFLOW = true;
<% Else %>
var EW_MOBILE_REFLOW = false;
<% End If %>
</script>
<% End If %>
<% If gsExport = "" Or gsExport = "print" Then %>
<script type="text/javascript" src="js/jsrender.min.js"></script>
<script type="text/javascript" src="js/pom_ew11.js"></script>
<script type="text/javascript" src="js/pom_ewvalidator.js"></script>
<% End If %>
<% If gsExport = "" Then %>
<script type="text/javascript" src="js/pom_userfn11.js"></script>
<script type="text/javascript">
<%= Language.ToJSON() %>
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="generator" content="ASPMaker v11.0.3">
</head>
<body>
<% If gsExport = "" Or gsExport = "print" Then %>
<% If ew_IsMobile() Then %>
<div data-role="page">
	<div data-role="header">
		<a href="pom_mobilemenu.asp"><%= Language.Phrase("MobileMenu") %></a>
		<h1 id="ewPageTitle"></h1>
	<% If IsLoggedIn() Then %>
		<a href="pom_logout.asp"><%= Language.Phrase("Logout") %></a>
	<% ElseIf Right(Request.ServerVariables("URL"), Len("pom_login.asp")) <> "pom_login.asp" Then %>
		<a href="pom_login.asp"><%= Language.Phrase("Login") %></a>
	<% End If %>
	</div>
<% End If %>
<% End If %>
<% If Not gbSkipHeaderFooter Then %>
<% If gsExport = "" Then %>
<div class="ewLayout">
<% If Not ew_IsMobile() Then %>
	<!-- header (begin) --><!-- *** Note: Only licensed users are allowed to change the logo *** -->
  <div id="ewHeaderRow" class="ewHeaderRow"><img src="images/aspmkrlogo1.png" alt="" style="border: 0;"></div>
	<!-- header (end) -->
<% End If %>
<% If ew_IsMobile() Then %>
	<div data-role="content" data-enhance="false">
	<table id="ewContentTable" class="ewContentTable">
		<tr>
<% Else %>
	<!-- content (begin) -->
	<table id="ewContentTable" class="ewContentTable">
		<tr><td class="ewMenuColumn">
			<!-- left column (begin) -->
<% Server.Execute("pom_ewmenu.asp") %>
			<!-- left column (end) -->
		</td>
<% End If %>
		<td id="ewContentColumn" class="ewContentColumn">
			<!-- right column (begin) -->
				<h4 class="ewSiteTitle"><%= Language.ProjectPhrase("BodyTitle") %></h4>
<% End If %>
<% End If %>
