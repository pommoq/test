<%

		' Display elapsed time
		If EW_DEBUG_ENABLED Then Response.Write ew_CalcElapsedTime(StartTimer)
%>
<% If gsExport = "" Then %>
<% If Not gbSkipHeaderFooter Then %>
			<!-- right column (end) -->
		</td></tr>
	</table>
	<!-- content (end) -->
<% If Not ew_IsMobile() Then %>
	<!-- footer (begin) --><!-- *** Note: Only licensed users are allowed to remove or change the following copyright statement. *** -->
	<div id="ewFooterRow" class="ewFooterRow">
		<div class="ewFooterText"><%= Language.ProjectPhrase("FooterText") %></div>
		<!-- Place other links, for example, disclaimer, here -->		
	</div>
	<!-- footer (end) -->	
<% End If %>
</div>
<% End If %>
<% If gsExport = "" Or gsExport = "print" Then %>
<% If ew_IsMobile() Then %>
	</div>
	<!-- footer (begin) --><!-- *** Note: Only licensed users are allowed to remove or change the following copyright statement. *** -->
<!-- *** Remove comment lines to show footer for mobile
	<div data-role="footer">
		<h4><%= Language.ProjectPhrase("FooterText") %></h4>
	</div>
*** -->
	<!-- footer (end) -->	
</div>
<script type="text/javascript">
$("#ewPageTitle").html($("#ewPageCaption").text());
</script>
<% End If %>
<% if Request.QueryString("_row").Count > 0 Then %>
<script type="text/javascript">
jQuery.later(1000, null, function() {
	jQuery("#<%= Request.QueryString("_row") %>").each(function() { this.scrollIntoView(); }
});
</script>
<% End If %>
<% End If %>
<!-- message box -->
<div id="ewMsgBox" class="modal hide" data-backdrop="false"><div class="modal-body"></div><div class="modal-footer"><a href="#" class="btn btn-primary ewButton" data-dismiss="modal" aria-hidden="true"><%= Language.Phrase("MessageOK") %></a></div></div>
<!-- tooltip -->
<div id="ewTooltip"></div>
<% End If %>
<% If gsExport = "" Then %>
<script type="text/javascript">
// Write your global startup script here
// document.write("page loaded");
</script>
<% End If %>
</body>
</html>
