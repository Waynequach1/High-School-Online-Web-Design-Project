<%@ Language=VBscript %>
<% Option Explicit %>

<%
	Session("Verified") = false
	Session("ErrorMsg") = "You have logged out sucessfully."
	Response.Redirect("login.asp")

%>