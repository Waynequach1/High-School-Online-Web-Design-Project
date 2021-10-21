<%@ Language=VBscript %>
<% Option Explicit %>

<%
	Dim Remaining
	Dim SecondsLeft
	Dim MinutesLeft
	Dim CurTime
	CurTime = Timer
	Remaining = Int(Session("TimeLeft") - CurTime)
	Response.Write("<FONT SIZE=6>")
	Response.Write("<CENTER>")
	SecondsLeft = Remaining
	if SecondsLeft > 0 then
		Response.Write("You have " & SecondsLeft & " seconds left.")
	else
		Response.Write("You ran out of time. Hit Refresh or submit to send in your score.")
	end if
	Response.Write("<br>")
	Response.Write("Your current score: " & Session("TestScore"))
	
%>

<HTML>
	<HEAD>
		<TITLE>Test Timer</TITLE>
	</HEAD>
		<meta http-equiv="refresh" content="0.5" >
	<BODY>
	</BODY>
</HTML>