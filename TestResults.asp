<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	end if
%>
<HTML>
	<HEAD>
		<TITLE>Test Results</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			margin: 0 20px 0 20px;
		}
		.toolbarImg{
			margin: 0 25px 0 25px;
		}
		</STYLE>
	</HEAD>
	
	<BODY style="background-color: 50B2DF;"> 
		<hr size="5" style="color: black; background-color: black;">
		<p>
		<center>
		<a href="http://cozmathematics.1apps.com/main.asp">
		<img src="Images\CMLogo.png" width="50%"></img><a></center>
		</p>
		
		<br>
		<center>
		<a href="http://cozmathematics.1apps.com/main.asp">
		<abbr title="Home"><img class="toolbarImg" src="Images\Home_SYM.png" width="8%"></img></abbr></a>
		<a href="http://cozmathematics.1apps.com/Practice.asp">
		<abbr title="Practice/Quizzes"><img class="toolbarImg" src="Images\Practice_SYM.png" width="8%"></img></abbr></a>
		<a href="http://cozmathematics.1apps.com/test.asp">
		<abbr title="Test"><img class="toolbarImg" src="Images\Test.png" width="8%"></img></abbr></a>
		<a href="http://cozmathematics.1apps.com/Leaderboards.asp">
		<abbr title="Leaderboards"><img class="toolbarImg" src="Images\Leaderboard_SYM.png" width="8%"></img></abbr></a>
		
		<a href="http://cozmathematics.1apps.com/Logout.asp">
		<abbr title="Logout"><img class="toolbarImg" src="Images\Logout_SYM.png" width="8%"></img></abbr></a>
		</center>
		<hr size="5" style="color: black; background-color: black;">
		
		</Center>
		
		<FONT SIZE=5>
		<%
			Const adLockOptimistic = 3
	
			Dim objConn
			Dim strConnection
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\Accounts.mdb")
									
			objConn.Open (strConnection)
			
			Dim strSQL
			strSQL = "SELECT * FROM Accounts"
			
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open strSQL, objConn, , adlockOptimistic
			
			Do while not objRS.EOF
				Dim AUser
				AUser = objRS("UserName")
				
				
				if Ucase(Session("Username")) = UCase(AUser) then
					Dim PlayerRank 
					PlayerRank = objRS("Rank")
					if int(PlayerRank) < int(Session("TestScore")) then
						Session("Rank") = Session("TestScore")
						objRS("Rank") = Session("Rank")
						objRS.Update
						Response.Write("You Managed to set a new high score. Nice! <br>")
						Exit do
					else
						Response.Write("The score you got wasn't as good as your personal best. Try Again! <br>")
						Exit Do
					end if
				end if
				objRS.MoveNext
			Loop
			
			
			Response.Write("In 60 seconds you managed to score a total of " & Session("TestScore") & " points.")
			Response.Write("<BR>To check if you made it onto the leaderboards click <a href=""leaderboards.asp"">Here</a>.")
			Response.Write("<br>To take the test again. Click <a href=""Test.asp"">Here</a>")
			
			Session("TestScore") = 0
		%>
		
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>