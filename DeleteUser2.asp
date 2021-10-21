<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	elseif Session("AdminLevel") < 4 then
		Session("ErrorMsg") = "You do not have permission to visit this page."
		Response.Redirect("http://cozmathematics.1apps.com/main.asp")
	end if
	Dim RNum
	Dim RMemeLink1
	Dim RMemeLink2
	
	Randomize
	
	RNum = int(rnd()*17) + 1
	RMemeLink1 = "memes\" & RNum & ".png"
	RNum = int(rnd()*17) + 1
	RMemeLink2 = "memes\" & RNum & ".png"
%>
<HTML>
	<HEAD>
		<TITLE>Delete User</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			margin: 0 20px 0 20px;
		}
		cite {
			font-size: 15;
			
		}
		
		H3 {
			font-size: 15;
		}
		table{
			width: 100%;
			border: 1px solid black;
		}
		td {
			text-align: center;
			vertical-align: middle;
			height: 150px;
		}
		.boxed{
			border: 1px solid red;
			width: 50%;
			background-color: red;
		}
		
		.toolbarImg{
			margin: 0 25px 0 25px;
		}
		.memeLeft{
			float: left;
			width: 17%;
			height: 20%;
		}
		.memeRight{
			float: right;
			width: 17%;
			height: 20%;
		}
		.logo{
			display: block;
			margin-left: auto;
			margin-right: auto;
		}
		</STYLE>
	</HEAD>
	<BODY style="background-color: 50B2DF;"> 
		<hr size="5" style="color: black; background-color: black;">
		<p>
		<img class="memeLeft" src=<%=RMemeLink1%>><img>
		<img class="memeRight" src=<%=RMemeLink2%>><img>
		<center><a href="http://cozmathematics.1apps.com/main.asp">
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
		
		<table border=1>
		</Center>
		<FONT SIZE = 4>
		<%
		
		Dim User
		dim UserResponse
		User = Session("DeleteUser")
		UserResponse = Request.Form("txtFinalizeDeletion")
		
		if User = "" then
			Session("ErrorMsg") = "No request for deleting a user received. Redirected back to main page."
			Response.Redirect("http://cozmathematics.1apps.com/main.asp")
		end if	
		
		
		If Ucase(UserResponse) = "YES" then
			Response.Write("You request for the deletion of User: " & User & " was received.")
			Const adLockOptimistic = 3
	
			Dim objConn
			Dim strConnection
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\Accounts.mdb")
									
			objConn.Open (strConnection)
			
			Dim strSQL
			strSQL = "Select * FROM Accounts"
			
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open strSQL, objConn, , adlockOptimistic
			
			do while not objRS.EOF
				Dim AUser
				AUser = UCase(objRS("UserName"))
				if UCase(User) = AUser then
					objRS("Admin") = -1
					objRS.Update
					exit do
				end if
				objRS.MoveNext
			loop
			objRS.Close
			Set objRS = Server.CreateObject("ADODB.Recordset")
			strSQL = "DELETE FROM Accounts WHERE Admin ='-1'"
			objRS.Open strSQL, objConn, , adlockOptimistic
			Response.Write("<br>The user has now been completely removed from the database.")
			Response.Write("<br>The account will not be accessible the next time the user attempts to log in.")
		elseif Ucase(UserResponse) = "NO" then
			Response.Write("The Deletion of User: " & User & " has been stopped.")
			Session("DeleteUser") = ""
		else
			Response.Write("No Valid Response Received. The Deletion of User: " & User & " has been stopped to protect the user. Restart the process if this was a mistake.")
			Session("DeleteUser") = ""
		end if
		
		%>
		<br>To return back to the main page. Click <a href="main.asp">Here</a>
		</FONT>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>
