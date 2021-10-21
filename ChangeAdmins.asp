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
		<TITLE>Change Admin Level/Delete User</TITLE>
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
		<CENTER>
		<FONT SIZE=5>
		Change Admin Levels
		</FONT>
		<%
		Dim objConn
		Dim strConnection
		Set objConn = Server.CreateObject("ADODB.Connection")
		strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
		   Server.MapPath("data\accounts.mdb")
								
		objConn.Open (strConnection)
		
		Dim strSQL
		strSQL = "SELECT * FROM Accounts"
		strSQL = strSQL & " ORDER BY Admin DESC"
		
		Dim objRS
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open strSQL, objConn
		
		Response.Write("<tr>")
		Response.Write("<td>Account</td>")
		Response.Write("<td>Admin Level</td>")
		Response.Write("<td>Change Admins</td>")
		Response.Write("<td>Delete User</td>")
		Response.Write("</tr>")
		
		Dim X
		X = 0
		Do While not objRS.EOF
			X = x + 1
			Response.Write("<tr>")
			Response.Write("<td>" & objRS("UserName") & "</td>")
			Response.Write("<td>" & objRS("Admin") & "</td>")
			Response.Write("<td> <A HREF=""http://cozmathematics.1apps.com/ChangeAdmins2.asp?txtChange=" & objRS("UserName") & """>Change</a>" & "</td>")
			Response.Write("<td> <A HREF=""http://cozmathematics.1apps.com/DeleteUser.asp?txtChange=" & objRS("UserName") & """>Delete</a>" & "</td>")
			Response.Write("</tr>")
			objRS.MoveNext
		loop

		
		%>
		</table>
		</CENTER>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>