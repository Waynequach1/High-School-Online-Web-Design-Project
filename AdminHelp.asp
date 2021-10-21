<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	elseif Session("AdminLevel") < 2 then
		Session("ErrorMsg") = "You do not have permission to visit this page."
		Response.Redirect("http://cozmathematics.1apps.com/main.asp")
	end if
	
	Dim Admin
	Dim RNum
	Dim RMemeLink1
	Dim RMemeLink2
	
	Randomize
	
	RNum = int(rnd()*17) + 1
	RMemeLink1 = "memes\" & RNum & ".png"
	RNum = int(rnd()*17) + 1
	RMemeLink2 = "memes\" & RNum & ".png"
	Admin = Session("AdminLevel")
%>

<HTML>
	<HEAD>
		<TITLE>Coz Mathematics</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
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
		body{
			margin: 0 20px 0 20px;
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
		
		<font size=5>
		</Center>
		<%
			if Admin = 2 then
				Response.Write("<br>")
				Response.Write("Your admin level is <b>2</b>.")
				Response.Write("<br>")
				Response.Write("You are a Quiz Helper.")
				Response.Write("<br>")
				Response.Write("You have the ability to add new quizzes according to the format described on the create quiz page.")
			elseif Admin = 3 then
				Response.Write("<br>")
				Response.Write("Your admin level is <b>3</b>.")
				Response.Write("<br>")
				Response.Write("You are a Leadboards Admin.")
				Response.Write("<br>")
				Response.Write("You have the ability to remove/change scores on the leaderboards if invalid.")
				Response.Write("<br>")
				Response.Write("You also have the ability to add new quizzes according to the format described on the create quiz page.")
			elseif Admin = 4 then
				Response.Write("<br>")
				Response.Write("Your admin level is <b>4</b>.")
				Response.Write("<br>")
				Response.Write("You are an Adminstrator and of the highest level.")
				Response.Write("<br>")
				Response.Write("You have the ability to change admin levels of all verified users.")
				Response.Write("<br>")
				Response.Write("You additionally have the ability to remove/change scores on the leaderboards if invalid.")
				Response.Write("<br>")
				Response.Write("You also have the ability to add new quizzes according to the format described on the create quiz page.")
			end if
		%>
		
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HEAD>
