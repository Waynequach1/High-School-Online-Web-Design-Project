<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
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
		<TITLE>Leaderboards</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			margin: 0 20px 0 20px;
		}
		.toolbarImg{
			margin: 0 25px 0 25px;
		}
		table{
			width: 80%;
			border: 1px solid black;
			background-color:green;
		}
		td {
			text-align: center;
			vertical-align: middle;
			height: 150px;
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
		fieldset{
			
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
		
		</Center>
		<CENTER>
		<H2>Welcome to the leaderboards.</H2>
		<FIELDSET>
		<%
			Dim Rank
			Dim objConn
			Dim strConnection
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\accounts.mdb")
									
			objConn.Open (strConnection)
			
			Dim strSQL
			strSQL = "SELECT * FROM Accounts"

			
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open strSQL, objConn
			
			if Session("Rank") > 0 then
				Response.Write("Your current score on the leaderboards is: " & Session("Rank"))
			else
				Response.Write("You do not have a rank yet. Take the test under tests in the header to obtain a rank.")
			end if
			
			Response.Write("<TABLE BORDER=1>")
			Response.Write("<tr>")
			Response.Write("<td>Rank</td>")
			Response.Write("<td>Player Name</td>")
			Response.Write("<td>Score</td>")
			if Session("AdminLevel") > 3 then
				Response.Write("<td>Change Admins</td>")
			end if
			Response.Write("</tr>")
			
			Dim X
			X = 0
			Rank = 0
			Dim AllRanks(10000)
			Dim AllNames(10000)
			Dim NumRanks
			
			Do While not objRS.EOF
				if Int(objRS("Rank")) > 0 then
					NumRanks = NumRanks + 1
					AllRanks(NumRanks) = objRS("Rank")
					AllNames(NumRanks) = objRS("UserName")
				end if
				objRS.MoveNext
			loop
			
			Dim Sentinel
			Sentinel =  false
			If NumRanks > 0 then
				Do while Sentinel = false
					For X = 1 to NumRanks
						If Int(AllRanks(X)) < Int(AllRanks(X + 1)) then
							Dim TempRank
							Dim TempName
							
							TempRank = AllRanks(X)
							TempName = AllNames(X)
							AllRanks(X) = AllRanks(X + 1)
							AllNames(X) = AllNames(X + 1)
							AllRanks(X + 1) = TempRank
							AllNames(X + 1) = TempName

						end if
					Next
					
					For X = 1 to NumRanks
						Sentinel = true
						
						if Int(AllRanks(X)) < Int(AllRanks(X + 1)) then
							Sentinel = false
							exit for
						end if
					Next
				Loop
			end if
			
			if NumRanks > 10 then
				NumRanks = 10
			end if
			
			For X = 1 to NumRanks
				Rank = Rank + 1
				Response.Write("<tr>")
				if X > 1 then
					Dim GiveRank
					Dim Counter
					GiveRank = Rank
					Counter = 1
					do while Int(AllRanks(X)) = Int(AllRanks(X-Counter))
						GiveRank = GiveRank - 1
						Counter = Counter + 1
					loop
					response.Write("<td>" & GiveRank & "</td>")
				else
					response.Write("<td>" & Rank & "</td>")
				end if
				Response.Write("<td>" & AllNames(X) & "</td>")
				Response.Write("<td>" & AllRanks(X) & "</td>")
				
				if Session("AdminLevel") > 3 then
					Response.Write("<td> <A HREF=""http://cozmathematics.1apps.com/ChangeScore.asp?txtChange=" & AllNames(X) & """>Change</a>" & "</td>")
				end if
				Response.Write("</tr>")
			Next

		%>
		</TABLE>
		</FIELDSET>
		<HR>
		</CENTER>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>