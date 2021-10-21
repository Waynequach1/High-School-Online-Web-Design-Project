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
		<TITLE>Quizzes</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
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
		
		</Center>
		
		<FONT SIZE=5>
		Welcome to the quizzes section of practice. 
		<br>Quizzes consist of practice quesitons that have been created by our developing team.
		<br>To start click one of the following quizzes to start.
		<br><br>
		
		<%
			'Shows the availble quizzes or no quizzes if none available.
			Dim NumQuizzes
			Dim objConn
			Dim strConnection
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\Quizzes.mdb")
									
			objConn.Open (strConnection)
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open "Quizzes", objConn ,,, 2
			
			Do While not ObjRS.EOF
				Dim Temporary
				Temporary = objRS("Questions")
				Dim Desc
				Desc = objRS("Description")
				if Trim(Temporary) = "*" then
					NumQuizzes = NumQuizzes + 1
					if NumQuizzes = 1 then
						Response.Write("<form method=""Post"" action=""GenerateQuiz.asp"">")
						Response.Write("<label for=""lstQuiz"">Select A Quiz <br></label>")
						Response.Write("<select name=""lstQuiz"">")
					end if
					Response.Write("<option value =" & NumQuizzes & ">" & "Quiz Number " & NumQuizzes & ": " & Desc & "</option>")
				end if
				objRS.MoveNext
			Loop
			
			if NumQuizzes < 1 then
				Response.Write("<br><br><b>Unfortuantely there are no quizzes at the moment.<br>")
				Response.Write("Please come back another time when we have added some quizzes.</b>")
			else
				Response.Write("<input type=""submit"" value=""Continue"">")
				Response.Write("<form>")
			end if
		%>
		
		</Font>
		<%
			'Checks if the user is admin and is allowed to visit a page to add quizzes.
			if Session("AdminLevel") > 1 then
				Response.Write("<br><br><b>ADMIN:</b> You can add new quizzes. Click <a href=""GenerateQuestion.asp"">Here</a> to add new quizzes.")
			end if
		%>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>