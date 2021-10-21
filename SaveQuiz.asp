<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	elseif Session("AdminLevel") < 2 then
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
		<TITLE>Edit Quiz</TITLE>
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
		
		<%
			Const adLockOptimistic = 3
	
			Dim objConn
			Dim strConnection
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\Quizzes.mdb")
									
			objConn.Open (strConnection)
			
			Dim strSQL
			strSQL = "SELECT * FROM Quizzes"
			
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open strSQL, objConn, , adlockOptimistic
			
			Dim QuizVersion
			QuizVersion = Session("QuizVersion")
			
			if QuizVersion = 0 then
				Session("ErrorMsg") = "No Quiz Selected. Redirected back to main page."
				Response.Redirect("main.asp")
			end if
			
			Dim CurrentQuiz
			
			Do While not objRS.EOF
				'Loops until it finds the correct quiz version.
				if objRS("Questions") = "*" then
					CurrentQuiz = CurrentQuiz + 1
				end if
				
				if Trim(CurrentQuiz) = Trim(QuizVersion) then
					exit do
				end if
				objRS.MoveNext
			loop
			Dim NewDesc
			NewDesc = Request.Form("txtDesc")
			objRS("Description") = NewDesc
			objRS.MoveNext
			
			Dim X
			Dim Question
			Dim Answer
			X = 0
			Do While not objRS.EOF
				if Trim(objRS("Questions")) = "*" then
					'if star is reached means new quiz means exit
					exit do
				end if
				
				x = x + 1
				Question = Request.Form("txtQuestion" & Trim(x))
				Answer = Request.Form("txtAnswer" & Trim(x))
				objRS("Questions") = Question
				objRS("Answers") = Answer
				objRS.MoveNext
			loop
			
		%>
		<FONT SIZE=5>
		The following questions and answers have been updated. Click <a href="quiz.asp">Here</a> to return to the quiz section.
		</FONT>
	<HR>
	<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>