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
		<form method="post" action="CheckAnswers2.asp">
		<FONT SIZE =5 face="courier">
		<FIELDSET>
		
		<%
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
			objRS.Open "Quizzes", objConn ,,, 2
			
			Dim QuizVersion
			QuizVersion = Session("QuizVersion")
			
			Dim CurrentQuiz
			CurrentQuiz = 0
			if QuizVersion <> 0 then
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
				
				Response.Write("Current Quiz: " & ObjRS("Description"))
				Response.Write("<br>")
				Response.Write("<br>")
				objRS.MoveNext
				
				Dim Questions(1000)
				Dim Answers(1000)
				Dim X
				
				Do While not objRS.EOF
					if Trim(objRS("Questions")) = "*" then
						'if star is reached means new quiz means exit
						exit do
					else
						X = X + 1
						Questions(X) = objRS("Questions")
						Answers(X) = objRS("Answers")
						objRS.MoveNext
					end if
				loop
				
				
				Dim QuestionNum(1000)
				Dim StuAnswer(1000)
				Dim Correct
				Correct = 0
				Dim NumQuestions
				NumQuestions = X
				
				For X = 1 to NumQuestions
					StuAnswer(X) = Request.Form("txtAnswer" & Trim(x))
					if StuAnswer(X) = Answers(X) then
						Response.Write("Question  " & x & ": " & Questions(X))
						Response.Write("<FONT COLOR=YELLOW>")
						Response.Write("<br>")
						Response.Write("Your Answer: " & StuAnswer(X))
						Response.Write(" is correct.")
						Response.Write("</FONT>")
						Response.Write("<br>")
						correct = correct + 1
					else
						Response.Write("Question " & x & ": " & Questions(X))
						Response.Write("<br>")
						Response.Write("<FONT COLOR=RED>")
						Response.Write("Your Answer: " & StuAnswer(X))
						Response.Write(" is incorrect.")
						Response.Write("<br>")
						Response.Write("The correct answer is: " & Answers(X))
						Response.Write("</FONT>")
						Response.Write("<br>")
					end if
					Response.Write("<br>")
				Next
			else
				Session("ErrorMsg") = "No Quiz was selected. Redirected back to main page."
				Response.Redirect("main.asp")	
			end if
			
			Response.Write("You got " & Correct & " out of " & NumQuestions & " questions correct.")
			Response.Write("<br>")
			if Correct / NumQuestions < 0.6 then
				Response.Write("Oh no! You should get more practice before trying this quiz again.")
				Response.Write("<br>")
				Response.Write("Click <A HREF=""http://cozmathematics.1apps.com/practice.asp"">Here</a> to get more practice.")
			elseif Correct / NumQuestions < 0.8 then
				Response.Write("Nice! Keep practicing to get a perfect score.")
				Response.Write("<br>")
				Response.Write("Click <A HREF=""http://cozmathematics.1apps.com/practice.asp"">Here</a> to get more practice.")
			else
				Response.Write("Excellent! Maybe you should move on to harder quizzes.")	
				Response.Write("<br>")
				Response.Write("Click <A HREF=""http://cozmathematics.1apps.com/practice.asp"">Here</a> to get more practice.")				
			end if
		
		%>