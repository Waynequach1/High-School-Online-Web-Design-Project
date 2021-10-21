<%@ Language=VBscript %>
<% Option Explicit %>

<%
	
	Dim Version
	Version = Session("PracticeVer")
	
	if Version = -1 then
		Response.Redirect("http://cozmathematics.1apps.com/Practice.asp")
	end if
	
	Dim X
	Dim NumQuestions
	Dim objConn
	Dim strConnection
	Set objConn = Server.CreateObject("ADODB.Connection")
	strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
	   Server.MapPath("data\Questions" & Version & ".mdb")
	Dim OpenPath
	OpenPath = "Questions" & Version
	objConn.Open (strConnection)
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open OpenPath, objConn ,,, 2

	Dim Questions(1000)
	Dim Answers(1000)

	Do while not objRS.EOF
		X = X + 1
		Questions(X) = objRS("Question")
		Answers(X) = objRS("Answer")
		objRS.MoveNext
	loop
	NumQuestions = X
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
		<TITLE>Practice Questions</TITLE>
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
		
		<FONT SIZE=4>
		<%
			Dim QuestionNum(11)
			Dim StuAnswer(11)
			Dim Y
			Dim Correct
			Correct = 0
			for X = 1 to 10
				Response.Write("<br>")
				QuestionNum(X) = Request.Form("Question" & Trim(x))
				StuAnswer(X) = Request.Form("txtAnswer" & Trim(X))
				
				for Y = 1 to NumQuestions
					if Trim(QuestionNum(X)) = Trim(Y) then
						if Trim(StuAnswer(X)) = Trim(Answers(Y)) then
							Response.Write("<b>")
							Response.Write("Question " & X & ": " & Questions(QuestionNum(X)) & "<br>")
							Response.Write("</b>")
							Response.Write("<FONT COLOR=YELLOW>")
							Response.Write("Your Answer: " & StuAnswer(X) & " is correct." & "<br>")
							Response.Write("<br>")
							Response.Write("</FONT>")
							Correct = Correct + 1
						else
							Response.Write("<b>")
							Response.Write("Question " & X & ": " & Questions(QuestionNum(X)) & "<br>")
							Response.Write("</b>")
							Response.Write("<FONT COLOR=RED>")
							Response.Write("Your Answer: " & StuAnswer(X) & " is incorrect." & "<br>")
							Response.Write("The Correct Answer is: " & Answers(Y) & "<br>")
							Response.Write("<br>")
							Response.Write("</FONT>")
						end if
					end if
				Next			
			Next
			
			Response.Write("</FONT>")
			Response.Write("<br>")
			Response.Write("<br>")
			Response.Write("<FONT SIZE=5>")
			if Correct >= 8 then
				Response.Write("Nice! You Got " & Correct & "/10 Questions Correct.")
				Response.Write("<br>")
				Response.Write("Think your ready for harder questions? Click <A HREF=""http://cozmathematics.1apps.com/Practice.asp""" & ">Here</a>")
				Response.Write("<br>")
				Response.Write("Need More Practice? Click <A HREF=""http://cozmathematics.1apps.com/Questions" & Version &  ".asp""" & ">Here</a>")
			elseif Correct >= 6 then
				Response.Write("Cool! You Got " & Correct & "/10 Questions Correct.")
				Response.Write("<br>")
				Response.Write("Your almost ready for harder questions! If you think your ready Click <A HREF=""http://cozmathematics.1apps.com/Practice.asp""" & ">Here</a>")
				Response.Write("<br>")
				Response.Write("Want More Practice? Click <A HREF=""http://cozmathematics.1apps.com/Questions" & Version &  ".asp""" & ">Here</a>")
			elseif Correct < 6 then
				Response.Write("Uh Oh! You Got " & Correct & "/10 Questions Correct.")
				Response.Write("<br>")
				Response.Write("You Need More Practice! Click <A HREF=""http://cozmathematics.1apps.com/Questions" & Version &  ".asp""" & ">Here</a> for more practice.")
			end if
			Response.Write("</FONT>")
			Session("PracticeVer") = -1
		%>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>
