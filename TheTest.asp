<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	end if	
	
	Dim Remaining
	Dim SecondsLeft
	Dim MinutesLeft
	Dim CurTime
	CurTime = Timer
	Remaining = Int(Session("TimeLeft") - CurTime)
	
	SecondsLeft = Remaining
%>
<HTML>
	<HEAD>
		<TITLE>CM Test</TITLE>
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
		<FONT SIZE=5>
		</Center>
		<form method="post" action="TheTest.asp">
		<%
			if SecondsLeft <= 0 then
				If Session("TestScore") > 0 then
					Response.Redirect("TestResults.asp")
				else
					Session("TimeLeft") = Timer + 60
					Remaining = Int(Session("TimeLeft") - CurTime)
					SecondsLeft = Remaining
					Session("TestScore") = 0
				end if
			else
				if Request("txtAnswer") = Session("TestA") then
					Session("TestScore") = Session("TestScore") + Session("Points")
					Response.Write("<FONT COLOR=YELLOW>You got that question correct! You got " & Session("Points") & " points.<br></FONT>")
				elseif Trim(Request("txtAnswer")) = "" then
					Response.Write("<FONT COLOR=BLACK>You skipped the question.")
				else
					Response.Write("<FONT COLOR=RED>Your answer was incorrect. The Correct Answer was " & Session("TestA") & "<br></FONT>")
				end if
			end if
		%>
		<%
			CONST TYPEQUESTIONS = 5
			Dim X
			Dim NumQuestions
			
			Dim objConn
			Dim strConnection
			Dim QType
			Dim File
			
			File = "Questions"
			Randomize
			QType = int(rnd() * TYPEQUESTIONS ) + 1
			Session("Points") = QType
			File = File & Trim(QType)
			Response.Write("<FONT SIZE=5>")
			Response.Write("<CENTER>This question is worth: " & QType & " points.")
			Response.Write("<br><br>")
			
			Set objConn = Server.CreateObject("ADODB.Connection")
			strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
			   Server.MapPath("data\" & File & ".mdb")
									
			objConn.Open (strConnection)
			
			Dim objRS
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open File, objConn ,,, 2
			
			Dim Questions(1000)
			Dim Answers(1000)
			
			Do while not objRS.EOF
				x = x + 1
				Questions(X) = objRS("Question")
				Answers(X) = objRS("Answer")
				objRS.MoveNext
			loop
			NumQuestions = X
			
			Dim Random
			Dim Sentence1
			Dim Sentence2
			
			Random = int(rnd() * NumQuestions ) + 1
			Session("TestA") = Answers(Random)
			
			Sentence1 = "<label for=""txtQuestion"" >Question: &nbsp;&nbsp;" & Trim(Questions(Random)) & "</label>"
			
			if QType = 1 then
				do while len(Sentence2) < (74 - len(Sentence1)) * 6
					Sentence2 = Sentence2 & "&nbsp;"
				loop
			elseif QType = 2 then
				do while len(Sentence2) < (78 - len(Sentence1)) * 6
					Sentence2 = Sentence2 & "&nbsp;"
				loop
			elseif QType = 3 then
				do while len(Sentence2) < (82 - len(Sentence1)) * 6
					Sentence2 = Sentence2 & "&nbsp;"
				loop
			elseif QType = 4 then
				do while len(Sentence2) < (86 - len(Sentence1)) * 6
					Sentence2 = Sentence2 & "&nbsp;"
				loop
			elseif QType = 5 then
				do while len(Sentence2) < (90 - len(Sentence1)) * 6
					Sentence2 = Sentence2 & "&nbsp;"
				loop
			end if
			
			Sentence2 = Sentence2 & "<input autofocus type = ""TEXT"" name=""txtAnswer"">"
			Response.Write(Sentence1)
			Response.Write(Sentence2)

		%>
		</FONT>
		<br><br><br>
		<input type="submit" value="Submit" />
		<HR COLOR=BLACK>
		<iFRAME frameborder=1 width="40%" name="main" src="DisplayTime.asp"></iFRAME>
		<br>
		Coz Mathematics Test Version 1.0.0 ALPHA
		</CENTER>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>