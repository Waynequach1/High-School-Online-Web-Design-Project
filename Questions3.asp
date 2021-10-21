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
		
		</Center>
		
		<H2>Practice 3: Questions with mixed operators.</H2>
		<br>
		<H3>Try and answer the following questions:</H3>
		<br> <H4>Note: This is not a test, take your time and be as precise as possible</H4>
		<form method="post" action="CheckAnswers1.asp">
			<fieldset>
				<ul>
					<%
						Session("PracticeVer") = 3
						Dim X
						Dim NumQuestions
						
						Dim objConn
						Dim strConnection
						Set objConn = Server.CreateObject("ADODB.Connection")
						strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
						   Server.MapPath("data\Questions3.mdb")
												
						objConn.Open (strConnection)
						Dim objRS
						Set objRS = Server.CreateObject("ADODB.Recordset")
						objRS.Open "Questions3", objConn ,,, 2
						
						Dim Questions(1000)
						Dim Answers(1000)
						
						Do while not objRS.EOF
							x = x + 1
							Questions(X) = objRS("Question")
							Answers(X) = objRS("Answer")
							objRS.MoveNext
						loop
						NumQuestions = X
						
					%>
					<FONT SIZE =5 face="courier">
					<%
						Dim Random
						Randomize
						For X = 1 to 10
							do
							Random = int(rnd() * NumQuestions ) + 1
							loop while Questions(Random) = "" 
							Dim Sentence1
							Dim Sentence2
							if x < 10 then
								Sentence1 = "<label for=""txtQuestion" & x & """>" & "Question " & X & ":&nbsp;&nbsp;" & Trim(Questions(Random)) & "</label>"
							else
								Sentence1 = "<label for=""txtQuestion" & x & """>" & "Question " & X & ":&nbsp;" & Trim(Questions(Random)) & "</label>"
							end if
							if x < 10 then
								do while len(Sentence2) < (88 - len(Sentence1)) * 6
									Sentence2 = Sentence2 & "&nbsp;"
								loop
							else
								do while len(Sentence2) < (84 - len(Sentence1)) * 6
									Sentence2 = Sentence2 & "&nbsp;"
								loop
							end if
							Sentence2 = Sentence2 & "<input type = ""TEXT&"" name=""txtAnswer" & X & """>"
							Response.Write(Sentence1)
							Response.Write(Sentence2)
							Response.Write("<br>")
							Questions(Random) = "" 
							Sentence2 = ""
							Response.Write("<input type = ""Hidden"" name = ""Question" & x & """ value = """ & Random & """></input>")
						Next
					%>
					<br>
					
				<input type="submit" value="Submit Answers" />
				</FONT>
				</ul>
			</fieldset>
		</form>
		
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>