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
		<TITLE>Generate Question</TITLE>
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
		
		Welcome to Quiz Creation. Enter in the number of questions that are in your quiz.
		<br>Sample Description: Easy Addition
		<br><br>
		<% 
			Response.Write("<form method=""Post"" action=""NewQuiz.asp"">")
			Response.Write("<label for=""txtNumQuestions"">Number Of Questions: </label>")
			Response.Write("<input type = ""TEXT"" name=""txtNumQuestions"">")
			Response.Write("<br>")
			Response.Write("<label for=""txtDesc"">Description of Quiz Type: </label>")
			Response.Write("<input type = ""TEXT"" name=""txtDesc"">")
			Response.Write("<br><br>")
			Response.Write("<input type=""submit"" value=""Generate Template"">")
			Response.Write("</form>")
			Response.Write("<br>")
			Response.Write("<br>")
			Response.Write("<hr color=black>")
			Response.Write("<hr color=black>")
			Response.Write("Edit an existing Quiz. It is not possible to change the number of questions in a quiz.")
			Response.Write("<br>")
			Response.Write("<br>")
			
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
						Response.Write("<form method=""Post"" action=""EditQuiz.asp"">")
						Response.Write("<label for=""lstQuiz"">Select A Quiz <br></label>")
						Response.Write("<select name=""lstQuiz"">")
					end if
					Response.Write("<option value =" & NumQuizzes & ">" & "Quiz Number " & NumQuizzes & ": " & Desc & "</option>")
				end if
				objRS.MoveNext
			Loop
			
			
			if NumQuizzes < 1 then
				Response.Write("There are no quizzes at the moment. Add some first to edit one.</b>")
			else
				Response.Write("<input type=""submit"" value=""EDIT QUIZ"">")
				Response.Write("</form>")
			end if
		%>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>