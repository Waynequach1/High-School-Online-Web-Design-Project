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
		<TITLE>New Quiz</TITLE>
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
		
		</Center>
		
		Welcome to Quiz Creation.
		<br>Enter in the questions and answer to the questions.
		<br>To format questions follow this guide:
		<br>Insert a space inbetween all symbols and numbers.
		<br>Always include an equal sign in the question 
		<br>ex: 4 + 5 =
		<% 
			Dim Questions
			Questions = Request.Form("txtNumQuestions")
			Dim Description
			Description = Request.Form("txtDesc")
			
			if Questions = "" then	
				Session("ErrorMsg") = "You cannot create a quiz with no questions. You have been redirected back to the main page."
				Response.Redirect("Main.asp")
			elseif Description = "" then
				Session("ErrorMsg") = "You cannot create a quiz with no description. You have been redirected back to the main page."
				Response.Redirect("Main.asp")
			elseif Questions = 0 then
				Session("ErrorMsg") = "You cannot create a quiz with no questions. You have been redirected back to the main page."
				Response.Redirect("Main.asp")
			end if
			
			Session("Desc") = Description
			Response.Write("<br>")
			Response.Write("<hr>")
			Response.Write("Please remember to follow your quiz description: " & Description)
			Dim X
			Session("CreateNumQuestions") = Questions
			if int(Questions) > 0 then
				Response.Write("<form method=""Post"" action=""AddQuiz.asp"">")
				For X = 1 to Questions
					if x < 10 then
						Response.Write("<br>")
						Response.Write("<label for=""txtQuestion" & x & """> Question " & X & ": &nbsp;&nbsp;&nbsp;&nbsp;" & "</label>")
						Response.Write("<input type = ""TEXT"" name=""txtQuestion" & x & """>")
						Response.Write("<br>")
						Response.Write("<label for=""txtAnswer" & x & """> Answer " & X  & ":&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &"</label>")
						Response.Write("<input type = ""TEXT"" name=""txtAnswer" & x & """>")
						Response.Write("<br>")
					else
						Response.Write("<br>")
						Response.Write("<label for=""txtQuestion" & x & """> Question " & X & ": &nbsp;&nbsp;" & "</label>")
						Response.Write("<input type = ""TEXT"" name=""txtQuestion" & x & """>")
						Response.Write("<br>")
						Response.Write("<label for=""txtAnswer" & x & """> Answer " & X  & ":&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &"</label>")
						Response.Write("<input type = ""TEXT"" name=""txtAnswer" & x & """>")
						Response.Write("<br>")
					end if
				Next
				Response.Write("<br>")
				Response.Write("<input type=""submit"" value=""Add Quiz"">")
				Response.Write("<form>")
			else
				Response.Write("Invalid Input. Please Try Again.")
			end if
			
		%>
	<HR>
	<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>