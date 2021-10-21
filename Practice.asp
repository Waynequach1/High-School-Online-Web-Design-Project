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
		
		<FONT SIZE=4>
		<BR><B>NOTICE FROM OUR TEAM:</B>
		<br>We strive to deliver the best website for practicing BEDMAS questions.
		<br>However as our team is small it is difficult to consistently check that all answers are correct.
		<br>If the answer to a questions is incorrect, it will be corrected shortly after.
		<br>Thank you for your cooperation.
		<br>
		<H2>Symbols and Their Meaning</H2>
		To make it easier for our team to create questions. We use the following symbols.
		<br><u>Symbol:</u> + is equivalent to Addition.
		<br><u>Symbol:</u> - is equivalent to Subtraction.
		<br><u>Symbol:</u> x is equivalent to multiplication.
		<br><u>Symbol:</u> / is equivalent to division.
		<br><u>Symbol:</u> ^ is equivalent to exponents.
		<br><u>Symbol:</u> () is equivalent to brackets.
		<br>
		<br>
		<FIELDSET>
		<H2>Select one of the following types of BEDMAS questions to practice.</H2>
		<a href="Questions1.asp">Practice 1:</a> Questions with basic addition and subtraction.
		<br><a href="Questions2.asp">Practice 2:</a> Questions with basic multiplication and division.
		<br><a href="Questions3.asp">Practice 3:</a> Questions with mixed operators.
		<br><a href="Questions4.asp">Practice 4:</a> Questions with mixed operators and exponents.
		<br><a href="Questions5.asp">Practice 5:</a> Questions with mixed operators, exponents and brackets.
		</FIELDSET>
		<br>
		<br>
		<FIELDSET>
		<H2>Want to try out one of our quizzes? Click <A HREF="Quiz.asp">Here</A>
		</FIELDSET>
		<br>
		<br>
		</FONT>
		<FONT SIZE=2>
		Can't find what your looking for? Come back again and we will have new topics to study.
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>