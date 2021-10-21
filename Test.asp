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
	Session("TestScore") = 0
%>
<HTML>
	<HEAD>
		<TITLE>Test</TITLE>
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
		</Center>
		<H2>Test your skills in our test. Compete to achieve the rank of number one on our leaderboards!</H2>
		<br>
		<FIELDSET>
		<H3>Not ready to take on our test? Practice some questions <a href="http://cozmathematics.1apps.com/practice.asp">Here</a></H3>
		</FIELDSET>
		<br><B><U>INSTRUCTIONS:</B></U>
		<br>
		The test will consist of every type of Bedmas question so be prepared. 
		<br>You will be given a minute to answer as many questions as possible correctly.
		<br>Afterwards you will be scored and compared to every other player who has taken the test.
		<br>Points are rewarded based on the difficulty of the question.
		<br>You can skip a question by leaving the answer blank. Sometimes it is best to attempt questions you know best.
		<br>Pressing Back during the test may ruin your results and incorrectly display the questions. Please do not hit the back button.
		<br>Visiting the test page automatically resets your score to prevent cheating. Please do not leave the test.
		<br>Getting a score of zero will not send you to the results page and instead start the test again.
		<br><b>Note: Leaving in the middle of a test will not stop the test. The test will end ONLY after 60 seconds have passed.</b>
		</FONT>
		<br><br>
		<FIELDSET>
		<FONT SIZE=5>
			<br>The Test has been released. 
			Click <A HREF="TheTest.asp">Here</a> To take the test.
		<br>
		</FIELDSET>
		<br>
		</FONT>
		<FONT SIZE=3>
		<br>Note: The test is currently in development and is still in alpha.
		New test questions will be added in the future come back from time to time and test your skills.
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>