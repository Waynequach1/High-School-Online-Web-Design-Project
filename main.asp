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
	
	RNum = int(rnd()*45) + 1
	RMemeLink1 = "memes\" & RNum & ".png"
	RNum = int(rnd()*45) + 1
	RMemeLink2 = "memes\" & RNum & ".png"
	
%>
<HTML>
	<HEAD>
		<TITLE>Coz Mathematics</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			margin: 0 20px 0 20px;
			font-family: "Verdana"
		}
		cite {
			font-size: 15;
			
		}
		
		H3 {
			font-size: 15;
		}
		table{
			width: 100%;
			border: 1px solid black;
		}
		td {
			text-align: center;
			vertical-align: middle;
			height: 150px;
		}
		.boxed{
			border: 1px solid red;
			width: 50%;
			background-color: red;
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
		<br>
		<hr size="5" style="color: black; background-color: black;">
		
		
		</Center>
		<center>
			<%
				if Len(Session("ErrorMsg")) > 0 then
					Response.Write("<div class=" & "boxed" & ">")
					Response.Write(Session("ErrorMsg"))
					Session("ErrorMsg") = ""
					Response.Write("</div>")
				end if
			%>
		</center>
		
		<P>Welcome <%=Session("Username")%> to Coz Mathematics! Please select a topic of BEDMAS to study from the toolbar above!</P>
		
		<FONT SIZE = 5>
		<P>Hello Verified Users! Unverified users may no longer access this site.</P>
		<br>Welcome! Coz Mathematics has officially launched. Come practice your skills in the practice zone or take on some challenging Quizzes!
		<br>Or do you to take on other users and compete for the top position?
		<H3>Be a responsible user. Please do not share your password with anyone else!</H3>
		<br><img src="Images\Home_SYM.png" float=left><b>Home</b> Click home to return back to this page.<br>
		<br><img src="Images\Practice_SYM.png" float=left><b>Practice/Quizzes</b> Click practice to practice your skills in BEDMAS, take it easy here and hone your skills.<br>
		<br><br><img src="Images\Test.png" float=left><b>Test</b> Think you've practiced enough? Take your skills here and put it to the test, race against time to
		show what you've learned.<br>
		<br><img src="Images\Leaderboard_SYM.png" float=left><b>Leaderboard</b> Think you've done it all? Complete more tests to reach a higher ranking and aim for the top!<br>
		<br><img src="Images\Logout_SYM.png" float=left><b>Log Out</b> Log out and prevent others from accessing your account.<br>
		
		<HR>
		</FONT>
		<font size=4>
		
		<%
			if Session("AdminLevel") = 4 then
				Response.Write("<FIELDSET>")
				Response.Write("<P><B>You are an administrator. Your Admin level is " & Session("AdminLevel") & ". For more details about your admin powers. Click <A HREF=""http://cozmathematics.1apps.com/AdminHelp.asp""" & ">Here</a></B></P>")
				Response.Write("</FIELDSET>")
			elseif Session("AdminLevel") = 3 then
				Response.Write("<FIELDSET>")
				Response.Write("<P><B>You are a Leadboards Admin. Your Admin level is " & Session("AdminLevel") & ". For more details about your admin powers. Click <A HREF=""http://cozmathematics.1apps.com/AdminHelp.asp""" & ">Here</a></B></P>")
				Response.Write("</FIELDSET>")
			elseif Session("AdminLevel") = 2 then
				Response.Write("<FIELDSET>")
				Response.Write("<P><B>You are a Quiz Helper. Your Admin level is " & Session("AdminLevel") & ". For more details about your admin powers. Click <A HREF=""http://cozmathematics.1apps.com/AdminHelp.asp""" & ">Here</a></B></P>")
				Response.Write("</FIELDSET>")
			end if
			
			If Session("AdminLevel") = 4 then
				Response.Write("<FIELDSET>")
				Response.Write("<b>Administrator " & Session("Username") & " you have the ability to change admin levels/delete users. Click <A HREF=""http://cozmathematics.1apps.com/ChangeAdmins.asp""" & ">Here</a></b>")
				Response.Write("</FIELDSET>")
			end if
		%>
		<font size=4>
			<p>Changelog:</p>
			<br>6/02/2019: Changed up main page to display up to 45 different memes from 17.
			<br>6/02/2019: Fixed leaderboards from infinitely looping due to trying to reach 0 from 1 when no accounts have a proper rank.
			<br>6/02/2019: Changed the user creation to prevent the creation of accounts with empty spaces.
			<br>6/02/2019: Added two step verification to deleting a user.
			<br>6/02/2019: Deleting of accounts now work correctly and only delete the user selected.
			<br>6/02/2019: All users account details have been rerolled due to the testing of deleting records resulting in the deletion of all accounts.
			<br>6/02/2019: Added pages to delete users.
			<br>6/02/2019: Added new column to admin table.
			<br>6/01/2019: More memes added into the database to decrease the chances of seeing the same memes.
			<br>6/01/2019: Fixed some more database questions from displaying incorrect answers.
			<br>6/01/2019: Added a few more quizzes onto quiz page.
			<br>6/01/2019: Fixed the test results page from changing your score if your score was better not worse.
			<br>6/01/2019: Fixed Leaderboards only sorting scores once resulting in scores above 100 to not be listed correctly.
			<br>6/01/2019: Some Major bugs fixed.
			<br>5/31/2019: Leaderboards now only selects the top 10 players instead of all users.
			<br>5/31/2019: Leaderboard now gives users with the same score the same rank.
			<br>5/31/2019: Some optimization to prevent abonormal changes from adding new test.
			<br>5/31/2019: Leaderboard admins and Adiministrators can now change user scores on leaderboards.
			<br>5/31/2019: Leaderboards now correctly sort the ranks.
			<br>5/31/2019: Tests now take your best score instead of your most recent score.
			<br>5/30/2019: Changed up some text for better clarification.
			<br>5/30/2019: Added more detailed explanation to the practice page.
			<br>5/30/2019: Added more detailed explanation to the test page.
			<br>5/30/2019: Changed links with a fieldset to make them clearer to see.
			<br>5/30/2019: Randomized math memes added to headers.
			<br>5/30/2019: Leaderboards ranks players based on a string basis (needs changing).
			<br>5/30/2019: Leaderboards page added.
			<br>5/30/2019: Tests now save users scores and rank them.
			<br>5/29/2019: Test questions are marked and scored.
			<br>5/29/2019: Test framework completed. 
			<br>5/29/2019: Test page now links to actual test.
			<br>5/29/2019: Changed display for other site helpers to clarify their status.
			<br>5/29/2019: Administrators can now change the level of ALL users including themselves.
			<br>5/29/2019: Added new pages for changing admin levels.
			<br>5/28/2019: Added page to change admin ranks.
			<br>5/28/2019: Main page edited slightly to be more helpful to the user.
			<br>5/28/2019: Minor changes to quiz creation to prevent crashing of pages.
			<br>5/28/2019: Databases have been checked again and incorrect answers have been corrected.
			<br>5/28/2019: Changes to certain titles to fix displaying of incorrect titles.
			<br>5/28/2019: Some minior changes to fix existing pages displaying incorrectly.
			<br>5/28/2019: You can now highlight over the header to find out what the images represent.
			<br>5/28/2019: Quiz answers are now marked correctly.
			<br>5/28/2019: Users can now attempt to solve quizzs.
			<br>5/28/2019: Admins can now edit existing quizzes.
			<br>5/28/2019: Admins can now add new quizzes.
			<br>5/26/2019: Quizzes now have descriptions.
			<br>5/26/2019: Added help on quiz pages for creating quizzes.
			<br>5/26/2019: Added administrator help page.
			<br>5/26/2019: Added new database for adding quizzes.
			<br>5/24/2019: Added Admin quiz pages.
			<br>5/24/2019: Added Quiz Page.
			<br>5/17/2019: Added some description and help to administrators.
			<br>5/17/2019: Fixed attempts display after being locked out from saying 0 to 1.
			<br>5/17/2019: Added new logo on all page titles. 
			<br>5/15/2019: All accounts have been reset! Accounts now come with ranks and admin levels.
			<br>5/15/2019: Users are now locked out for one minute when they attempt to log in incorrectly 3 or more times.
			<br>5/15/2019: Users can no longer sign up without a username or password.
			<br>5/15/2019: Error messages are now displayed with a style.
			<br>5/15/2019: Login and User account creation style have been updated.
			<br>5/15/2019: All Login pages now required correct verification to view.
			<br>5/13/2019: Main header finalized and changed on all pages.
			<br>5/13/2019: Changes to some pages to redirect instead of transfer.
			<br>5/13/2019: Sign in now correctly checks that usernames are case-insensitive.
			<br>5/13/2019: Practice question correctly marks practice questions.
			<br>5/12/2019: Marking of practice question asp added.
			<br>5/12/2019: Practice pages display questions correctly.
			<br>5/12/2019: New practice questions added.
			<br>5/12/2019: Practice Question pages added.
			<br>5/11/2019: Leaderboards page added.
			<br>5/11/2019: Working image links added.
			<br>5/11/2019: Quiz/Test Page added.
			<br>5/11/2019: Pratice page added.
			<br>5/9/2019: Main Page header added. 
			<br>5/9/2019: Working Sign up Page added.
			<br>5/9/2019: Working Login Page added.
			<br>5/7/2019: Sign up page added.
			<br>5/7/2019: Main landing page added.
			<br>5/7/2019: Temporary Login Page was created.
			<br>5/7/2019: Site was Created.
			
		<br><br><b>Current Goals: (In no particular order)</b>
			<br>No Goals Currently.
			<br>
			<br> <b>Goals In Progress:</b>
			<br>No Goals Currently.
			<br>
			<br> <b>Goals Completed:</b>
			<br> Allow Level 3+ admins to reset/delete users from leaderboards ~ Done
			<br> Add more questions to databases~ Done
			<br> Randomly select test questions and save results ~ Done
			<br> Test Page Works  ~ Done
			<br> Add changing math memes on the side of header ~ Done
			<br> Test questions are marked correctly ~ Done
			<br> Allow Level 4 admins to change adminerstrator levels of others ~ Done
			<br> Allow users to log out ~ Done
			<br> Quiz Page Works  ~ Done
			<br> Quiz answers are marked correctly ~ Done
			<br> Allow Level 2+ admins to edit/delete quizzes ~ Done
			<br> Allow Level 2+ admins to create quizzes ~ Done
			<br> All pages titles consist of logo ~ Done
			<br> Users are required to be verified to visit pages ~ Done
			<br> Login and Sign up pages style changes ~Done
			<br> Create Home Page ~ Done
			<br> Marking of practice questions ~ Done
			<br> Pratice Page Works ~ Done
			<br> Randomly select math questions from database for practice ~ Done
			<br> Create Leaderboards Page ~ Done
			<br> Create Practice page ~ Done
			<br> Create Quiz/Test Page ~ Done
			<br> Create Log in Page ~ Done
			<br> Create Sign in Page ~ Done
			<br> Users Can Log In ~ Done
			<br> New Users Can Sign In ~ Done
		</font>
			<BR><BR><BR><BR>
		
		</font>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
		
	</BODY>
</HTML>
