<%@ Language=VBscript %>
<% Option Explicit %>

<%
	If Session("Verified") = False then
		Session("ErrorMsg") = "Please Log in to Continue"
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	elseif Session("AdminLevel") < 4 then
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
		<TITLE>Delete User</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			margin: 0 20px 0 20px;
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
		<hr size="5" style="color: black; background-color: black;">
		
		<table border=1>
		</Center>
		<FONT SIZE = 4>
		<%

		Dim User
		User = Request.QueryString(1)
		Response.Write("<b>Note: You can not reverse the effects of deleting a user. Delete Users only if necessary and required.</b>")
		Session("DeleteUser") = User
		Response.Write("<br>")
		Response.Write("<br>")
		Response.Write("<br>To finalize the deletion of User: " & User & " enter <b>Yes</b> into the textbox.")
		Response.Write("<br>To stop the request of deleting User: " & User & " enter <b>No</b> into the textbox.")
		Response.Write("<br>Do <b>NOT</b> press the back button. You must answer Yes or No.")
		Response.Write("<form method=""Post"" action=""DeleteUser2.asp"">")
		Response.Write("<label for=""txtFinalizeDeletion"">Your Response </label>")
		Response.Write("<input type = ""TEXT"" name=""txtFinalizeDeletion"">")
		Response.Write("<input type=""submit"" value=""Submit Response."">")
		Response.Write("</form>")
		%>
		</FONT>
		<HR>
		<CITE>&copy; 2019 Coz Mathematics</CITE>
	</BODY>
</HTML>
