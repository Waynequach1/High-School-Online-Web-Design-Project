<%@ Language=VBscript %>
<% Option Explicit %>

<html>
	<head>
		<title>Login To CM</title>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>

		table{
			width: 100%;
			border: 1px solid black;
		}
		td {
			text-align: center;
			vertical-align: middle;
			height: 150px;
		}
		form{
			max-width: 500px;
			background-color: white;
		}
		fieldset{
			border: 0px;
		}
		label{
			font-size: 12pt;
		}
		
		input{
			max-width: 300px;
			max-height: 80px;
			font-size: 10pt;
		}
		body{
			font-family: verdana;
			background-color: 50B2DF;
			margin: 0 20px 0 20px;
		}
		.boxed{
			border: 1px solid red;
			width: 100%;
			background-color: red;
		}
		
		</STYLE>
	</head>
	<body>
		<center><img src="Images\CMLogo.png" width="40%"></img>
		<br>
		
		
		
		<font size="5">Don't have an account? <a href="CreateNew.asp"> SIGN UP ITS FREE!</a></font>
		<br>
		<div id="login">
			<form method="post" action="verifyUser.asp">
				<fieldset>
					
					<%
						if len(Session("ErrorMsg")) > 0 then
							Response.Write("<div class=" & "boxed" & ">")
								Response.Write(Session("ErrorMsg"))
							Response.Write("</div>")
							Session("ErrorMsg") = ""				
						end if
			
					%>
					
					<h2>Please log in</h2>
					<hr>
                                        <BR>
					<label for="txtUserName">Username: </label>
					<input type = "TEXT" name="txtUserName">
					<BR><BR>
					<label for="txtPassword">Password:&nbsp;</label>
					<input type = "password" name="txtPassword">
					<BR><BR>					
					<input type="submit" value="Login"/>
					
				</fieldset>
			</form>
		</div></center>
	<center><cite> &copy; Coz Mathematics 2019 </cite></center>
</html>
