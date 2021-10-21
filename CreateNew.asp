<html>
	<head>
		<title>Create Account</title>
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
		
		
		
		<font size="5">Already have an account? Click <a href="Login.asp">Here</a> to log in.</font>
		<br>
		<div id="login">
			<form method="post" action="CreateUser.asp">
				<fieldset>
					<%
						if len(Session("ErrorMsg")) > 0 then
							Response.Write("<div class=" & "boxed" & ">")
								Response.Write(Session("ErrorMsg"))
							Response.Write("</div>")
							Session("ErrorMsg") = ""				
						end if
			
					%>
					<h2>Create An Account.</h2>
					<hr>
					<ul>
					<li type="disc">Note: Account names must be unique and are not case sensitive (i.e USERNAME = username)</li>
					</ul>
					<br>
					<label for="txtUserName">Username: </label>
					<input type = "TEXT" name="txtUserName">
					<BR><BR>
					<label for="txtPassword">Password:&nbsp;</label>
					<input type = "password" name="txtPassword">
					<BR><BR>					
					<input type="submit" value="Create Account"/>
					
				</fieldset>
			</form>
		<center><cite> &copy; Coz Mathematics 2019 </cite></center>
</HTML>

