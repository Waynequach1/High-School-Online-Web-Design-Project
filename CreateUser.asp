<%@ Language=VBscript %>
<% Option Explicit %>
	
<%
	Const adLockOptimistic = 3
	Dim AUsername
	Dim Password
	Dim ErrorMsg
	Dim Unique
	
	Unique = true
	AUsername = Request.Form("txtUserName")
	Password = Request.Form("txtPassword")
	
	Dim objConn
	Dim strConnection
	Set objConn = Server.CreateObject("ADODB.Connection")
	strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
	   Server.MapPath("data\accounts.mdb")
							
	objConn.Open (strConnection)
	
	Dim strSQL
	strSQL = "SELECT * FROM Accounts"
	
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open "Accounts", objConn ,,, 2
	
	Dim AUser
	Do while not objRS.EOF
		AUser = Trim(objRS("UserName"))
		If Ucase(AUser) = Ucase(AUsername) then
			ErrorMsg = "The Username has already been taken. Please Enter a new username."
			Unique = false
			exit do
		End if
		objRS.MoveNext
	Loop
	
	if Trim(AUsername) = "" then
		ErrorMsg = "The Username cannot be blank."
		Unique = false
	elseif Trim(Password) = "" then
		ErrorMsg = "The Password cannot be blank."
		Unique = false
	end if
	
	
	if Unique = true then
		objRS.Close
		objConn.Close
		set objConn = nothing
		set objRS = nothing
		set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open (strConnection)
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open strSQL, objConn, , adlockOptimistic
		
		objRS.AddNew
		objRS("UserName") = AUsername
		objRS("Password") = Password
		objRS("Rank") = -1
		objRS("Admin") = 1
		objRS.Update
	elseif Unique = false then
		Session("ErrorMsg") = ErrorMsg
		Response.Redirect("http://cozmathematics.1apps.com/createaccount.asp")
	End if
	
%>

<HTML>
	<HEAD>
		<TITLE>New Account</TITLE>
		<LINK REL="shortcut icon" HREF="Images\Logo.ico">
		<STYLE>
		body{
			font-family: verdana;
			background-color: 50B2DF;
			font-size: 20pt;
			margin: 0 20px 0 20px;
		}
		</STYLE>
	</HEAD>
	<BODY>
		<center><img src="Images\CMLogo.png" width="40%"></img></center>
		<BR><BR><BR>
		<center>The account: <B><U><font color="red"><%=AUsername%></font></U></B> has been created sucessfully. To log in Click <a href="http://cozmathematics.1apps.com/login.asp">here</a>.</center>
	</BODY>
</HTML>