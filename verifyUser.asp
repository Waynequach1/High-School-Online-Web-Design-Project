<%@ Language=VBscript %>
<% Option Explicit %>

<%
	Dim Username, Password
	Dim ErrorMsg, Verified, Attempts, Lockout
	
	Username = Request("txtUserName")
	Password = Request("txtPassword")
	Attempts = Session("PwdAttempts")
	Lockout = Session("LockOut")
	
	Attempts = Attempts + 1
	Session("PwdAttempts") = Attempts
	Public Sub Login
		
		Dim objConn
		Dim strConnection
		Set objConn = Server.CreateObject("ADODB.Connection")
		strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & _
		   Server.MapPath("data\accounts.mdb")
								
		objConn.Open (strConnection)
		
		Dim objRS
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "Accounts", objConn ,,, 2
		
		Do while not objRS.EOF
			Dim AUser
			AUser = objRS("UserName")
			
			if UCase(Username) <> UCase(AUser) then
				Verified = False
				ErrorMsg = "Invalid Login credentials. Number of attempts: " & Attempts
			else
				Dim Pass
				Pass = objRS("Password")
				
				if Password = Pass then
					Session("AdminLevel") = objRS("Admin")
					Session("Rank") = objRS("Rank")
					Verified = true
					exit do
				else
					Verified = false
					ErrorMsg = "Invalid Login credentials. Number of attempts: " & Attempts
					exit do
				end if
			end if
			objRS.MoveNext
		Loop
		
		objConn.Close
		set objConn = Nothing
	End Sub
	if Lockout = "No" then
		if Attempts <= 2 then
			Login
		else
			Login
			if Verified = false then
				ErrorMsg = "Too many invalid attempts. You have been locked out for 1 minute."
				Session("LockOut") = Timer + 60
			end if	
		end if
	else
		Dim CurTime
		CurTime = Timer
		if CurTime =< Lockout then
			Dim Remaining
			Remaining = Int(Lockout - CurTime)
			Dim SecondsLeft
			Dim MinutesLeft
			
			Do while Remaining >= 60
				MinutesLeft = MinutesLeft + 1
			loop
			
			SecondsLeft = Remaining
			if MinutesLeft > 0 then
				ErrorMsg = "You are still locked out for " & MinutesLeft & " minutes and " & SecondsLeft & " seconds due to too many invalid login attempts"
			else
				ErrorMsg = "You are still locked out for " & SecondsLeft & " seconds due to too many invalid login attempts"
			end if
			Verified = false
		else
			Attempts = 1
			Session("PwdAttempts") = Attempts
			Session("Lockout") = "No"
			Login
		end if
	end if
	
	if Verified then	
		Session("Verified") = True
		Session("Username") = Username
		Session("ErrorMsg") = ""
		Session("TestScore") = 0
		Session("PracticeVer") = -1
		Session("PwdAttempts") = 0
		Response.Redirect("http://cozmathematics.1apps.com/main.asp")
	else
		Session("Verified") = False
		Session("Username") = ""
		Session("ErrorMsg") = ErrorMsg
		Response.Redirect("http://cozmathematics.1apps.com/login.asp")
	end if
%>
