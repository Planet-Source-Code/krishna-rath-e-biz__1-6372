Attribute VB_Name = "Module1"
'***********************************************************
'               E-BIZ. V 1.0
'           KRISHNA RATH. email: krath@engineer.com
'               March 2000
'***********************************************************
'
'E-biz is an interactive "bussiness WinCGI software". It uses the standard POST
'and GET Methods to get the user information and handles it accordingly.
'It has also a user database. Registered users can shop online!
'Unregistered users can create a new account.
'If the user name has been registered then you will be asked to enter again!
' And in case you forget your password, you can retrive it by giving the email you entered while signing up!
'
'This Program uses CGI32.BAS module to handle all CGI calls. I like this module very much!
'
'
'Please read the "install.html" which comes with the zip file to understand
' the installation process
'
'This program shows:
'1. How to register a new user
'2. How to accept an already registered user
'3. Retriving a lost password of a registered user
'4. Online shopping...shop till you run out of money!
'
'You can add many web pages, add a VISA card membership have the address of
' the users . Corresponding changes have to be made in this program.
'
'Oh yes! E-com programs are not made using CGI calls...cause they can be easily trapped!
'
'Sorry for any typing or spelling mistakes...I made this program in a big hurry! I had
'to design web pages along with the corresponding changes in this program etc and
'then get back to my college work!
'
'Since this is Version 1.0 There are not many error handling codes.
'Please report Bugs and modifications at krath@engineer.com
'If you make any decent change in this program please mail me the source code
' so that I can too work on it.
'

Public db As Database
Public rs As Recordset
Public qd  As QueryDef
Public LoginSuccess, UserRegistered As Boolean  'flags that help while login
Public UserName, UserPassword, UserEmail As String  'the username, password and email
Public PurchasesMade As String

Type MONEY                  '2 currencies are used: US$ and Indian Ruppes
    USdollar As Long
    IndianRu As Long
End Type

Global B_Amount(8) As MONEY      'Book_amount for each book.
Global CD_Amount(4) As MONEY     'the number represents the
Global MB_Amount(4) As MONEY     'cost of each item covered in a check box

Public USamount, IndAmount As Long  'Books_countryAmount


Sub Inter_Main()
'Interactive start
Dim x As String
x = x & "E-biz is an interactive 'bussiness WinCGI software'. It uses the standard POST "
x = x & "and GET Methods to get the user information and handles it accordingly."
x = x & " It has also a user database. Registered users can shop online!"
x = x & " Unregistered users can create a new account."
x = x & " And in case you forget your password, you can retrive it by giving the email you entered while signing up!" & vbNewLine & vbNewLine
x = x & " This Program uses CGI32.BAS module to handle all CGI calls. I like this module very much!"
x = x & " Please read the 'install.html' which comes with the zip file to understand"
x = x & " the installation process" & vbNewLine & vbNewLine
x = x & " This program shows:" & vbNewLine
x = x & " 1. How to register a new user" & vbNewLine
x = x & " 2. How to accept an already registered user" & vbNewLine
x = x & " 3. Retriving a lost password of a registered user" & vbNewLine
x = x & " 4. Online shopping...shop till you run out of money!" & vbNewLine & vbNewLine
x = x & " E-Biz By Krishna Rath, krath@engineer.com, http://rath.8k.com"
MsgBox x, vbOKOnly, "About E-Biz"


End Sub

Sub CGI_Main()
On Error GoTo errhan

If CGI_RequestMethod = "POST" Then 'If POST then
    'Check for the value of "formname" of the web page
    'formname is a hidden field in the web page to identify
    'which part is been called
    
    Select Case GetSmallField("formname")
    
        'Check if it is the user name and Password
        Case "login"
            'get the user name  and password
            UserName = GetSmallField("T1")
            UserPassword = GetSmallField("T2")
            'Check if the user exits. If not display the register form
            CheckUserName UserName, UserPassword
            'after checking the login succes
                If LoginSuccess = True Then ShowShop  'show the shopping centre
        
        Case "showpass"    'if the user has forgotten his password and asks for it
            'get the username and email address
            UserName = GetSmallField("T1")
            UserEmail = GetSmallField("T2")
            ShowPassword UserName, UserEmail
            
        Case "register"   'If the user wants to register for the first time
            UserName = GetSmallField("T1")
            UserPassword = GetSmallField("T2")
            UserEmail = GetSmallField("T3")
            RegisterUser UserName, UserPassword, UserEmail
    
    
    'Case for books
    'the books and CDs section work on the same principle..so I have not decorated
    'the CD section Web-page the way I did for the books.
        Case "book"
            'No need for username here as it has already been entered and loginsucces is true
            'We have to get the CheckBoxes that were ticked. I have given 2 prices
            'US$ and Indian Rs. This is because most e-shoping sites have options
            'of paying in diffrent currcency. I am in India, and an Indian
            'will give the money in Indian Rs, while for someone else in another country  will have
            ' to pay in US$.
            '
            'We put the book section in a separate sub as it would mess up this loop!
            BookSection 'Goto the book section
                    
    'case for the CDs
        Case "cd"
            CDsection
        
    'case for the Movie Theatres
        Case "mb"
        'Book the tickets
        BookTicket
        
        
        'if there was an error in the formname or it is not available
        'display an error
        Case Else
            MsgBox "There was an error in the server"
                
    End Select 'end the getsmallfield("formname")
    
End If  'Ends the POST session

errhan:
'simply exit the sub if any error occurs.. Mostly it occurs if a POST is done
'without giving a key called formname.

End Sub
Public Sub CheckUserName(ByVal uname As String, ByVal UPassword As String)
'checks whether the username and password exists or not

Dim dbPath, SQL_str As String 'the database path
dbPath = App.Path & "\ebiz.mdb"     'set the database path
'Opening the database and checking for username and password
Set db = OpenDatabase(dbPath, dbOpenDynaset)
SQL_str = "SELECT user FROM UID WHERE user='" & UserName & "' AND pass='" & UserPassword & "';"
Set rs = db.OpenRecordset(SQL_str)

If rs.RecordCount = 0 Then
    ErrorLogin    'there was an error while login in.
Else
    LoginSuccess = True
End If

Set db = Nothing
Set rs = Nothing

End Sub

Public Sub ErrorLogin()
' The error while logining  can be
'1. The password was incorrect
'2. The user name does not exits
'In this program we ask the email of the user. Show the user the 2 options he
' wants. Ie. Login in again with correct password or as a new user
'Showing the Options HTML file...

' the HTML page is a bit complex due to the tables!

LoginSuccess = False 'Login Succes was a failure

Send "<html> <head> <title>ERROR IN LOGIN</title> </head>"
Send "<body> <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
Send "<tr><td width=""100%"" bgcolor=""#800000"" valign=""top""><font color=""#80FFFF"">Rath-India website."
Send "http://rath.8k.com</font></td></tr><tr>"
Send "<td width=""100%"" bgcolor=""#008000"" valign=""top"" align=""center""><h1><font color=""#FFFF00"">E-Biz"
Send "</font></h1> </td></tr>  <tr>"
Send "<td width=""100%"" bgcolor=""#800000"" valign=""top""><font color=""#FFFF00""><p align=""right"">The"
Send "Ultimate shoping centre in the world</font></td></tr><tr>"
Send "<td width=""100%"" valign=""top""><h3 align=""left""><font color=""#FF0000"">Sorry! There was an"
Send "Error while login</font></h3><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
Send "<tr><td width=""100%"">Please try your password again by<a href=http://" & CGI_RemoteAddr & "/ebiz.htm" & "> login in again</a></td>"
Send "</tr><tr><td width=""100%""><br>If you have forgotten you Password type in your username and your email address<table"
Send "border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr><td width=""25%""><font color=""#FF0000""><strong>User Name</strong></font></td>"
Send "<td width=""75%""><font color=""#FF0000""><strong>Email Address</strong></font></td>"
Send "</tr></table><form method=""POST"" action=""/cgi-win/cgiebiz/ebiz.exe""><p><input type=""hidden"" name=""formname"" value=""showpass""> <input type=""text"" name=""T1"" size=""20""><input type=""text"" name=""T2"" size=""20""><input"
Send "type=""submit"" value=""Show Me My Password"" name=""B1""></p></form></td></tr><tr><td width=""100%""><strong><font color=""#004080""></font></strong></td>"
Send "</tr><tr><td width=""100%""><strong><font color=""#004080""><a href=http://" & CGI_RemoteAddr & "/ebiz/register.htm" & ">Unregistered users can sign up here</a></font></strong></td></tr></table>"
Send "<p><font color=""#400080""><em>NOTE: The whole excercise is just a 'DEMO&quot; program showing WinCGI using Visual Basic. All the items shown are fictious and any resemblance to anyone living is just a 'coincidence'. </em></font></td>"
Send "</tr></table></body></html>"


End Sub
Public Sub ShowPassword(uname, uemail)

Dim dbPath, SQL_str As String 'the database path
dbPath = App.Path & "\ebiz.mdb"     'set the database path
'Opening the database and checking for username and email
Set db = OpenDatabase(dbPath, dbOpenDynaset)
SQL_str = "SELECT pass FROM UID WHERE user='" & UserName & "' AND email='" & UserEmail & "';"
Set rs = db.OpenRecordset(SQL_str)

If rs.RecordCount = 0 Then
ErrorLogin    'there was an error while login in.
Else
'display the password
Send ("<html>")

Send ("<head>")
Send ("<meta http-equiv=""Content-Type""")
Send ("content=""text/html"">")
Send ("<title>Your Password</title>")
Send ("</head>")
Send ("<body bgcolor=""#FFFFF0"">")
Send ("<h1 align=""center"">Your Password</h1>")
Send ("<p>The password is : <strong>" & rs.Fields(0) & "</strong></p>")
Send ("<p>Login <a href=http://" & CGI_RemoteAddr & "/ebiz.htm" & ">here</a></p>")
Send ("</body></html>")
End If

Set db = Nothing
Set rs = Nothing

End Sub

Public Sub RegisterUser(x, y, z)
' Register the new user
' we have to also see that a username is already registered or not.
'if the name is registered prompt the user to put another name.

AlreadyUser UserName    'Check if the user name is registered or not

Dim dbPath, SQL_str As String   'define the filename etc
dbPath = App.Path & "\ebiz.mdb"


    If UserRegistered = True Then
        Set db = Nothing
        Exit Sub
    Else    'else if new User
    'insert the information
        Set db = OpenDatabase(dbPath, dbOpenDynaset) 'open the database
        SQL_str = "INSERT INTO UID VALUES('" & UserName & "','" & UserEmail & "','" & UserPassword & "');"
        db.QueryDefs.Delete ("add_user")
        Set qd = db.CreateQueryDef("add_user")
        qd.SQL = SQL_str
        db.Execute ("add_user") 'execute the code
        
    End If

Set db = Nothing
Set qd = Nothing

LoginSuccess = True ' new user logs in and ...
ShowShop    'Since the user is registered he can directly do to shop

End Sub

Public Sub AlreadyUser(x)
Dim dbPath, SQL_str As String
dbPath = App.Path & "\ebiz.mdb"     'set the database path
'Opening the database and checking for username and email
Set db = OpenDatabase(dbPath, dbOpenDynaset)
SQL_str = "SELECT user FROM UID WHERE user='" & UserName & "';"
Set rs = db.OpenRecordset(SQL_str)

If rs.RecordCount = 0 Then
UserRegistered = False   ' the username does not exits
Else
'the user name exits in the database..so promt the user to register again
UserRegistered = True
Send ("<html>")

Send ("<head>")
Send ("<meta http-equiv=""Content-Type""")
Send ("content=""text/html"">")
Send ("<title>Register</title>")
Send ("</head>")
Send ("<body bgcolor=""#FFFFF0"">")
Send ("<h1 align=""center"">Register again</h1>")
Send ("<p>The user name that you slected has already been taken Please click <a href=http://" & CGI_RemoteAddr & "/ebiz/register.htm" & " >here</a> to register again </p>")
Send ("</body></html>")
End If

Set db = Nothing
Set rs = Nothing

End Sub

Public Sub ShowShop()
'Main entrance for showing the shopping world after registering, logining etc
'Display the index for shopping
'This page was created using Frontpage and then pasted here...so you might
'not be able to make much sense out of!
'You can create a link to a web page on the server itself unlike
'creating a "dynamic" one like here
'Why I did this was to include the user name in the web page.
'I have included 2  advertisments just for fun!


Send "<html><head><meta name=""GENERATOR"" content=""Microsoft FrontPage 3.0""><title>E-Biz</title></head>"
Send "<body><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr><td width=""100%"" bgcolor=""#800000"" valign=""top"" colspan=""2""><font color=""#80FFFF"">Rath-India"
Send "website. <a href=""http://rath.8k.com"">http://rath.8k.com</a></font></td></tr><tr><td width=""100%"" bgcolor=""#008000"" valign=""top"" colspan=""2"" align=""center""><font color=""#FFFF00""><h1>E-Biz </h1></font></td></tr><tr>"
Send "<td width=""100%"" bgcolor=""#800000"" valign=""top"" colspan=""2""><font color=""#FFFF00""><palign=""right""></font><font color=""#80FFFF"">The Ultimate shoping centre in the world</font></td></tr><tr><td width=""24%"" valign=""top""><div align=""left""><table border=""0"" cellpadding=""0"""
Send "cellspacing=""0"" width=""100%"" align=""left""><tr><td width=""89%"" bgcolor=""#000080""><strong><font color=""#80FF80""><small>Choose the Sections to Browse through the virtual Shopping centre</small><br></font><font color=""#400080""><small>temmp</small></font></strong></td><td width=""11%"" rowspan=""7""></td>"
Send "</tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/books.htm" & "><strong><font color=""#FFFFFF"">Books</font></strong></a></td></tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/cd.htm" & "><strong><font"
Send "color=""#FFFFFF"">CDs</font></strong></a></td></tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/mb.htm" & "><strong><font"
Send "color=""#FFFFFF"">Movie Bookings</font></strong></a></td></tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><font color=""#400080"">Temp</font></td></tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><strong><font color=""#80FF00"">Sign Out</font></strong></td>"
Send "</tr><tr><td width=""89%"" bgcolor=""#000080"" valign=""top""><strong><font color=""#FFFF00"">About E-Biz<br>Author:<br>Krishna Rath<br>krath@engineer.com</font></strong></td>"
Send "</tr></table></div></td><td width=""76%"" valign=""top""><br>Welcome to E-Biz, <strong>" & UserName & "</strong>  Here you can shop for books, CDs and book tickets for movies. All here for your service!<br>Choose the section and feel free to shop anything here.<br> <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""105%""><tr><td width=""82%"" valign=""top"" align=""left"" bgcolor=""#808000""><p align=""center""><a"
Send "href=http://" & CGI_RemoteAddr & "/ebiz/books.htm" & "><font color=""#FFFF00""><strong>Books</strong></font></a></td><td width=""23%"" valign=""top"" align=""left"" bgcolor=""#008000""><font color=""#FFFF00"">Visit<a href=""http://rath.8k.com""> rath.8k.com </a></font></td></tr><tr><td width=""82%"" valign=""top"" align=""left"">Buy the latest books online. Select the newest books from all over the world. If you buy my book, you will get a special discount"
Send "of 35%. <br></td><td width=""23%"" rowspan=""5"" valign=""top"" align=""left""><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" height=""176""><tr><td width=""100%"" height=""72"" valign=""top""< a href=""http://rath.8k.com""><img src=http://" & CGI_RemoteAddr & "/ebiz/ad1.jpg" & " width=""70"" height=""70""alt=""ad1.jpg (5818 bytes)""></a></td></tr>"
Send "<tr><td width=""100%"" height=""104"" valign=""top""><a href=""http://planet-source-code.com""><img src=http://" & CGI_RemoteAddr & "/ebiz/psc.gif" & " width=""175"" height=""101"" alt=""psc.gif""></a></td></tr></table></td></tr><tr><td width=""82%"" valign=""top"" align=""left"" bgcolor=""#808000""><p align=""center""><a href=http://" & CGI_RemoteAddr & "/ebiz/cd.htm" & "><font color=""#FFFF00""><strong>CDs</strong></font></a></td>"
Send "</tr><tr><td width=""82%"" valign=""top"" align=""left"">Crazy for music? Then shop the CD section. Collection ranges from the oldies to the latest.<br></td></tr><tr><td width=""82%"" valign=""top"" align=""left"" bgcolor=""#808000""><p align=""center""><a href=http://" & CGI_RemoteAddr & "/ebiz/mb.htm" & "><font color=""#FFFF00""><strong>Movie Booking</strong></font></a></td></tr><tr><td width=""82%"" valign=""top"" align=""left"">Book a ticket in a movie theatre in you town. This offer is valid only in India.<br></td></tr></table>"
Send "<p>&nbsp;</p><p><font color=""#400080""><em>NOTE: The whole excercise is just a 'DEMO&quot; program showing WinCGI using Visual Basic. All the items shown are fictious and any resemblance to anyone living is just a 'coincidence'. Nothing actually works out, expect the calls you make to the server. You are not going to spend money here!</em></font></td>"
Send "</tr></table></body></html>"


End Sub

Public Sub BookSection()
'We know that we have 8 choices to make between the various books
'Also the ammount is in US$ and Rs...And note that the currencies have no exchange rate relation

'To find out whether a CheckBox is checked or not we use the Function FieldPresent.
'If a Checkbox is checked then the function returns true
'
'Find if the field is present or not. If present then add the total money

Dim ItemName(8) As String   'Name of each item

'first define the cost and name of each item
'we can put this in a database.
' I feel that will be better for larger data. But since here there are just 8
'entries it will be better to put in the information here itself

B_Amount(1).IndianRu = 450: B_Amount(1).USdollar = 25: ItemName(1) = "Chinese Intrigues by Richie Bernard"
B_Amount(2).IndianRu = 550: B_Amount(2).USdollar = 30: ItemName(2) = "American Conspiracy by Xin Chan"
B_Amount(3).IndianRu = 1050: B_Amount(3).USdollar = 75: ItemName(3) = "Soccer my life by Romaldo "
B_Amount(4).IndianRu = 100: B_Amount(4).USdollar = 5: ItemName(4) = "The Cricket Scandal by M Azziz."
B_Amount(5).IndianRu = 1300: B_Amount(5).USdollar = 50: ItemName(5) = "The Day by William Shookspeare"
B_Amount(6).IndianRu = 750: B_Amount(6).USdollar = 28: ItemName(6) = "The Night by David Boon"
B_Amount(7).IndianRu = 75: B_Amount(7).USdollar = 5: ItemName(7) = "Hacking made easy by Krishna Rath"
B_Amount(8).IndianRu = 75: B_Amount(8).USdollar = 5: ItemName(8) = "E-biz: The future by Krishna Rath"


Dim n As String 'just a dummy to indicate the CheckBox C1,C2 etc
Dim i As Integer

PurchasesMade = "The Books bought are :<br>"
For i = 1 To 8      'for each Checkbox in Book section
    n = "C" & i     'increase the CheckBox
    If FieldPresent(n) Then 'if true then
        USamount = USamount + B_Amount(i).USdollar
        IndAmount = IndAmount + B_Amount(i).IndianRu
        PurchasesMade = PurchasesMade & "<br>" & ItemName(i)
    End If
Next i

DisplayResult       'Display the result of the purchase

End Sub

Public Sub CDsection()
'Define the Cd section terms
Dim ItemName(4) As String   'Name of each item

'first define the cost and name of each item
'we can put this in a database.
' I feel that will be better for larger data. But since here there are just 4
'entries it will be better to put in the information here itself

CD_Amount(1).IndianRu = 425: CD_Amount(1).USdollar = 5: ItemName(1) = "Jennifer Lopez:On the 6"
CD_Amount(2).IndianRu = 400: CD_Amount(2).USdollar = 5: ItemName(2) = "Ricky Martin"
CD_Amount(3).IndianRu = 300: CD_Amount(3).USdollar = 10: ItemName(3) = "Kaho Na Pyaar Hai"
CD_Amount(4).IndianRu = 330: CD_Amount(4).USdollar = 12.25: ItemName(4) = "Pukar"

Dim n As String 'just a dummy to indicate the CheckBox C1,C2 etc
Dim i As Integer
PurchasesMade = "The CDs Perchased are: <br>"
For i = 1 To 4      'for each Checkbox in Book section
    n = "C" & i     'increase the CheckBox
    If FieldPresent(n) Then 'if true then
        USamount = USamount + CD_Amount(i).USdollar
        IndAmount = IndAmount + CD_Amount(i).IndianRu
        PurchasesMade = PurchasesMade & "<br>" & ItemName(i)
    End If
Next i

DisplayResult       'Display the result of the purchase


End Sub

Public Sub BookTicket()
'Since the ticket cost is the same ie Rs 75...we can give a for loop
'By the way since the Service was not allowed but India, US$ has been conviently ommited
'This section can be improved...I have not added a day function.ie which date will the user go to the theatre

Dim ItemName(4) As String
ItemName(1) = "Some Theatre: The World is more than enough"
ItemName(2) = "Another Theatre: Shout: Part 3, The last one"
ItemName(3) = "Hind Theatres: Khoon Ki Aag"
ItemName(4) = "A Theatre: Tum Kab Shadi Karogay"

Dim i As Integer
For i = 1 To 4
    MB_Amount(i).IndianRu = 75
Next i

Dim n As String 'just a dummy to indicate the CheckBox C1,C2 etc
Dim m As String     'indiacte the show time...combobox
PurchasesMade = "The Movies Booked are along with the show time: <br>"
For i = 1 To 4      'for each Checkbox in Book section
    n = "C" & i     'increase the CheckBox
    m = "D" & i     'increase the Drop Down Box
    If FieldPresent(n) Then 'if true then
        IndAmount = IndAmount + MB_Amount(i).IndianRu
        PurchasesMade = PurchasesMade & "<br>" & ItemName(i) & "; " & GetSmallField(m)
    End If
Next i

DisplayResult       'Display the result of the purchase

End Sub
Public Sub DisplayResult()
'Display the result of the purchase
    'Again the page was created using FrontPage and pasted here. The place where The username and purchases are
    'to be inserted is shown clearly among the hapazard code!


Send "<html><head><meta name=""GENERATOR"" content=""Microsoft FrontPage 3.0""><title>Result of Purchase.E-Biz</title></head>"
Send "<body><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">  <tr>    <td width=""100%"" bgcolor=""#800000"" valign=""top"" colspan=""2""><font color=""#80FFFF"">Rath-India  website. <a href=""http://rath.8k.com"">http://rath.8k.com</a></font></td>"
Send "</tr><tr><td width=""100%"" bgcolor=""#008000"" valign=""top"" colspan=""2"" align=""center""><font color=""#FFFF00""><h1>E-Biz </h1></font></td></tr><tr>"
Send "<td width=""100%"" bgcolor=""#800000"" valign=""top"" colspan=""2""><font color=""#FFFF00""><p align=""right""></font><font color=""#80FFFF"">The Ultimate shoping centre in the world</font></td></tr><tr><td width=""24%"" valign=""top""><div align=""left""><table border=""0"" cellpadding=""0"""
Send "cellspacing=""0"" width=""100%"" align=""left""><tr><td width=""89%"" bgcolor=""#000080""><strong><font color=""#80FF80""><small>Choose the Sections to Browse through the virtual Shopping centre</small><br></font><font color=""#400080""><small>temmp</small></font></strong></td><td width=""11%"" rowspan=""7""></td>"
Send "</tr><tr> <td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/books.htm" & "><strong><font color=""#FFFFFF"">Books</font></strong></a></td> </tr>"
Send "<tr> <td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/cd.htm" & "><strong><font color=""#FFFFFF"">CDs</font></strong></a></td></tr> <tr>"
Send "<td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz/mb.htm" & "><strong><font color=""#FFFFFF"">Movie Bookings</font></strong></a></td></tr> <tr><td width=""89%"" bgcolor=""#000080"" align=""left""><font color=""#400080"">Temp</font></td>"
Send "</tr><tr><td width=""89%"" bgcolor=""#000080"" align=""left""><a href=http://" & CGI_RemoteAddr & "/ebiz.htm" & "><strong><font color=""#80FF00"">Sign Out</font></strong></a></td></tr> <tr>"
Send "<td width=""89%"" bgcolor=""#000080"" valign=""top""><strong><font color=""#FFFF00"">About E-Biz<br>Author:<br> Krishna Rath<br>krath@engineer.com</font></strong></td></tr>"
Send "</table> </div></td><td width=""76%"" valign=""top"">&nbsp;<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">     <tr>"
Send "<td width=""20%"" align=""center"" bgcolor=""#808000""><font color=""#FFFF00""><strong>User Info</strong></font></td> <td width=""80%"" align=""center"" bgcolor=""#800080""><font color=""#FFFF00""><strong>Purchases Made</strong></font></td>"
Send " </tr> <tr> <td width=""20%"" valign=""top"" rowspan=""3""><strong>"
'Print the username and date
Send UserName & "</strong><br><strong>" & Date & "</strong></td>"
Send "<td width=""80%"" valign=""top""><blockquote><p><em>"
'Give the infromation about purchases
Send PurchasesMade & "<br></em></p><p>&nbsp;</p></blockquote></td>"
Send "</tr> <tr><td width=""80%"" valign=""top"" bgcolor=""#808000""><font color=""#FFFF00""><strong>Total cost :</strong></font></td></tr>"
Send "<tr><td width=""80%"" valign=""top""><blockquote><blockquote>"
'print US amount
Send "<p align=""left"">In US$: " & USamount & "<br>"
'print Indian amount
Send "In Indian Ruppees: " & IndAmount & "</p>"
Send "</blockquote> </blockquote>  </td> </tr> </table> <p><font color=""#400080""><em><br>"
Send "NOTE: The whole excercise is just a 'DEMO&quot; program showing WinCGI using Visual Basic. All the items shown are dummy data.. Nothing actually works out, expect the calls you make to the server. You are not going to spend money here!</em></font></td>"
Send "</tr></table><p align=""right""><small><font color=""#0000A0""><em>©Krishna Rath 2000</em></font></small></p></body></html>"

End Sub
