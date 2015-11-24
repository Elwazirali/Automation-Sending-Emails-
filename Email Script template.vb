
'This script is to pull information from an Excel sheet and email a bunch of people specific information
'The variable fields in this script are the person's email and the specific link needed to be accessed by that person
'There are five paragraphs initialized as strings. You can add more paragraph by using this code: Dim newParagraph As String new paragraph =""
'and add whatever is needed in between the double quotation like so: "This is the new paragraph"
'To access this code, you can simply open an excel sheet and enable the developer option. After that, just open up VBA and copy this code
'This code was developed by Ali Elwazir. For any questions please contact me at elwazirali@gmail.com
'Note that I used HTML syntax to allow for importing pictures and to display the text in acceptable format


Sub SendMain()

	'Initialize objects of Outlook Application
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
	'Hyper creating hyper-link object
    Dim HL As Hyperlink
    
    
    'English sectoin
    Dim fileLink As String
    Text = ""
    Dim intro As String
    intro = ""
    Dim firsParagraph As String
    firsParagraph = ""
    Dim secondParagraph As String
    secondParagraph = ""
    Dim thirdParagraph As String
    thirdParagraph = ""
    Dim fourthParagraph As String
    fourthParagraph = ""
    Dim fifthParagraph As String
    fifthParagraph = ""
    
    
    'French Section
    Dim Text2 As String
    Text2 = ""
    Dim firstFrench As String
    firstFrench = ""
    Dim secondFrench As String
    secondFrench = ""
    Dim thirdFrench As String
    thirdFrench = ""
    Dim fourthFrench As String
    fourthFrench = ""
    Dim fifthFrench As String
    fifthFrench = ""
    
		'Extracts the actual HTTP links from the hyper links in an excel file
        For Each HL In ActiveSheet.Hyperlinks
            HL.Range.Offset(0, 0).Value = HL.Address
        Next
		'To send emails from specified number of cell in the For i=1 To (which ever number you would like to specify)
    For i = 1 To 1 'change the columns as needed
        Set olApp = New Outlook.Application
        Set olMail = olApp.CreateItem(olMailItem)
        On Error Resume Next
        
        With olMail
			'information present in the Excel sheet
            .To = Cells(i, 1).Value
            .Subject = Cells(i, 3).Value
             HL.Address = Cells(i, 2).Value
             Cells(i, 2) = Replace((Cells(i, 2).Value), " ", "%20")
			 'Organizing the strings into HTML format 
            .HTMLBody = intro + "<br><br>" + firsParagraph + 
			"<br><br>" + 
			secondParagraph 
			+ "<br><br>" + thirdParagraph + 
			"<br><br>" + fourthParagraph + "<br><br>" 
			+ Cells(i, 2).Value 
			+ "<br><br>" + fifthParagraph + 
			"<br><br><img src = ><br><br>" + Text + "<br><br>" + 
			"<hr style=height:.2em color=black>" + "<br><br>" + 
			firstFrench + "<br><br>" + secondFrench + 
			"<br><br>" + thirdFrench + "<br><br>" + 
			fourthFrench + "<br><br>" & vbCrLf & Cells(i, 2).Value + 
			"<br><br>" + fifthFrench + "<br><br><img src = ><br><br>" + Text2
            'Displays message
			.Display
            '\/Wait function used to allow some time for the HTML elements to Render
             Application.Wait (Now + TimeValue("0:00:01"))
            .Send
            
            
        End With
            
        Set olMail = Nothing
        Set olApp = Nothing
    
    Next

End Sub




