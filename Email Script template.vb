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
    
    Dim boot As String
    boot = "<link href=https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css rel=stylesheet integrity=sha256-7s5uDGW3AHqw6xtJmNNtr+OBRJUlgkNJEo78P4b0yRw= sha512-nNo+yCHEyn0smMxSswnf/OnX6/KwJuZTlNZBjauKhTK0c+zT+q5JOCx0UFhXQ6rJR9jg6Es8gPuD2uZcYDLqSw== crossorigin=anonymous>"
    
    Dim image As String
    image = "<div class=text-center ><img src=http://i.imgur.com/WdanDK8.jpg ></div>"
    
    
    'English sectoin
    Dim fileLink As String
    Text = ""
    Dim intro As String
    intro = ""
    Dim firsParagraph As String
    firsParagraph = "<h1 class = text-center>Hello worldlings, I bring you peace</h1>"
    Dim secondParagraph As String
    secondParagraph = "<p class=text-center>We, the kingdom of sheep of dimension C31, would like to offer you a deveoper position in our dimension.</p>"
    Dim thirdParagraph As String
    thirdParagraph = "<p class=text-center>Your skills were noticed and we would like you to join our great sheep developers.</p>"
    Dim fourthParagraph As String
    fourthParagraph = "<p class=text-center>We hope you accept our offer.</p>"
    Dim fifthParagraph As String
    fifthParagraph = "<p class=text-center>Contact us at sheepDevelopers@sheepy.C31</p>"
    
    
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
        'Uncomment if you want to extract a link
        'For Each HL In ActiveSheet.Hyperlinks
        '    HL.Range.Offset(0, 0).Value = HL.Address
        'Next
        
        
        'To send emails from specified number of cell in the For i=1 To (which ever number you would like to specify)
    For i = 1 To 1 'change the columns as needed
        Set olApp = New Outlook.Application
        Set olMail = olApp.CreateItem(olMailItem)
        On Error Resume Next
        
        With olMail
            'information present in the Excel sheet
            .To = Cells(i, 1).Value
            .Subject = "Message from the Sheep King"
            
             'Organizing the strings into HTML format
            .HTMLBody = "<head>" + boot + "</head>" + "<body style=background-color:#e0e0d1>" + "<br><br>" + firsParagraph + "<p class=text-center>" + "<br>" + "Dear" + " " + Cells(i, 3) + "," + "</p>" + secondParagraph + thirdParagraph + "<br>" + fourthParagraph + fifthParagraph + "<br>" + image + "</body>"
          
            'Displays message
            .Display
            '\/Wait function used to allow some time for the HTML elements to Render
             Application.Wait (Now + TimeValue("0:00:01"))
            '.Send
            
            
        End With
            
        Set olMail = Nothing
        Set olApp = Nothing
    
    Next

End Sub

