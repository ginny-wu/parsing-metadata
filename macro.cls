Public Sub findTags() ' Begin VBA procedure

' Declare variables
Dim http As Object
Dim html As New HTMLDocument
Dim startups As Object
Dim startup As HTMLHtmlElement
Dim tags As Object
Dim tag As HTMLHtmlElement
Dim i As Integer
Dim s As Integer
Dim t As Integer

' Retrieve data
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://www.startupstadium.tw/startups/", False
http.send
html.body.innerHTML = http.responseText ' Strips down unnecessary information

' Populate list of startups and tags
Set startups = html.getElementsByClassName("summary-thumbnail-container sqs-gallery-image-container")
Set tags = html.getElementsByClassName("summary-metadata-item summary-metadata-item--tags")

i = 2 ' Start at row below header row
s = 0 ' Start at first index
t = 0

Set startup = startups(s)
Set tag = tags(t)

For Unique = 0 To tags.Length - 1 Step 3

For Count = 0 To tag.getElementsByTagName("a").Length - 1 ' Populate row based on number of tag occurences (per startup). Number of tags is equal to number of hyperlink tags.     
    Set startup = startups(s)
    Cells(i + Count, 2).Value = tag.getElementsByTagName("a")(Count).innerText
    Cells(i + Count, 1).Value = startup.getAttribute("data-title")
        
Next   
    t = t + 3
    Set tag = tags(t)
    i = i + Count ' Move onto next blank row
    s = s + 1 ' Move to next index
    
Next

End Sub
