 To download and compare milestones of Agile Central Rally using Excel VBA, you can follow these steps:

1. Create a new Excel workbook and save it as a macro-enabled workbook (.xlsm).
2. Open the Visual Basic Editor (VBE) by pressing Alt + F11.
3. In the VBE, create a new module by clicking on Insert > Module.
4. In the module, create a new function called "DownloadMilestones" that takes in the URL of the Agile Central Rally project as a parameter.
5. In the function, use the "WinHttp.WinHttpRequest" object to send a GET request to the Agile Central Rally API to retrieve the milestones.
6. Parse the JSON response from the API and extract the milestone data.
7. Create a new Excel worksheet to store the milestone data.
8. Use the "Range" object to write the milestone data to the worksheet.
9. Create a new function called "CompareMilestones" that takes in the URL of the Agile Central Rally project as a parameter.
10. In the function, use the "WinHttp.WinHttpRequest" object to send a GET request to the Agile Central Rally API to retrieve the milestones.
11. Parse the JSON response from the API and extract the milestone data.
12. Compare the milestone data from the two API calls and highlight any differences.

Here is an example of the code for the "DownloadMilestones" function:
```
Function DownloadMilestones(url As String)
    
    ' Create a new WinHttp.WinHttpRequest object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Send a GET request to the Agile Central Rally API
    http.Open "GET", url, False
    http.Send
    
    ' Parse the JSON response from the API
    Dim json As Object
    Set json = JsonConverter.ParseJson(http.ResponseText)
    
    ' Extract the milestone data from the JSON response
    Dim milestones As Object
    Set milestones = json("milestones")
    
    ' Create a new Excel worksheet to store the milestone data
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    ' Write the milestone data to the worksheet
    Dim i As Long
    For i = 1 To milestones.Count
        ws.Cells(i, 1).Value = milestones(i)("name")
        ws.Cells(i, 2).Value = milestones(i)("description")
        ws.Cells(i, 3).Value = milestones(i)("startDate")
        ws.Cells(i, 4).Value = milestones(i)("endDate")
    Next i
    
    ' Clean up
    Set http = Nothing
    Set json = Nothing
    Set milestones = Nothing
    Set ws = Nothing
    
End Function
```
And here is an example of the code for the "CompareMilestones" function:
```
Function CompareMilestones(url1 As String, url2 As String)
    
    ' Create a new WinHttp.WinHttpRequest object
    Dim http1 As Object
    Set http1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Send a GET request to the Agile Central Rally API
    http1.Open "GET", url1, False
    http1.Send
    
    ' Parse the JSON response from the API
    Dim json1 As Object
    Set json1 = JsonConverter.ParseJson(http1.ResponseText)
    
    ' Extract the milestone data from the JSON response
    Dim milestones1 As Object
    Set milestones1 = json1("milestones")
    
    ' Create a new WinHttp.WinHttpRequest object
    Dim http2 As Object
    Set http2 = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Send a GET request to the Agile Central Rally API
    http2.Open "GET", url2, False
    http2.Send
    
    ' Parse the JSON response from the API
    Dim json2 As Object
    Set json2 = JsonConverter.ParseJson(http2.ResponseText)
    
    ' Extract the milestone data from the JSON response
    Dim milestones2 As Object
    Set milestones2 = json2("milestones")
    
    ' Compare the milestone data from the two API calls
    Dim i As Long
    For i = 1 To milestones1.Count
        If milestones1(i)("name") <> milestones2(i)("name") Then
            ws.Cells(i, 1).Value = "Milestone name mismatch"
        End If
        If milestones1(i)("description") <> milestones2(i)("description") Then
            ws.Cells(i, 2).Value = "Milestone description mismatch"
        End If
        If milestones1(i)("startDate") <> milestones2(i)("startDate") Then
            ws.Cells(i, 3).Value = "Milestone start date mismatch"
        End If
        If milestones1(i)("endDate") <> milestones2(i)("endDate") Then
            ws.Cells(i, 4).Value = "Milestone end date mismatch"
        End If
    Next i
    
    ' Clean up
    Set http1 = Nothing
    Set json1 = Nothing
    Set milestones1 = Nothing
    Set http2 = Nothing
    Set json2 = Nothing
    Set milestones2 = Nothing
    Set ws = Nothing
    
End Function
```
Note that this code assumes that the Agile Central Rally API is available and that the JSON response from the API is in the correct format. You may need to modify the code to handle any errors or exceptions that may occur.
