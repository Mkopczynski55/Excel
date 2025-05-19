Attribute VB_Name = "PQFormatter"
Option Explicit

Public Const API_URL As String = "https://m-formatter.azurewebsites.net/api/v2"
'JsonCoverter.bas needs to be imorted to the same project + Microsoft Scripting Runtime added as reference in the References

Sub SendFormatterRequest()

    Dim awb As Workbook
    Dim qryName As String
    Dim rawCode As String
    Dim esc As String
    Dim jsonBody As String
    Dim http As Object
    Dim rawResponse As String
    Dim parsed As Object
    Dim resultValue As String
    
    Set awb = ActiveWorkbook
    
    qryName = Application.InputBox("Provide query's name:", "Power Query Formatter", Type:=2)

    If qryName = "False" Then Exit Sub
    
    On Error GoTo QueryNotFound
    
    rawCode = awb.Queries(qryName).Formula
    
    On Error GoTo 0
    
    esc = rawCode
    esc = Replace(esc, vbCrLf, " ")
    esc = Replace(esc, vbCr, " ")
    esc = Replace(esc, vbLf, " ")
    esc = Replace(esc, vbTab, " ")
    esc = Replace(esc, "\", "\\")
    esc = Replace(esc, """", "\""")
    
    jsonBody = _
                    "{" & _
                                """code"":""" & esc & """," & _
                                """resultType"":""text""," & _
                                """lineWidth"":50," & _
                                """indentationLength"":2," & _
                                """includeComments"":true," & _
                                """surroundBracesWithWs"":false," & _
                                """indentSectionMembers"":true," & _
                                """alignLineCommentsToPosition"":40," & _
                                """alignPairedLetExpressionsByEqual"":""singleline""," & _
                                """alignPairedRecordExpressionsByEqual"":""singleline""," & _
                                """indentation"":""spaces""," & _
                                """ws"":"" ""," & _
                                """lineEnd"":""\n""" & _
                    "}"
                            
    'Debug.Print "OUTGOING JSON:" & vbCrLf & jsonBody

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    With http
    
        .Open "POST", API_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .send jsonBody
        
        rawResponse = .responseText
        
        'Debug.Print "RAW RESPONSE:" & vbCrLf & rawResponse
        
        Set parsed = JsonConverter.ParseJson(rawResponse)
        
        If Not parsed("success") Then
        
            MsgBox "PQ Formatter Error: " & parsed("errors")(1)("message"), vbCritical
            
            Exit Sub
            
        End If
        
        resultValue = parsed("result")
        awb.Queries(qryName).Formula = resultValue
        MsgBox "Query was successfully formatted.", vbOKOnly + vbInformation, "Formatting complete"
        
    End With
    
    Set http = Nothing
    Set parsed = Nothing
    
    Exit Sub
    
QueryNotFound:
    
    MsgBox "Query not found!", vbOKOnly + vbCritical, "Failed to find query"
    
End Sub


