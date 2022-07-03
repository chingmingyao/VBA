Sub 桿號查詢_DMQS_批次查詢()
Dim Mysheetname As String
Mysheetname = ActiveSheet.Name
Application.ScreenUpdating = 0
Range("A3:B65535") = ""
With Sheets(Mysheetname)
last_row = Cells(65535, 3).End(xlUp).Row

For i = 3 To last_row

 Range("A" & i) = 桿號查詢_DMQS_2(Range("C" & i))
 Range("B" & i) = Tailoc_to_loc(Range("A" & i))
    If Range("A" & i) <> "無此桿號" Then
         Call 座標加超連結("B", i)
    Else
        
         Range("B" & i) = ""
    End If

Next
End With
Application.ScreenUpdating = 1
End Sub

Sub 座標TO桿號_DMQS()
Application.ScreenUpdating = 0
Dim My_Coordinate, My_pole As String
With Sheets("桿號查詢")
If .Range("F2") = 0 Or .Range("F2") = "" Then
Exit Sub
End If

    Dim rawResponseText As String
    Dim oXML As Object
    Dim myploe As String
    Set oXML = CreateObject("MSXML2.XMLHTTP")
 
    My_Coordinate = .Range("F2").Value
    'myploe = "%E5%9F%A4%E9%A0%AD%E9%AB%98%E5%B9%B9#2-1"
    My_Coordinate = Trim(CStr(My_Coordinate))
    
    With oXML
        '.Open "GET", "http://10.210.35.218/DMQService/api/PoleInformation/GetPoleInformation?TPCLID=G4227GA57&P_NUMB=&COUNTY=&DISTRICT=&LI=&CODETXT=", 0
        .Open "GET", "http://10.210.35.218/DMQService/api/PoleInformation/GetPoleInformation?TPCLID=" & My_Coordinate & "&P_NUMB=&COUNTY=&DISTRICT=&LI=&CODETXT=", 0
        .setRequestHeader "Connection", "keep-alive"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Cookie", "ASP.NET_SessionId=0txljjjpxnuvxsauqnjsbejj"
        .send

        rawResponseText = convertraw(.responseBody, "UTF-8")
        'Debug.Print rawResponseText
      
    End With
    Set oXML = Nothing
    
    
    
    '檢查有無桿號
    If InStr(1, rawResponseText, "p_Num") <> 0 Then
        
    '爬桿號
     My_pole = Split(Split(rawResponseText, """,""p_Numb"":""")(1), """,""p_Len1")(0)
     Range("G2") = My_pole
    Else
    Range("G2") = "無此桿號"
    Application.ScreenUpdating = 1
    Exit Sub
    End If
    
End With
Application.ScreenUpdating = 1
End Sub
  
 Sub 座標加超連結(col As String, row_number) 'col請輸入開始之欄位之字串

'加入超連結
Application.ScreenUpdating = 0
Dim Mysheetname As String
Mysheetname = ActiveSheet.Name
    With Worksheets(Mysheetname)
            .Hyperlinks.Add Anchor:=.Range(col & row_number), _
            Address:="https://www.google.com.tw/maps/place/" & .Range(col & row_number), _
            ScreenTip:="Microsoft Web Site", _
            TextToDisplay:="" & .Range(col & row_number) & ""
    End With
'查詢桿號
   
End Sub 
 
