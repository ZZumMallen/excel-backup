' I think this might actually be done

Option Explicit
Sub Spray_GigaLoop()

    Dim ws As Worksheet
    Set ws = ActiveSheet
  
    If ActiveSheet.Name Like "*-Spray-*" Then
        Dim tbl As ListObject
        Dim cell As Range
        Dim pc As String
        Dim dbh As String
        Dim counter As Integer
        Dim finalprice As Integer
    
        For Each tbl In ws.ListObjects
            counter = 0
            For Each cell In tbl.ListColumns("Rec").DataBodyRange.Rows
                If cell.EntireRow.Hidden = False Then
                    If counter = 0 Then
                        pc = get_price_code(cell.Value)
                        dbh = get_size_code(cell.Offset(, -4).Value)
                        finalprice = firstPrice(pc, dbh)
                    Else
                        pc = get_price_code(cell.Value)
                        dbh = get_size_code(cell.Offset(, -4).Value)
                        finalprice = additionalPrice(pc, dbh)
                    End If
                    cell.Offset(, 3).Value = finalprice
                    With cell.Offset(, -1)
                        .Value = cell.Value
                        .Interior.ColorIndex = 36
                    End With
                    Debug.Print finalprice
                    counter = counter + 1
                End If
            Next cell

        Next tbl

    End If

End Sub
Function get_price_code(sprayCommon As String) As Variant

    Dim dict As New Dictionary
'    Set dict = CreateObject("Scripting.dictionary")
    
    'Adding Sprays to the dictionary
    dict.Add "AnthOak1", "PC1"
    dict.Add "AnthOak2", "PC1"
    dict.Add "AnthOak3", "PC1"
    dict.Add "AppleScab1", "PC1"
    dict.Add "AppleScab2", "PC1"
    dict.Add "AppleScab3", "PC1"
    dict.Add "AnthAsh1", "PC1"
    dict.Add "AnthAsh2", "PC1"
    dict.Add "AnthAsh3", "PC1"
    dict.Add "AnthMapl2", "PC1"
    dict.Add "AnthMapl3", "PC1"
    dict.Add "NBDoth1", "PC2"
    dict.Add "NBDoth2", "PC2"
    dict.Add "NCRhizo1", "PC2"
    dict.Add "NCRhizo2", "PC2"
    dict.Add "NCRhizo3", "PC2"
    dict.Add "ShBSp1", "PC2"
    dict.Add "ShBSp2", "PC2"
    dict.Add "ShBSp3", "PC2"
    dict.Add "ShBSpDo1", "PC2"
    dict.Add "ShBSpDo2", "PC2"
    dict.Add "ShBSpDo3", "PC2"
    dict.Add "BePB1", "PC4"
    dict.Add "BePB2", "PC4"
    dict.Add "BePB3", "PC4"
    dict.Add "BeTwgGd1", "PC4"
    dict.Add "BeTwgPn1", "PC4"
    dict.Add "BoCW1", "PC4"
    dict.Add "BoCW2", "PC4"
    dict.Add "BoPSI", "PC4"
    dict.Add "BoPSLep", "PC4"
    dict.Add "BoPSX", "PC4"
    dict.Add "FruitRed1", "PC4"
    dict.Add "MiSpider1", "PC4"
    dict.Add "MiSpider2", "PC4"
    
    If dict.Exists(sprayCommon) Then
        get_price_code = dict(sprayCommon)
    Else
        get_price_code = "PC3"
    End If
    
End Function


Function get_size_code(dbh As Double) As Variant

    Select Case dbh
        Case Is < 3: get_size_code = "S1" 'Petite
        Case 3 To 6.99: get_size_code = "S2" 'X-small
        Case 7 To 10.99: get_size_code = "S3" 'Small
        Case 11 To 14.99: get_size_code = "S4" 'Medium
        Case 15 To 18.99: get_size_code = "S5" 'Large
        Case 19 To 23.99: get_size_code = "S6" 'X-Large
        Case Is > 24: get_size_code = "S7" 'Jumbo
    End Select

End Function

Function additionalPrice(pc As String, dbh As String) As Variant
    Dim codeDict As New Dictionary
    Dim priceDict As Dictionary
     
    Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 23
        .Add "S2", 35
        .Add "S3", 41
        .Add "S4", 46
        .Add "S5", 69
        .Add "S6", 114
        .Add "S7", 173
    End With
    codeDict.Add "PC1", priceDict
    
        Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 23
        .Add "S2", 35
        .Add "S3", 41
        .Add "S4", 46
        .Add "S5", 69
        .Add "S6", 114
        .Add "S7", 173
    End With
    codeDict.Add "PC2", priceDict
    
        Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 35
        .Add "S2", 46
        .Add "S3", 51
        .Add "S4", 58
        .Add "S5", 86
        .Add "S6", 127
        .Add "S7", 196
    End With
    codeDict.Add "PC3", priceDict
    
        Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 23
        .Add "S2", 35
        .Add "S3", 41
        .Add "S4", 46
        .Add "S5", 69
        .Add "S6", 114
        .Add "S7", 173
    End With
    codeDict.Add "PC4", priceDict
    
    If codeDict.Exists(pc) Then
        additionalPrice = codeDict(pc)(dbh)
    Else
        additionalPrice = "error"
    End If
    
End Function

Function firstPrice(pc As String, dbh As String) As Variant
    Dim codeDict As New Dictionary
    Dim priceDict As Dictionary
     
    Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 121
        .Add "S2", 127
        .Add "S3", 137
        .Add "S4", 150
        .Add "S5", 178
        .Add "S6", 206
        .Add "S7", 281
    End With
    codeDict.Add "PC1", priceDict
    
    Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 121
        .Add "S2", 127
        .Add "S3", 137
        .Add "S4", 150
        .Add "S5", 178
        .Add "S6", 206
        .Add "S7", 281
    End With
    codeDict.Add "PC2", priceDict
    
    Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 144
        .Add "S2", 155
        .Add "S3", 166
        .Add "S4", 178
        .Add "S5", 212
        .Add "S6", 246
        .Add "S7", 327
    End With
    codeDict.Add "PC3", priceDict
    
    Set priceDict = New Dictionary
    With priceDict
        .Add "S1", 160
        .Add "S2", 166
        .Add "S3", 178
        .Add "S4", 189
        .Add "S5", 224
        .Add "S6", 264
        .Add "S7", 249
    End With
    codeDict.Add "PC4", priceDict
    
    If codeDict.Exists(pc) Then
        firstPrice = codeDict(pc)(dbh)
    Else
        firstPrice = "error"
    End If
    
End Function
