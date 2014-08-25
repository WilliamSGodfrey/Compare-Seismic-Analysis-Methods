Attribute VB_Name = "ReadCTIFile"
Option Explicit

Sub ReadCTIFile()

Dim bararea As Dictionary

Dim j As Long
Dim i As Long
Dim groupList As Range                          '[List of groups]
Dim NumFiles As Long
Dim MaxTBSteel As Double
Dim MaxSideSteel As Double
Dim ctiPath As String

Dim lineStr As String                           '[Generic String information]
Dim sizeStr As String                           '[String containing section size information]
Dim barsStr As String                           '[String containing rebar information]
Dim coverStr As String                          '[String containing cover information]
Dim sectInfo As String                       '[String containing cover information]
Dim aosInfo As String                        '[String containing cover information]
Dim numTBBars As Double
Dim sizeTBBars As Double
Dim numSideBars As Double
Dim sizeSideBars As Double
Dim AreaTBSteel As Double
Dim AreaSideSteel As Double
Dim MaxTBBars As String
Dim MaxSideBars As String
Dim Fnum1 As Long

Dim MinIC As Double
Dim IC As Double
Dim LC As String
Dim file As String


Set bararea = New Dictionary

bararea.Add 3, 0.11
bararea.Add 4, 0.2
bararea.Add 5, 0.31
bararea.Add 6, 0.44
bararea.Add 7, 0.6
bararea.Add 8, 0.79
bararea.Add 9, 1
bararea.Add 10, 1.27
bararea.Add 11, 1.56
bararea.Add 14, 2.25


For j = 1 To 3
    '[Find the range of column groups]
    With Worksheets("Main")
        .Cells(1, j).Select
        If .Cells(2, j) <> "" Then .Range(Selection, Selection.End(xlDown)).Select
        Set groupList = Selection
        .Range("A1").Select
    End With
    
    NumFiles = groupList.Count - 1
    
    MaxTBBars = ""
    MaxSideBars = ""
    MaxTBSteel = 0
    MaxSideSteel = 0
    MinIC = 9999999

    For i = 0 To NumFiles - 1
        ctiPath = Split(Worksheets("Main").Cells(2 + i, j).Value, ".cti")(0) & ".out"
        Fnum1 = FreeFile()
        Open ActiveWorkbook.Path & "\" & ctiPath For Input As #Fnum1
    
        Do While Not EOF(1)
    
            Line Input #Fnum1, lineStr
            '[Reads in the section size]
            If lineStr = "   Section:" Then
                Line Input #Fnum1, lineStr
                Line Input #Fnum1, sizeStr
                sectInfo = sizeStr
            End If
    
    '        '[Reads in the tie information]
    '        If Left(lineStr, 17) = "      Confinement" Then
    '            tieSize1 = Mid(lineStr, 27, 1)
    '            tieSize2 = Mid(lineStr, 51, 1)
    '        End If
    
            '[Reads in the rebar information]
            If lineStr = "      Layout: Rectangular" Then
                Line Input #Fnum1, lineStr
                Line Input #Fnum1, aosInfo
                Line Input #Fnum1, lineStr
                Line Input #Fnum1, lineStr
                Line Input #Fnum1, barsStr
                Line Input #Fnum1, coverStr
                
                numTBBars = CDbl(Mid(barsStr, 22, 2))
                sizeTBBars = CDbl(Mid(barsStr, 27, 2))
                numSideBars = CDbl(Mid(barsStr, 48, 2))
                sizeSideBars = CDbl(Mid(barsStr, 53, 2))
                AreaTBSteel = numTBBars * bararea.Item(sizeTBBars)
                AreaSideSteel = numSideBars * bararea.Item(sizeSideBars)
                
                
                If AreaTBSteel > MaxTBSteel Then
                    MaxTBSteel = AreaTBSteel
                    MaxTBBars = numTBBars & " #" & sizeTBBars
                End If
                If AreaSideSteel > MaxSideSteel Then
                    MaxSideSteel = AreaSideSteel
                    MaxSideBars = numSideBars & " #" & sizeSideBars
                End If
                'Exit Do
            End If
            
            'Read in IC information
            If Len(lineStr) = 86 Then
                IC = CDbl(Mid(lineStr, 70, 9))
                If IC < MinIC Then
                    MinIC = IC
                    LC = Mid(lineStr, 1, 9)
                    file = ctiPath
                End If
            End If
        Loop
        Close #Fnum1
    
    Next i 'Next CTI Files
    
    Worksheets("Main").Cells(i + 3, j).Value = MaxTBBars & ", " & MaxTBSteel & " in^2, T&B"
    Worksheets("Main").Cells(i + 4, j).Value = MaxSideBars & ", " & MaxSideSteel & " in^2, each side"
    Worksheets("Main").Cells(i + 5, j).Value = MinIC & ", " & LC & ", " & file

    'Also print result in columns E, F, & G for convenience
    Worksheets("Main").Cells(1, j + 4).Value = Worksheets("Main").Cells(1, j).Value
    Worksheets("Main").Cells(3, j + 4).Value = MaxTBBars & ", " & MaxTBSteel & " in^2, T&B"
    Worksheets("Main").Cells(4, j + 4).Value = MaxSideBars & ", " & MaxSideSteel & " in^2, each side"
    Worksheets("Main").Cells(5, j + 4).Value = MinIC & ", " & LC & ", " & file
    


Next j 'next combination type

End Sub






