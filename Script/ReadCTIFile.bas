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

Dim Pu As Double
Dim Mux As Double
Dim Muy As Double

Dim ICLC As Double
Dim ICfile As String

Dim MaxP As Double
MaxP = 0
Dim MinP As Double
MinP = 0

Dim MaxPMx As Double
Dim MaxPMy As Double
Dim MaxPLC As Double
Dim MaxPfile As String

Dim MinPMx As Double
Dim MinPMy As Double
Dim MinPLC As Double
Dim MinPfile As String

Dim MaxMux As Double
MaxMux = 0
Dim MaxMuxP As Double
Dim MaxMuxMy As Double
Dim MaxMuxLC As Double
Dim MaxMuxfile As String

Dim MaxMuy As Double
MaxMuy = 0
Dim MaxMuyP As Double
Dim MaxMuyMx As Double
Dim MaxMuyLC As Double
Dim MaxMuyfile As String

Dim MaxPIC As Double
Dim MinPIC As Double
Dim MaxMuxIC As Double
Dim MaxMuyIC As Double

Dim comma As String

Dim WS As Worksheet

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

For Each WS In Worksheets
    If WS.Name = "Results" Then
        Application.DisplayAlerts = False
        Worksheets("Results").Delete
        Application.DisplayAlerts = True
    Exit For
    End If
Next WS
Worksheets.Add.Name = "Results"

Worksheets("Results").Cells(1, 1).Value = "Combination Method:"
Worksheets("Results").Cells(2, 1).Value = "T&B Reinforcement:"
Worksheets("Results").Cells(3, 1).Value = "Side Reinforcement:"
Worksheets("Results").Cells(4, 1).Value = "IC & Controlling PCA Load & Input file:"
Worksheets("Results").Cells(5, 1).Value = "Max P Load Triplet & Controlling PCA Load & Input file:"
Worksheets("Results").Cells(6, 1).Value = "Min P Load Triplet & Controlling PCA Load & Input file:"
Worksheets("Results").Cells(7, 1).Value = "Max M2 Load Triplet & Controlling PCA Load & Input file:"
Worksheets("Results").Cells(8, 1).Value = "Max M3 Load Triplet & Controlling PCA Load & Input file:"
With Worksheets("Results").Range("A1:A10")
    .Font.FontStyle = "Bold"
End With
With Worksheets("Results").Range("A1:E1")
    .Font.FontStyle = "Bold"
End With




For j = 1 To 4
    '[Find the range of column groups]
    With Worksheets("Main")
        .Select
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
    MinIC = 999999
    MaxPIC = 999999
    MinPIC = 999999
    MaxMuxIC = 999999
    MaxMuyIC = 999999

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
            
            If Len(lineStr) = 86 Then
                'Read in Load and IC information
                IC = CDbl(Mid(lineStr, 70, 9))
                Pu = CDbl(Mid(lineStr, 10, 12))
                Mux = CDbl(Mid(lineStr, 22, 12))
                Muy = CDbl(Mid(lineStr, 34, 12))
                LC = Mid(lineStr, 1, 9)
                
                'Check if controlling IC
                If IC < MinIC Then
                    MinIC = IC
                    ICLC = LC
                    ICfile = ctiPath
                End If
                'Check if max P, min P, max M2x, or Max My
                If Pu >= MaxP And IC < MaxPIC Then
                    MaxP = Pu
                    MaxPMx = Mux
                    MaxPMy = Muy
                    MaxPLC = LC
                    MaxPfile = ctiPath
                    MaxPIC = IC
                End If
                If Pu <= MinP And IC < MinPIC Then
                    MinP = Pu
                    MinPMx = Mux
                    MinPMy = Muy
                    MinPLC = LC
                    MinPfile = ctiPath
                    MinPIC = IC
                End If
                If Abs(Mux) >= MaxMux And IC < MaxMuxIC Then
                    MaxMux = Abs(Mux)
                    MaxMuxP = Pu
                    MaxMuxMy = Abs(Muy)
                    MaxMuxLC = LC
                    MaxMuxfile = ctiPath
                    MaxMuxIC = IC
                End If
                If Abs(Muy) >= MaxMuy And IC < MaxMuyIC Then
                    MaxMuy = Abs(Muy)
                    MaxMuyP = Pu
                    MaxMuyMx = Abs(Mux)
                    MaxMuyLC = LC
                    MaxMuyfile = ctiPath
                    MaxMuyIC = IC
                End If
            
            End If
        Loop
        Close #Fnum1
    
    Next i 'Next CTI Files
    
    Worksheets("Main").Cells(i + 3, j).Value = MaxTBBars & ", " & MaxTBSteel & " in^2, T&B"
    Worksheets("Main").Cells(i + 4, j).Value = MaxSideBars & ", " & MaxSideSteel & " in^2, each side"
    Worksheets("Main").Cells(i + 5, j).Value = MinIC & ", " & ICLC & ", " & ICfile

    'Also print result in columns E, F, & G for convenience
    Worksheets("Main").Cells(1, j + 4).Value = Worksheets("Main").Cells(1, j).Value
    Worksheets("Main").Cells(3, j + 4).Value = MaxTBBars & ", " & MaxTBSteel & " in^2, T&B"
    Worksheets("Main").Cells(4, j + 4).Value = MaxSideBars & ", " & MaxSideSteel & " in^2, each side"
    Worksheets("Main").Cells(5, j + 4).Value = MinIC & ", " & ICLC & ", " & ICfile
    
    comma = ", "
    
    Worksheets("Results").Cells(1, j + 1).Value = Worksheets("Main").Cells(1, j).Value
    Worksheets("Results").Cells(2, j + 1).Value = MaxTBBars & comma & MaxTBSteel & " in^2"
    Worksheets("Results").Cells(3, j + 1).Value = MaxSideBars & comma & MaxSideSteel & " in^2"
    Worksheets("Results").Cells(4, j + 1).Value = MinIC & comma & ICLC & comma & ICfile
    Worksheets("Results").Cells(5, j + 1).Value = MaxP & comma & MaxPMx & comma & MaxPMy & comma & MaxPLC & comma & MaxPfile
    Worksheets("Results").Cells(6, j + 1).Value = MinP & comma & MinPMx & comma & MinPMy & comma & MinPLC & comma & MinPfile
    Worksheets("Results").Cells(7, j + 1).Value = MaxMuxP & comma & MaxMux & comma & MaxMuxMy & comma & MaxMuxLC & comma & MaxMuxfile
    Worksheets("Results").Cells(8, j + 1).Value = MaxMuyP & comma & MaxMuyMx & comma & MaxMuy & comma & MaxMuyLC & comma & MaxMuyfile
    
    'Reset variables for next combination type
    MaxTBBars = ""
    MaxTBSteel = 0
    MaxSideBars = ""
    MaxSideSteel = 0
    ICLC = 0
    ICfile = ""
    MaxP = 0
    MinP = 0
    MaxPMx = 0
    MaxPMy = 0
    MaxPLC = 0
    MaxPfile = ""
    MinPMx = 0
    MinPMy = 0
    MinPLC = 0
    MinPfile = ""
    MaxMux = 0
    MaxMuxP = 0
    MaxMuxMy = 0
    MaxMuxLC = 0
    MaxMuxfile = ""
    MaxMuy = 0
    MaxMuyP = 0
    MaxMuyMx = 0
    MaxMuyLC = 0
    MaxMuyfile = ""
    
Next j 'next combination type

With Worksheets("Results")
    .Columns("A:F").AutoFit
    .Select
End With


End Sub






