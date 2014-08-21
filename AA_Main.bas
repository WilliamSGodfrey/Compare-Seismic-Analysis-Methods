Attribute VB_Name = "AA_Main"
Option Explicit

Public strArray(1 To 86) As String          '[An array of strings used to create the
                                             'CTI input file]
Public Ret As Long
Public fpc As Double                        '[Concrete strength]
Public fy As Double                         '[Yield strength of reinforcement]

Sub Main()

Dim NumBmGroups As Single                   'This is the number of beams in the SAP model
Dim BmGrpNm As String                       'This is the name of the current beam group being post processed
Dim NumFrames As Double                     'This is the number of frames in the current frame group
Dim NumTS As Double                         'This is the number of time steps in the input time histories
Dim NumStaticLC As Single                   'This is the number of different static load cases in the SAP model
Dim StaticLC() As String                    'This stores the names of the static load cases in the SAP model
Dim DEADForces() As Double                  'This stores the forces from the static DEAD load cases
Dim LIVEForces() As Double                  'This stores the forces from the static LIVE load cases
Dim EQForcesSRSS() As Double                'This stores the forces from the EQ load cases combines by SRSS
Dim EQForcesHund() As Double                'This stores the forces from the EQ load cases combines by 100-40-40
Dim EQForcesASUM() As Double                'This stores the forces from the EQ load cases combines by ASUM
Dim ModelPath As String                     'This is the path to the SAP model
Dim NumEQTypes As Single                    'This is the number of different seismic load types (e.g. Time history, equivalent static, response spectra, etc.)
Dim NumEQComboType As Single                'This is the number of seismic effect combination methods
Dim tempCombos As String                    'Used temporarily to split string of seismic combination string
Dim CombineEQList As Dictionary             'This stores the method used to combine the seismic effects for each seismic load type (e.g. SRSS, algebraic sum, 100-40-40, etc.)
Dim tempEQLC As String                      'Used to temporarily store string that will be split into EQ LCs
Dim EQLC() As String                        'This stores the EQ load case names

Dim Key As String                           'Used to add entries to dictionaries

Dim NumObjs As Long                         'Required for SAP OAPI, populated with number of elements in SAP group
Dim ObjType() As Long                       'Required for SAP OAPI, populated with object type for elected object
Dim ObjIDs() As String                      'Required for SAP OAPI, populated with element IDs of selected objects

Dim PosNeg() As Integer                     'an array that is (1,-1), used to permute forces
Dim HunForFor() As Single                   'an array that is (1, .4, .4, 1, .4), used to permute forces

Dim EQAnalysis As String                    'A string that stores the type of EQ analysis

Dim NumberResults As Long
Dim Obj() As String
Dim ObjSta() As Double
Dim Elm() As String
Dim ElmSta() As Double
Dim LoadCase() As String
Dim StepType() As String
Dim StepNum() As Double

Dim P() As Double
Dim V2() As Double
Dim V3() As Double
Dim T() As Double
Dim M2() As Double
Dim M3() As Double

Dim P_EQ1() As Double
Dim V2_EQ1() As Double
Dim V3_EQ1() As Double
Dim T_EQ1() As Double
Dim M2_EQ1() As Double
Dim M3_EQ1() As Double
Dim P_EQ2() As Double
Dim V2_EQ2() As Double
Dim V3_EQ2() As Double
Dim T_EQ2() As Double
Dim M2_EQ2() As Double
Dim M3_EQ2() As Double
Dim P_EQ3() As Double
Dim V2_EQ3() As Double
Dim V3_EQ3() As Double
Dim T_EQ3() As Double
Dim M2_EQ3() As Double
Dim M3_EQ3() As Double

Dim P_EQcom As Double
Dim V2_EQcom As Double
Dim V3_EQcom As Double
Dim T_EQcom As Double
Dim M2_EQcom As Double
Dim M3_EQcom As Double

Dim PDL As Double
Dim PLL As Double
Dim M2DL As Double
Dim M2LL As Double
Dim M3DL As Double
Dim M3LL As Double
Dim tempDL()
Dim tempLL()

Dim TempArray As Variant

Dim numSRSSCTIFiles As Long
Dim numHundCTIFiles As Long
Dim numASUMCTIFiles As Long

Dim Fnum2 As Long
Dim Fnum3 As Long

Dim StartIndex As Long
Dim EndIndex As Long

Dim CTICount As Long

Dim BeamGroup_i As Single                   'Counter
Dim i As Single                             'Counter
Dim j As Single                             'Counter
Dim ii As Single                            'Counter
Dim jj As Single                            'Counter
Dim k As Single                             'Counter
Dim kk As Single                            'Counter
Dim q As Double                             'Counter
Dim qq As Double                            'Counter
Dim r As Double                             'Counter
Dim EQIndex As Double                       'Counter

Dim SRSScounter As Double                   'Counter
Dim Hundcounter As Double                   'Counter
Dim ASUMcounter As Double                   'Counter

Dim BmWidth As Long
Dim BmHeight As Long


'[Initialize batch file]
Fnum3 = FreeFile()
Open ActiveWorkbook.Path & "\pcaBatchFile.bat" For Output As #Fnum3
Close #Fnum3


'Open SAP model and set units
ModelPath = Worksheets("Input").Cells(2, 3)
BB_OpenSAP.Open_SAP_Model (ModelPath)

'Initialize
ReDim PosNeg(1 To 2)
PosNeg(1) = 1
PosNeg(2) = -1
ReDim HunForFor(1 To 5)
HunForFor(1) = 1
HunForFor(2) = 0.4
HunForFor(3) = 0.4
HunForFor(4) = 1
HunForFor(5) = 0.4


'Count some stuff
NumBmGroups = Worksheets("Group Def").Range("B1").End(xlDown).Row - 1
NumEQTypes = Worksheets("Input").Range("I1").End(xlDown).Row - 1

'Store Static LC names
NumStaticLC = Worksheets("Input").Range("G1").End(xlDown).Row - 1
ReDim StaticLC(0 To NumStaticLC - 1)
For j = 0 To NumStaticLC - 1
    StaticLC(j) = Worksheets("Input").Cells(2 + j, 7).Value
Next j

CTICount = 0
Worksheets("Main").Range("A2:C65536").Clear


For BeamGroup_i = 1 To NumBmGroups
    
    BmGrpNm = Worksheets("Group Def").Cells(1 + BeamGroup_i, 2).Value
    fpc = Worksheets("Group Def").Cells(1 + BeamGroup_i, 4) * 1000
    fy = Worksheets("Group Def").Cells(1 + BeamGroup_i, 5) * 1000
    BmWidth = Worksheets("Group Def").Cells(1 + BeamGroup_i, 6)
    BmHeight = Worksheets("Group Def").Cells(1 + BeamGroup_i, 7)
    
   
    'Select current frame group
    Ret = SapObject.SapModel.SelectObj.ClearSelection
    Ret = SapObject.SapModel.SelectObj.Group(BmGrpNm, False) 'False = select objects, True = deselect objects

    'Set case and combo output selections for DEAD Load
    Ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
    Ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")
    
    'Get DEAD forces from SAP and store in array
    Ret = SapModel.Results.FrameForce(BmGrpNm, 2, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)
    ReDim DEADForces(0 To UBound(Elm), 0 To 2)
    For j = 0 To UBound(Elm)
        DEADForces(j, 0) = P(j)
        DEADForces(j, 1) = M2(j)
        DEADForces(j, 2) = M3(j)
    Next j

    'Set case and combo output selections for LIVE Load
    Ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
    Ret = SapModel.Results.Setup.SetCaseSelectedForOutput("LIVE")
    
    'Get LIVE forces from SAP and store in array
    Ret = SapModel.Results.FrameForce(BmGrpNm, 2, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)
    ReDim LIVEForces(0 To UBound(Elm), 0 To 2)
    For j = 0 To UBound(Elm)
        LIVEForces(j, 0) = P(j)
        LIVEForces(j, 1) = M2(j)
        LIVEForces(j, 2) = M3(j)
    Next j

    
    For i = 1 To NumEQTypes
        EQAnalysis = Worksheets("Input").Cells(1 + i, 9).Value
        'Store EQ load case names and pull forces from SAP
        tempEQLC = Worksheets("Input").Cells(1 + i, 10).Value
        ReDim EQLC(0 To 2)
        For j = 0 To 2
            EQLC(j) = Split(tempEQLC, ", ")(j)
        Next j

        'Store seismic effect combination methods
        Set CombineEQList = New Dictionary
        tempCombos = Worksheets("Input").Cells(1 + i, 11).Value
        NumEQComboType = UBound(Split(tempCombos))
        For j = 0 To NumEQComboType
            CombineEQList.Add Split(tempCombos, ", ")(j), ""
        Next j

        'set output time step option (Envelopes, each step, or last step)
        Ret = SapModel.Results.Setup.SetOptionModalHist(2) '2 is for each step
            
        'Set case and combo output selections
        Ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        Ret = SapModel.Results.Setup.SetCaseSelectedForOutput(EQLC(0))
        Ret = SapModel.Results.FrameForce(BmGrpNm, 2, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P_EQ1, V2_EQ1, V3_EQ1, T_EQ1, M2_EQ1, M3_EQ1)
        Ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        Ret = SapModel.Results.Setup.SetCaseSelectedForOutput(EQLC(1))
        Ret = SapModel.Results.FrameForce(BmGrpNm, 2, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P_EQ2, V2_EQ2, V3_EQ2, T_EQ2, M2_EQ2, M3_EQ2)
        Ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        Ret = SapModel.Results.Setup.SetCaseSelectedForOutput(EQLC(2))
        Ret = SapModel.Results.FrameForce(BmGrpNm, 2, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P_EQ3, V2_EQ3, V3_EQ3, T_EQ3, M2_EQ3, M3_EQ3)
        NumFrames = (UBound(DEADForces) + 1) / 3 ' Find the number of frames in the group
        NumTS = UBound(Obj) / NumFrames / 3 'Find the number of time steps
        SRSScounter = 0
        Hundcounter = 0
        ASUMcounter = 0
        
        'Initialize arrays that store combined forces if applicable to current analysis type
        If CombineEQList.Exists("SRSS") Then
            ReDim EQForcesSRSS(0 To NumberResults * 8 - 1, 0 To 2)
        End If
        If CombineEQList.Exists("100-40-40") Then
            ReDim EQForcesHund(0 To NumberResults * 24 - 1, 0 To 2)
        End If
        If CombineEQList.Exists("ASUM") Then
            ReDim EQForcesASUM(0 To NumberResults - 1, 0 To 2)
        End If
        
        For ii = 0 To NumFrames - 1 'Loop on the number of elements
            For r = 0 To NumTS ' Loop on the number of time steps
                For qq = 0 To 2 'Loop on the number of output stations

                    'Retrieve beam forces from Static Force Arrays
                    PDL = -DEADForces(3 * ii + qq, 0)
                    PLL = -LIVEForces(3 * ii + qq, 0)
                    M2DL = DEADForces(3 * ii + qq, 1)
                    M2LL = LIVEForces(3 * ii + qq, 1)
                    M3DL = DEADForces(3 * ii + qq, 2)
                    M3LL = LIVEForces(3 * ii + qq, 2)

                    EQIndex = (ii * (NumTS * 3) + r * 3 + qq) 'array index where forces for current frame, TS, and station are located
                    
                    If CombineEQList.Exists("SRSS") Then
                        For jj = 1 To 2
                            For k = 1 To 2
                                For kk = 1 To 2
                                    'Combine static beam forces with permuted seismic beam forces, use ACI349 9.2.1 LC4 only
                                    P_EQcom = PosNeg(jj) * (P_EQ1(EQIndex) ^ 2 + P_EQ2(EQIndex) ^ 2 + P_EQ3(EQIndex) ^ 2) ^ 0.5 + PDL + PLL
                                    M2_EQcom = PosNeg(k) * (M2_EQ1(EQIndex) ^ 2 + M2_EQ2(EQIndex) ^ 2 + M2_EQ3(EQIndex) ^ 2) ^ 0.5 + M2DL + M2LL
                                    M3_EQcom = PosNeg(kk) * (M3_EQ1(EQIndex) ^ 2 + M3_EQ2(EQIndex) ^ 2 + M3_EQ3(EQIndex) ^ 2) ^ 0.5 + M3DL + M3LL
                                    EQForcesSRSS(SRSScounter, 0) = P_EQcom
                                    EQForcesSRSS(SRSScounter, 1) = M2_EQcom
                                    EQForcesSRSS(SRSScounter, 2) = M3_EQcom
                                    SRSScounter = SRSScounter + 1
                                Next kk
                            Next k
                        Next jj
                    End If
                    If CombineEQList.Exists("100-40-40") Then
                        For jj = 1 To 2
                            For k = 1 To 2
                                For kk = 1 To 2
                                    For j = 1 To 3
                                        'Combine static beam forces with permuted seismic beam forces, use ACI349 9.2.1 LC4 only
                                        P_EQcom = (PosNeg(jj) * HunForFor(j) * -P_EQ1(EQIndex) + PosNeg(k) * HunForFor(j + 1) * -P_EQ2(EQIndex) + PosNeg(kk) * HunForFor(j + 2) * -P_EQ3(EQIndex)) + PDL + PLL
                                        M2_EQcom = (PosNeg(jj) * HunForFor(j) * M2_EQ1(EQIndex) + PosNeg(k) * HunForFor(j + 1) * M2_EQ2(EQIndex) + PosNeg(kk) * HunForFor(j + 2) * M2_EQ3(EQIndex)) + M2DL + M2LL
                                        M3_EQcom = (PosNeg(jj) * HunForFor(j) * M3_EQ1(EQIndex) + PosNeg(k) * HunForFor(j + 1) * M3_EQ2(EQIndex) + PosNeg(kk) * HunForFor(j + 2) * M3_EQ3(EQIndex)) + M3DL + M3LL
                                        EQForcesHund(Hundcounter, 0) = P_EQcom
                                        EQForcesHund(Hundcounter, 1) = M2_EQcom
                                        EQForcesHund(Hundcounter, 2) = M3_EQcom
                                        Hundcounter = Hundcounter + 1
                                        Next j
                                Next kk
                            Next k
                        Next jj
                    End If
                    If CombineEQList.Exists("ASUM") Then
                        'Combine static beam forces with permuted seismic beam forces, use ACI349 9.2.1 LC4 only
                        P_EQcom = (-P_EQ1(ii) + -P_EQ2(ii) + -P_EQ3(ii)) + PDL + PLL
                        M2_EQcom = (M2_EQ1(ii) + M2_EQ2(ii) + M2_EQ3(ii)) + M2DL + M2LL
                        M3_EQcom = (M3_EQ1(ii) + M3_EQ2(ii) + M3_EQ3(ii)) + M3DL + M3LL
                        TempArray = Array(P_EQcom, M2_EQcom, M3_EQcom)
                        EQForcesASUM(ASUMcounter, 0) = P_EQcom
                        EQForcesASUM(ASUMcounter, 1) = M2_EQcom
                        EQForcesASUM(ASUMcounter, 2) = M3_EQcom
                        ASUMcounter = ASUMcounter + 1
                    End If
                    
                Next qq 'Next output station
            Next r 'Next time step
        Next ii 'Next frame

        Dim NumberofSRSSResults As Double
        Dim NumberofHundResults As Double
        Dim NumberofASUMResults As Double
        
        'Write design force triplets, (P, M2, and M3), to PCA Column input file for design of beam section
        NumberofSRSSResults = SRSScounter
        NumberofHundResults = Hundcounter
        NumberofASUMResults = ASUMcounter

        '[Calculate the number of CTI files that will be needed for the given frame element]
        If CombineEQList.Exists("SRSS") Then
            If NumberofSRSSResults Mod 4500 = 0 Then
                numSRSSCTIFiles = Int(NumberofSRSSResults / 4500)
            Else
                numSRSSCTIFiles = Int(NumberofSRSSResults / 4500) + 1
            End If
        Else
            numSRSSCTIFiles = 0
        End If
        If CombineEQList.Exists("100-40-40") Then
            If NumberofHundResults Mod 4500 = 0 Then
                numHundCTIFiles = Int(NumberofHundResults / 4500)
            Else
                numHundCTIFiles = Int(NumberofHundResults / 4500) + 1
            End If
        Else
            numHundCTIFiles = 0
        End If
        If CombineEQList.Exists("ASUM") Then
            If NumberofASUMResults Mod 4500 = 0 Then
                numASUMCTIFiles = Int(NumberofASUMResults / 4500)
            Else
                numASUMCTIFiles = Int(NumberofASUMResults / 4500) + 1
            End If
        Else
            numASUMCTIFiles = 0
        End If
        'Create SRSS CTI files
        If numSRSSCTIFiles > 0 Then
            ReDim SRSSfilenameCTI(1 To numSRSSCTIFiles)
            ReDim numFactLoads(1 To numSRSSCTIFiles)
            For j = 1 To numSRSSCTIFiles
                SRSSfilenameCTI(j) = EQAnalysis & BmGrpNm & "-SRSS-" & j & "PCAInputFile.cti"

                '[Add CTI Files to PCA Batch File]
                Open ActiveWorkbook.Path & "\pcaBatchFile.bat" For Append As #Fnum3
                    Print #Fnum3, "cd " & ActiveWorkbook.Path
                    Print #Fnum3, "\\Snlvs5\sys3\Ops$\PCA198410\pcaColumn /i:" & SRSSfilenameCTI(j)
                Close #Fnum3

                If numSRSSCTIFiles = 1 Then                 '[Only one file]
                    numFactLoads(j) = NumberofSRSSResults
                ElseIf j <> numSRSSCTIFiles Then            '[It's not the last file
                    numFactLoads(j) = 4500                  'but more than one exist]
                ElseIf j = numSRSSCTIFiles Then             '[It's the last file]
                    numFactLoads(j) = NumberofSRSSResults - 4500 * (numSRSSCTIFiles - 1)
                End If

                CTICount = CTICount + 1
                Worksheets("Main").Cells(1 + j, 1) = SRSSfilenameCTI(j)

                '[Strings are populated with data that will be used to create the PCA CTI file]
                Call PCA.Strings(numFactLoads(j), BmWidth, BmHeight, "SRSS-File" & j & "/" & numSRSSCTIFiles)

                '[The CTI Input file is written below]
                Fnum2 = FreeFile()
                Open ActiveWorkbook.Path & "\" & SRSSfilenameCTI(j) For Output As #Fnum2

                StartIndex = (j - 1) * 4500
                EndIndex = ((j - 1) * 4500) + numFactLoads(j) - 1

                For q = 1 To 43
                    Print #Fnum2, strArray(q)
                Next q
                'Add loads to CTI input file
                For q = StartIndex To EndIndex
                    Print #Fnum2, EQForcesSRSS(q, 0) & "," & EQForcesSRSS(q, 1) & "," & EQForcesSRSS(q, 2)
                Next q

                For q = 44 To 86
                    Print #Fnum2, strArray(q)
                Next q
                Close #Fnum2
            Next j
        End If

        'Create ASUM CTI files
        If numASUMCTIFiles > 0 Then
            ReDim ASUMfilenameCTI(1 To numASUMCTIFiles)
            ReDim numFactLoads(1 To numASUMCTIFiles)
            For j = 1 To numASUMCTIFiles
                ASUMfilenameCTI(j) = EQAnalysis & BmGrpNm & "-ASUM-" & j & "PCAInputFile.cti"
                '[Add CTI Files to PCA Batch File]
                Open ActiveWorkbook.Path & "\pcaBatchFile.bat" For Append As #Fnum3
                    Print #Fnum3, "cd " & ActiveWorkbook.Path
                    Print #Fnum3, "\\Snlvs5\sys3\Ops$\PCA198410\pcaColumn /i:" & ASUMfilenameCTI(j)
                Close #Fnum3

                If numASUMCTIFiles = 1 Then                 '[Only one file]
                    numFactLoads(j) = NumberofASUMResults
                ElseIf j <> numASUMCTIFiles Then            '[It's not the last file
                    numFactLoads(j) = 4500                  'but more than one exist]
                ElseIf j = numASUMCTIFiles Then             '[It's the last file]
                    numFactLoads(j) = NumberofASUMResults - 4500 * (numASUMCTIFiles - 1)
                End If

                CTICount = CTICount + 1
                Worksheets("Main").Cells(1 + j, 2) = ASUMfilenameCTI(j)

                '[Strings are populated with data that will be used to create the PCA CTI file]
                Call PCA.Strings(numFactLoads(j), BmWidth, BmHeight, "ASUM-File" & j & "/" & UBound(ASUMfilenameCTI))

                '[The CTI Input file is written below]
                Fnum2 = FreeFile()
                Open ActiveWorkbook.Path & "\" & ASUMfilenameCTI(j) For Output As #Fnum2

                StartIndex = (j - 1) * 4500
                EndIndex = ((j - 1) * 4500) + numFactLoads(j) - 1

                For q = 1 To 43
                    Print #Fnum2, strArray(q)
                Next q
                'Add loads to CTI input file
                For q = StartIndex To EndIndex
                    Print #Fnum2, EQForcesASUM(q, 0) & "," & EQForcesASUM(q, 1) & "," & EQForcesASUM(q, 2)
                Next q

                For q = 44 To 86
                    Print #Fnum2, strArray(q)
                Next q
                Close #Fnum2
            Next j
        End If

        'Create 100-40-40 CTI files
        If numHundCTIFiles > 0 Then
            ReDim HundfilenameCTI(1 To numHundCTIFiles)
            ReDim numFactLoads(1 To numHundCTIFiles)
            For j = 1 To numHundCTIFiles
                HundfilenameCTI(j) = EQAnalysis & BmGrpNm & "-Hund-" & j & "PCAInputFile.cti"
                '[Add CTI Files to PCA Batch File]
                Open ActiveWorkbook.Path & "\pcaBatchFile.bat" For Append As #Fnum3
                    Print #Fnum3, "cd " & ActiveWorkbook.Path
                    Print #Fnum3, "\\Snlvs5\sys3\Ops$\PCA198410\pcaColumn /i:" & HundfilenameCTI(j)
                Close #Fnum3

                If numHundCTIFiles = 1 Then                 '[Only one file]
                    numFactLoads(j) = NumberofHundResults
                ElseIf j <> numHundCTIFiles Then            '[It's not the last file
                    numFactLoads(j) = 4500                  'but more than one exist]
                ElseIf j = numHundCTIFiles Then             '[It's the last file]
                    numFactLoads(j) = NumberofHundResults - 4500 * (numHundCTIFiles - 1)
                End If

                CTICount = CTICount + 1
                Worksheets("Main").Cells(1 + j, 3) = HundfilenameCTI(j)

                '[Strings are populated with data that will be used to create the PCA CTI file]
                Call PCA.Strings(numFactLoads(j), BmWidth, BmHeight, "Hund-File" & j & "/" & UBound(HundfilenameCTI))

                '[The CTI Input file is written below]
                Fnum2 = FreeFile()
                Open ActiveWorkbook.Path & "\" & HundfilenameCTI(j) For Output As #Fnum2

                StartIndex = (j - 1) * 4500
                EndIndex = ((j - 1) * 4500) + (numFactLoads(j) - 1)

                'Add general information from module "PCA"
                For q = 1 To 43
                    Print #Fnum2, strArray(q)
                Next q
                'Add loads to CTI input file
                For q = StartIndex To EndIndex
                    Print #Fnum2, EQForcesHund(q, 0) & "," & EQForcesHund(q, 1) & "," & EQForcesHund(q, 2)
                Next q

                'Add the remainder of the general information from module "PCA"
                For q = 44 To 86
                    Print #Fnum2, strArray(q)
                Next q
                Close #Fnum2
            Next j
        End If
    Next i ' next seismic analysis
Next BeamGroup_i


'Close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing


End Sub
