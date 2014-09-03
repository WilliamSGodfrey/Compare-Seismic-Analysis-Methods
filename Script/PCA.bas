Attribute VB_Name = "PCA"


Option Explicit
'----------------------------------------------------------------------------------------------
'The following code is used to create the CTI input file for PCA Column.
'If any row of data is not explained in detail, it can be found in the PCA Column Manual.
'For user convenience, part of the manual (For v4.1)
'can be found in the CTIInfoPCAColumnVersion4_1 module.]
'----------------------------------------------------------------------------------------------

Sub Strings(ByVal numberOfLoads As Long, ByVal columnWidth As Double, ByVal columnDepth As Double, ByVal strGroupName As String, ByVal Line As Double)

strArray(1) = "#pcaColumn Text Input (CTI) File:"
strArray(2) = "[pcaColumn version]"
strArray(3) = "4.100"
strArray(4) = "[Project]"
strArray(5) = Worksheets("Main").Range("k9")
strArray(6) = "[Column ID]"
strArray(7) = strGroupName
strArray(8) = "[Engineer]"
strArray(9) = Worksheets("Main").Range("k11")
strArray(10) = "[Investigation Run Flag]"
strArray(11) = "15"
strArray(12) = "[Design Run Flag]"
strArray(13) = "15"
strArray(14) = "[Slenderness Flag]"
strArray(15) = "31"
strArray(16) = "[User Options]"


'---------------------------------------------------------------------------------------------
'[Information below is for user options]
'---------------------------------------------------------------------------------------------
Dim investigationMode As Long
Dim units As Long
Dim code As Long
Dim axisRun As Long
Dim reservedDNE1 As Long
Dim slenderness As Long
Dim designType As Long
Dim reservedDNE2 As Long
Dim sectionType As Long
Dim barLayout As Long
Dim columnType As Long
Dim confinement As Long
Dim ltInvestigation As Long
Dim ltDesign As Long
Dim rLInvestigation As Long
Dim rLDesign As Long
Dim reservedDNE3 As Long
Dim numFactLoads As Long
Dim numServLoads As Long
Dim numPtsExtCol As Long
Dim numPtsIntSec As Long
Dim reservedDNE4 As Long
Dim reservedDNE5 As String
Dim coverInvest As Long
Dim coverDesign As Long
Dim numLC As Long
Dim comma As String
Dim i As Double
Dim k As Double


investigationMode = 1   '[0=Investigation Mode,1=Design Mode]
units = 0               '[0-English Unit; 1-Metric Units]
code = 2                '[0-ACI 318-02; 1- CSA A23.3-94; 2-ACI 318-05; 3-CSA A23.3-04]
axisRun = 2             '[0-X Axis Run; 1-Y Axis Run; 2-Biaxial Run]
reservedDNE1 = 0        '[Reserved. Do not edit]
slenderness = 0         '[0-Slenderness is not considered; 1-Slenderness in considered]
designType = 0          '[0-Design for minimum number of bars; 1-Design for minimum area
                         'of reinforcement]
reservedDNE2 = 0        '[Reserved. Do not edit]
sectionType = 0         '[0-Rectangular Column Section; 1-Circular Column Section;
                         '2-Irregular Column Section]
barLayout = 0           '[0-Rectangular reinforcing bar layout; 1-Circular reinforcing bar layout]
columnType = 0          '[0-Structural Column Section; 1-Architectural Column Section;
                         '2-User Defined Column Section;]
confinement = 0         '[0-Tied Confinement; 1-Spiral Confinement; 2-Other Confinement
ltInvestigation = -1    '[0-Factored; 1-Service; 2-Control Points; 3-Axial Loads
                        '-1 Since using Design Mode]
ltDesign = 0            '[0-Factored; 1-Service; 2-Control Points; 3-Axial Loads]
rLInvestigation = 2     '[0-All Side Equal; 1-Equal Spacing; 2-Sides Different; 3-Irregular Pattern]
rLDesign = 2            '[0-All Side Equal; 1-Equal Spacing; 2-Sides Different; 3-Irregular Pattern]
reservedDNE3 = 8        '[Reserved. Do not edit!!!!!!!!!!!!!!!!!!!!!!!!!!]
numFactLoads = numberOfLoads    '[Number of factored loads]
numServLoads = 0        '[Number of service loads]
numPtsExtCol = 0        '[Number of points on exterior column section]
numPtsIntSec = 0        '[Number of points on interior section opening]
reservedDNE4 = 0        '[Reserved. Do not edit]
reservedDNE5 = "0.000000"    '[Reserved. Do not edit]
coverInvest = 0         '[Cover type for investigation mode: 0-To transverse bar;
                         '1-To longitudinal bar]
coverDesign = 0         '[Cover type for design mode: 0-To transverse bar; 1-To longitudinal bar]
numLC = 13              '[Number of load combinations (Default value, NOT USED)]

comma = ","             '[Separator]
'-----------------------------------------------------------------------------------------------

strArray(17) = investigationMode & comma & units & comma & code & comma & axisRun & _
               comma & reservedDNE1 & comma & slenderness & comma & designType & comma & _
               reservedDNE2 & comma & sectionType & comma & barLayout & comma & _
               columnType & comma & confinement & comma & ltInvestigation & comma & _
               ltDesign & comma & rLInvestigation & comma & rLDesign & comma & _
               reservedDNE3 & comma & numFactLoads & comma & numServLoads & comma & _
               numPtsExtCol & comma & numPtsIntSec & comma & reservedDNE4 & comma & _
               reservedDNE5 & comma & coverInvest & comma & coverDesign & comma & numLC

strArray(18) = "[Irregular Options]"
'[strArray(19) is for irregular options.  Default string is used
strArray(19) = "-2,0,0,1,0.600000,50.000000,50.000000,-50.000000"
strArray(19) = strArray(19) & ",-50.000000,0.000000,0.000000,5.000000,5.000000"
strArray(20) = "[Ties]"
'[There are 3 values separated by commas in one line in this section. These values
 'are described below in the order they appear from left to right. (Menu Input |
 'Reinforcement | Confinement?)]
'[1. Index (0 based) of tie bars for longitudinal bars smaller that the one
 'specified in the 3rd item in this section in the drop-down list]
'[2. Index (0 based) of tie bars for longitudinal bars bigger that the one specified
 'in the 3rd item in this section in the drop-down list]
'[3. Index (0 based) of longitudinal bar in the drop-down list]
'If tieBar = 0 Then
    strArray(21) = "0,1,7"
'Else
'    strArray(21) = CStr(tieBar) & comma & CStr(tieBar) & comma & "7"
'End If
strArray(22) = "[Investigation Reinforcement]"
'[strArray(22)  is for investigation option. Default string is used]
strArray(23) = "0,0,0,0,0,0,0,0,0.000000,0.000000,0.000000,0.000000"
strArray(24) = "[Design Reinforcement]"

'-----------------------------------------------------------------------------------------------
'[Information below is for design reinforcement]
'-----------------------------------------------------------------------------------------------
Dim drArray(1 To 12) As String

If Worksheets("Input").Cells(1 + Line, 12) = "NO" Then

    drArray(1) = 2 * Worksheets("Input").Cells(1 + Line, 13)        '[Minimum number of top and bottom bars]
    drArray(2) = 2 * Worksheets("Input").Cells(1 + Line, 13)        '[Maximum number of top and bottom bars]
    drArray(3) = 2 * Worksheets("Input").Cells(1 + Line, 14)        '[Minimum number of left and right bars]
    drArray(4) = 2 * Worksheets("Input").Cells(1 + Line, 14)        '[Maximum number of left and right bars]
    
    '[0=#3bar,1=#4bar,2=#5bar,3=#6bar, ... ,7=#10bar,8=#11bar]
    
    drArray(5) = Worksheets("Input").Cells(1 + Line, 15) - 3         '[Index (0 based) of minimum size for top and bottom bars]
    drArray(6) = Worksheets("Input").Cells(1 + Line, 15) - 3         '[Index (0 based) of maximum size for top and bottom bars]
    drArray(7) = Worksheets("Input").Cells(1 + Line, 16) - 3         '[Index (0 based) of minimum size for left and right bars]
    drArray(8) = Worksheets("Input").Cells(1 + Line, 16) - 3         '[Index (0 based) of maximum size for left and right bars
Else

    drArray(1) = "4"           '[Minimum number of top and bottom bars]
    drArray(2) = "12"          '[Maximum number of top and bottom bars]
    drArray(3) = "4"           '[Minimum number of left and right bars]
    drArray(4) = "12"          '[Maximum number of left and right bars]
    
    '[0=#3bar,1=#4bar,2=#5bar,3=#6bar, ... ,7=#10bar,8=#11bar]
    
    drArray(5) = "3"           '[Index (0 based) of minimum size for top and bottom bars]
    drArray(6) = "8"           '[Index (0 based) of maximum size for top and bottom bars]
    drArray(7) = "3"           '[Index (0 based) of minimum size for left and right bars]
    drArray(8) = "8"           '[Index (0 based) of maximum size for left and right bars]
End If

    drArray(9) = "2.000000"    '[Clear cover to top and bottom bars]
    drArray(10) = "2.000000"   '[Reserved. Do not edit.]
    drArray(11) = "2.000000"   '[Clear cover to left and right bars]
    drArray(12) = "2.000000"   '[Reserved. Do not edit.]

'-----------------------------------------------------------------------------------------------
strArray(25) = ""
For i = 1 To 11
    strArray(25) = strArray(25) & drArray(i) & comma
Next i

strArray(25) = strArray(25) & drArray(12)

strArray(26) = "[Investigation Section Dimensions]"
'[strArray(27)  is for investigation option. Default string is used]
strArray(27) = "0.000000,0.000000"
strArray(28) = "[Design Section Dimensions]"

'-----------------------------------------------------------------------------------------------
Dim xColStartDim As String
Dim xColEndDim As String
Dim yColStartDim As String
Dim yColEndDim As String
Dim xIncr As String
Dim yIncr As String

'[1. Section width (along X) Start]
xColStartDim = CStr(Format(columnWidth * 12, "0.000000"))
'[2. Section depth (along Y) Start]
yColStartDim = CStr(Format(columnDepth * 12, "0.000000"))
'[3. Section width (along X) End]
xColEndDim = CStr(Format(columnWidth * 12, "0.000000"))
'[4. Section depth (along Y) End]
yColEndDim = CStr(Format(columnDepth * 12, "0.000000"))
'[5. Section width (along X) Increment]
xIncr = CStr(Format(CDbl(xColEndDim) - CDbl(xColStartDim), "0.000000"))
'[6. Section depth (along Y) Increment]
yIncr = CStr(Format(CDbl(yColEndDim) - CDbl(yColStartDim), "0.000000"))

'-----------------------------------------------------------------------------------------------
strArray(29) = xColStartDim & comma & yColStartDim & comma & xColEndDim & comma & yColEndDim & _
               comma & xIncr & comma & yIncr

strArray(30) = "[Material Properties]"
'-----------------------------------------------------------------------------------------------

Dim Ec As String
Dim fpc1 As String
Dim fy1 As String
Dim fc As String
Dim beta1 As Single
Dim b1 As String
Dim esu As String
Dim Es As String
Dim rDNE As String

Ec = CStr(Format(57 * fpc ^ 0.5, "0.000000"))       '[Concrete modulus of elasticity, Ec (ksi)]
fc = CStr(Format(0.85 * fpc / 1000, "0.000000"))    '[Concrete maximum stress, fc (ksi)]
beta1 = 0.85 - 0.05 * (fpc / 1000 - 4)
If beta1 > 0.85 Then beta1 = 0.85
If beta1 < 0.65 Then beta1 = 0.65
fpc1 = CStr(Format(fpc / 1000, "0.000000"))         '[Concrete strength, f?c (ksi)]
b1 = CStr(Format(beta1, "0.000000"))                '[Beta(1) for concrete stress block]
esu = "0.003000"                                    '[Concrete ultimate strain (in/in)]
fy1 = CStr(Format(fy / 1000, "0.000000"))           '[Steel yield strength, fy (ksi)]
Es = "29000.000000"                                 '[Steel modulus of elasticity, Es (ksi)]
rDNE = "0"                                          '[Reserved. Do not edit.]
'-----------------------------------------------------------------------------------------------
strArray(31) = fpc1 & comma & Ec & comma & fc & comma & b1 & comma & _
               esu & comma & fy1 & comma & Es & comma & rDNE

strArray(32) = "[Reduction Factors]"

'[strArray(33) is for reduction factors]
'[1. Phi(a) for axial compression]
'[2. Phi(b) for tension-controlled failure]
'[3. Phi(c) for compression-controlled failure]
'[4. Reserved. Do not edit.]

strArray(33) = "0.800000,0.900000,0.650000,0.100000"

strArray(34) = "[Design Criteria]"
'-----------------------------------------------------------------------------------------------
Dim minRR As String
Dim maxRR As String
Dim minClearSpace As String
Dim drRatio As String

minRR = "0.010000"          '[1. Minimum reinforcement ratio]
maxRR = "0.080000"          '[2. Maximum reinforcement ratio]
minClearSpace = "2.115000"  '[3. Minimum clear spacing between bars]
drRatio = "1.000000"        '[4. Design/Required ratio]
'-----------------------------------------------------------------------------------------------
strArray(35) = minRR & comma & maxRR & comma & minClearSpace & comma & drRatio

strArray(36) = "[External Points]"
strArray(37) = CStr(numPtsExtCol)                '[NOT USED]
strArray(38) = "[Internal Points]"
strArray(39) = CStr(numPtsIntSec)                '[NOT USED]
strArray(40) = "[Reinforcement Bars]"
strArray(41) = "0"
strArray(42) = "[Factored Loads]"
strArray(43) = CStr(numFactLoads)                '[NOT USED]



'[VBA Code reads to here, stop, reads in load combination forces, and continues below]

strArray(44) = "[Slenderness: Column]"                                      '|
strArray(45) = "0.000000,0.000000,0.000000,1,0,1.000000,1.000000"           '|
strArray(46) = "0.000000,0.000000,0.000000,1,0,1.000000,1.000000"           '|
strArray(47) = "[Slenderness: Column Above And Below]"                      '|
strArray(48) = "1,0.000000,0.000000,0.000000,4.000000,3605.000000"          '|
strArray(49) = "1,0.000000,0.000000,0.000000,4.000000,3605.000000"          '|
strArray(50) = "[Slenderness: Beams]"                                       '|NOT USED FOR
strArray(51) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|COLUMN DESIGN
strArray(52) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|Default values
strArray(53) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|have been placed
strArray(54) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|
strArray(55) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|
strArray(56) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|
strArray(57) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|
strArray(58) = "1,0.000000,0.000000,0.000000,0.000000,4.000000,3605.000000" '|

strArray(59) = "[EI]"
strArray(60) = "0.000000"
strArray(61) = "[SldOptFact]"
strArray(62) = "0"
strArray(63) = "[Phi_Delta]"
strArray(64) = "0.750000"
strArray(65) = "[Cracked I]"
strArray(66) = "0.350000,0.700000"
strArray(67) = "[Service Loads]"
strArray(68) = CStr(numServLoads)
strArray(69) = "[Load Combinations]"                                        '|
strArray(70) = CStr(numLC)                                                  '|
strArray(71) = "1.400000,0.000000,0.000000,0.000000,0.000000"               '|
strArray(72) = "1.200000,1.600000,0.000000,0.000000,0.500000"               '|
strArray(73) = "1.200000,1.000000,0.000000,0.000000,1.600000"               '|
strArray(74) = "1.200000,0.000000,0.800000,0.000000,1.600000"               '|
strArray(75) = "1.200000,1.000000,1.600000,0.000000,0.500000"               '|NOT USED FOR
strArray(76) = "0.900000,0.000000,1.600000,0.000000,0.000000"               '|COLUMN DESIGN
strArray(77) = "1.200000,0.000000,-0.800000,0.000000,1.600000"              '|Default values
strArray(78) = "1.200000,1.000000,-1.600000,0.000000,0.500000"              '|have been placed
strArray(79) = "0.900000,0.000000,-1.600000,0.000000,0.000000"              '|
strArray(80) = "1.200000,1.000000,0.000000,1.000000,0.200000"               '|
strArray(81) = "0.900000,0.000000,0.000000,1.000000,0.000000"               '|
strArray(82) = "1.200000,1.000000,0.000000,-1.000000,0.200000"              '|
strArray(83) = "0.900000,0.000000,0.000000,-1.000000,0.000000"              '|
strArray(84) = "[BarGroupType]"
strArray(85) = "1"        '[0-User defined;1-ASTM615;2-CSA G30.18;3-prEN 10080;4-ASTM615M]
strArray(86) = "[User Defined Bars]"

End Sub

