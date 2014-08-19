Attribute VB_Name = "BB_OpenSAP"
Option Explicit

Sub Open_SAP_Model(strPath As String)


'create Sap2000 object
Set SapObject = CreateObject("SAP2000.SapObject")

'Start SAP2000 application
SapObject.ApplicationStart

'create SapModel object
Set SapModel = SapObject.SapModel

'initialize model
Ret = SapModel.InitializeNewModel

'Open SAP Model
Ret = SapObject.SapModel.File.OpenFile(strPath)

'Set units to user input units
Ret = SapModel.SetPresentUnits(Worksheets("Input").Cells(2, 5))

End Sub

