Private Sub cmdGeStPart_Click()

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc
Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

'######## List of command execution steps ########

'1. Convert toolbox file to *.step
    '1-1 Check if the a 3dmodel is open or no
    '1-2. Copy path and toolbox file name in the File_Name Variabel
    '1-3. Save as Toolbox Part to STEP data in Windows temp folder (C:\temp\)
    '1-4. Close Toolbox Part
    '1-5. Open STEP

'2. Add custom property and save as *-New.sldprt
    '2-1 Add custom property
    '2-2 Save as *-New.sldprt


'########## Start ###########

'1. Convert toolbox file to *.step

    '1-1 Check if the a 3dmodel is open or no
If swModel Is Nothing Then 'Check to see if a document is loaded
    swApp.SendMsgToUser2 "Please open a Standard Part Document and try again.", swMbInformation, swMbOk
    UserForm_SPG.Hide
    Exit Sub
End If
    
    '1-2. Copy path and toolbox file name in the File_Name Variabel

sModelName = swModel.GetPathName
'MsgBox (sModelName)

    '1-3. Save as Toolbox Part to STEP data in Windows temp folder (C:\temp\)
'longstatus = Part.SaveAs3(sModelName + ".step", 0, 2)
longstatus = Part.SaveAs3("C:\temp\SW_Standard_Temp_Part.step", 0, 2)

    '1-4. Close Toolbox Part
Set Part = Nothing
swApp.CloseDoc sModelName

    '1-5. Open STEP and set Display style to Shaded with edges
'boolstatus = swApp.LoadFile2(sModelName + ".step", "r") 'Open STEP
boolstatus = swApp.LoadFile2("C:\temp\SW_Standard_Temp_Part.step", "r") 'Open STEP

Set Part = swApp.ActiveDoc
Dim activeModelView As Object
Set activeModelView = Part.ActiveView
activeModelView.DisplayMode = swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges

'2. Add custom property and save as *.sldprt

    '2-1 Add custom property
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActivateDoc("")

If swModel Is Nothing Then 'Check to see if a document is loaded
    swApp.SendMsgToUser2 "Please open a Part Document.", swMbInformation, swMbOk
    UserForm_SPG.Hide
    Exit Sub
End If

        'Add custom property to Existing Part
    Dim retval As Long
    Dim custPropMan As CustomPropertyManager
    Dim Benennung  As String
    
    If optButGewSti.Value = True Then
    Benennung = "Gewindestift"
    End If
    
    If optButLinSch.Value = True Then
    Benennung = "Linsenschraube"
    End If
    
    If optButSchei.Value = True Then
    Benennung = "Scheibe"
    End If
    
    If optButSecKanMut.Value = True Then
    Benennung = "Sechskantmutter"
    End If
    
    If optButSecKanSch.Value = True Then
    Benennung = "Sechskantschraube"
    End If
    
    If optButSenSch.Value = True Then
    Benennung = "Senkschraube"
    End If
    
    If optButSpeSch.Value = True Then
    Benennung = "Spezialschraube"
    End If
    
    If optButOther.Value = True Then
    Benennung = txbxOther.Text
    End If


    Set custPropMan = swModel.Extension.CustomPropertyManager("")
    retval = custPropMan.Add3("Zeichnung.- Nr.", swCustomInfoText, "$PRP:""SW-File Name""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Benennung", swCustomInfoText, Benennung, swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Werkstoff", swCustomInfoText, """SW-Material@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Weight", swCustomInfoText, """SW-Mass@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)

    '2-2 Save as *.sldprt
'longstatus = Part.SaveAs3(sModelName + "-New" + ".sldprt", 0, 2)
longstatus = Part.SaveAs3(sModelName, 0, 2)
MsgBox ("New Standard part had generated in: " + sModelName)


    '2-3 Close SW_Standard_Temp_Part.step
Dim sTempName As String
sTempName = swModel.GetPathName
Set Part = Nothing
swApp.CloseDoc sTempName

    '2-4 Open new Standard Part

'boolstatus = swApp.LoadFile2(sModelName, "r")
'boolstatus = swApp.OpenDoc6(sModelName, False, True, True, "")
'MsgBox (sModelName + "   =>  info for open")

End Sub
