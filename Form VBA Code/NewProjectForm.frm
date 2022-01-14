VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProjectForm 
   Caption         =   "UserForm2"
   ClientHeight    =   4680
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6920
   OleObjectBlob   =   "NewProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub submitBtn_Click()
    'Set workbook and sheet
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Tests&Projects")

    'start on second row (headers are first row)
    intRow = 3

    'Test value of Item textbox
    If (itemTxt.Value <> "") Then
    
        'Test value of class textbox
        If (classTxt.Value <> "") Then
        
            'Test value of Duration Textbox
            If (durationTxt.Value <> "") Then
    
                'Go through rows, if they contain data, increment
                Do While (ws.Cells(intRow, "B") <> "")
                
                    'Increment row counter
                    intRow = intRow + 1
                
                Loop
                                
                'Write date into cell
                ws.Cells(intRow, "B") = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
        
                'Format cell so Excel recognizes a date
                ws.Cells(intRow, "B").NumberFormat = "yyyy-mm-dd;@"
                
                'Write item into cell
                ws.Cells(intRow, "C") = itemTxt.Value
                 
                'Write Duration into cell
                ws.Cells(intRow, "E") = durationTxt.Value
                
                'Write Class into cell
                ws.Cells(intRow, "D") = classTxt.Value
    
            
            Else
                'Give error for no category
                MsgBox ("Please select a Duration")
            End If
        
        Else
            'Give error message for no class
            MsgBox ("Please enter a class")
        End If
        
    Else
        'Give error message for no item
        MsgBox ("Please enter an item")
    End If
    
End Sub
