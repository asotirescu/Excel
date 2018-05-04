Option Explicit
Sub addCBX()
Dim myCBX As CheckBox
Dim myCell As Range
    With ActiveSheet
    'The below is nice for testing, removes existing checkboxes
       .CheckBoxes.Delete
       'Alter the values in the parentheses below to specify where checkboxes should go
       For Each myCell In ActiveSheet.Range("W2:W50,X2:X50,Y2:Y50,Z2:Z50,AA2:AA50,AC2:AC50,AD2:AD50").Cells
           With myCell
           'Automatically adds checkboxes to sheet with proper sizing
               Set myCBX = .Parent.CheckBoxes.Add _
                               (Top:=.Top, Width:=.Width, _
                               Left:=.Left, Height:=.Height)
               With myCBX
                   .LinkedCell = myCell.Address(external:=True)
                   'Below only to be changed if caption is needed for checkbox
                   .Caption = ""
                   '.Name = "CBX_" & myCell.Address(0, 0) [IGNORE for now]
               End With
               'Formats True/False text inside cell to be invisible but still available for sorting purposes
               .NumberFormat = ";;;"
           End With
       Next myCell
   End With
End Sub
