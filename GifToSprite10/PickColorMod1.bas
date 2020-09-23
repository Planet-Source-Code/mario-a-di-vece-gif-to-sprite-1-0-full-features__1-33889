Attribute VB_Name = "PickColorMod1"
Public GetColorReturn&

'Who needs 16.777.216 colors anyway ???
'I think 32.768 colors is more than enough for most purposes...

Public Function GetColor(Optional OldCol As Variant) As Long
'On Error GoTo GetColor_error
GetColor = &H7FFFFFFF
With ColFrm
.OldLabel.BackColor = &HC0C0C0
If IsMissing(OldCol) Then GoTo Further
.OldLabel.BackColor = OldCol
Further:
ColFrm.Show 1
If GetColorReturn = 0 Then Exit Function
GetColor = .PickLabel.BackColor
Exit Function
End With
GetColor_error:
MsgBox Err.Number & vbCr & Err.Description, vbCritical, "PickColor"
End Function
