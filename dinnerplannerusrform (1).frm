VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8592.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8352.001
   OleObjectBlob   =   "dinnerplannerusrform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uprow As Long

Private Sub CommandButton1_Click()
'Add button
Dim r As Long
r = Sheets("dinner planner").Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(r, 1).Value = TextBox1.Value
Cells(r, 2).Value = TextBox2.Value
Cells(r, 3).Value = ListBox1.Value
Cells(r, 4).Value = ComboBox1.Value
If CheckBox1.Value = True Then
Cells(r, 5).Value = "June 13"
ElseIf CheckBox2.Value = True Then
Cells(r, 5).Value = "June 27"
ElseIf CheckBox3.Value = True Then
Cells(r, 5).Value = "June 20"
End If
If OptionButton1.Value = True Then
Cells(r, 6) = "Yes"
Else
Cells(r, 6) = "no"
End If
Cells(r, 7) = TextBox3.Value
End Sub

Private Sub CommandButton2_Click()
' update button
Cells(uprow, 1) = TextBox1.Value
Cells(uprow, 2) = TextBox2.Value
Cells(uprow, 3) = ListBox1.Value
Cells(uprow, 4) = ComboBox1.Value
If CheckBox1.Value = True Then
Cells(uprow, 5) = "June 13"
ElseIf CheckBox2.Value = True Then
Cells(uprow, 5) = "June 27"
ElseIf CheckBox3.Value = True Then
Cells(uprow, 5) = "June 20"
End If
If OptionButton1.Value = True Then
Cells(uprow, 6) = "Yes"
ElseIf OptionButton2.Value = True Then
Cells(uprow, 6) = "No"
End If
Cells(uprow, 7) = TextBox3.Value
End Sub

Private Sub CommandButton3_Click()
' Search button
Range("a1").Select
Dim rrow As Long
rrow = Cells.Find(What:=TextBox1.Text, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Row
uprow = rrow ' for the update button
MsgBox ("the row number in which name is present " & rrow)
TextBox1.Value = Cells(rrow, 1)
TextBox2.Value = Cells(rrow, 2)
ListBox1.Value = Cells(rrow, 3)
ComboBox1.Value = Cells(rrow, 4)
If Format(Cells(rrow, 5).Value, "mmmm dd") = "June 13" Then
CheckBox1.Value = True
ElseIf Format(Cells(rrow, 5).Value, "mmmm dd") = "June 20" Then
CheckBox3.Value = True
ElseIf Format(Cells(rrow, 5).Value, "mmmm dd") = "June 27" Then
CheckBox2.Value = True
End If
If Cells(rrow, 6).Value = "Yes" Then
OptionButton1.Value = True
Else
OptionButton2.Value = True
End If
TextBox3.Value = Cells(rrow, 7)
End Sub

Private Sub CommandButton4_Click()
'Clear button
TextBox1.Value = ""
TextBox2.Value = ""
ListBox1.Value = ""
ComboBox1.Value = ""
If CheckBox1.Value = True Then
CheckBox1.Value = False
ElseIf CheckBox2.Value = True Then
CheckBox2.Value = False
ElseIf CheckBox3.Value = True Then
CheckBox3.Value = False
End If
If OptionButton1.Value = True Then
OptionButton1.Value = False
Else
OptionButton2.Value = False
End If
TextBox3.Value = ""
End Sub

Private Sub CommandButton5_Click()
'Exit button
Unload UserForm1
End Sub


Private Sub CommandButton6_Click()
' Delete button
Dim ans
ans = MsgBox("Want to delete the records", vbYesNo)
If ans = vbYes Then
Cells(uprow, 1).EntireRow.Delete
End If
End Sub

Private Sub SpinButton1_Change()
TextBox3.Value = SpinButton1.Value
End Sub

Private Sub UserForm_initialize()
With ListBox1
.AddItem ("San Fransiciso")
.AddItem ("öakland")
.AddItem ("Richmond")
End With
With ComboBox1
.AddItem ("Vegetarian")
.AddItem ("Vegan")
.AddItem ("Seafood")
.AddItem ("No preference")
End With
Range("a1").Value = "Name"
Range("b1").Value = "Phone number"
Range("c1").Value = "City preference"
Range("d1").Value = "Dinner preference"
Range("e1").Value = "Date"
Range("f1").Value = "Do you have car"
Range("g1").Value = "Maximum to spend"
End Sub
