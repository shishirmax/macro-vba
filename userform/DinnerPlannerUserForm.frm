VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DinnerPlannerUserForm 
   Caption         =   "Dinner Planner"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   OleObjectBlob   =   "DinnerPlannerUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DinnerPlannerUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

Unload Me

End Sub

Private Sub ClearButton_Click()

Call UserForm_Initialize

End Sub

Private Sub DinnerComboBox_Change()

End Sub

Private Sub MoneySpinButton_Change()

MoneyTextBox.Text = MoneySpinButton.Value

End Sub

Private Sub OKButton_Click()

Dim emptyRow As Long

'Make Sheet1 active
Sheet1.Activate

'Determine emptyRow
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information
Cells(emptyRow, 1).Value = NameTextBox.Value
Cells(emptyRow, 2).Value = PhoneTextBox.Value
Cells(emptyRow, 3).Value = CityListBox.Value
Cells(emptyRow, 4).Value = DinnerComboBox.Value

If DateCheckBox1.Value = True Then Cells(emptyRow, 5).Value = DateCheckBox1.Caption

If DateCheckBox2.Value = True Then Cells(emptyRow, 5).Value = Cells(emptyRow, 5).Value & " " & DateCheckBox2.Caption

If DateCheckBox3.Value = True Then Cells(emptyRow, 5).Value = Cells(emptyRow, 5).Value & " " & DateCheckBox3.Caption

If CarOptionButton1.Value = True Then
    Cells(emptyRow, 6).Value = "Yes"
Else
    Cells(emptyRow, 6).Value = "No"
End If

Cells(emptyRow, 7).Value = MoneyTextBox.Value

End Sub

Private Sub UserForm_Initialize()

'Empty NameTextBox
NameTextBox.Value = ""

'Empty PhoneTextBox
PhoneTextBox.Value = ""

'Empty CityListBox
CityListBox.Clear

'Fill CityListBox
With CityListBox
    .AddItem "San Francisco"
    .AddItem "Oakland"
    .AddItem "Richmond"
    .AddItem "New Delhi"
    .AddItem "Mumbai"
    .AddItem "Bangalore"
    .AddItem "Pune"
End With

'Empty DinnerComboBox
DinnerComboBox.Clear

'Fill DinnerComboBox
With DinnerComboBox
    .AddItem "Italian"
    .AddItem "Chinese"
    .AddItem "Frites and Meat"
End With

'Uncheck DataCheckBoxes
DateCheckBox1.Value = False
DateCheckBox2.Value = False
DateCheckBox3.Value = False

'Set no car as default
CarOptionButton2.Value = True

'Empty MoneyTextBox
MoneyTextBox.Value = ""

'Set Focus on NameTextBox
NameTextBox.SetFocus

End Sub
