VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} POCDialog 
   Caption         =   "Project POC Dialog"
   ClientHeight    =   4620
   ClientLeft      =   132
   ClientTop       =   492
   ClientWidth     =   7884
   OleObjectBlob   =   "POCDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "POCDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' POCDialog is a dialog used to select update the three main project info related
' points of contact relating to a RFI bid log: the PM, TL and Tech Services POCs.

Private Sub okButton_Click()
    If Me.pmTextBox.Value <> "" Then
        ActiveSheet.Range("B3").Value = Me.pmTextBox.Value
    End If
    If Me.tlTextBox.Value <> "" Then
        ActiveSheet.Range("B4").Value = Me.tlTextBox.Value
    End If
    If Me.tsTextBox.Value <> "" Then
        ActiveSheet.Range("B5").Value = Me.tsTextBox.Value
    End If
    If Me.tsTextBox.Value <> "" Then
        ActiveSheet.Range("B6").Value = Me.corTextBox.Value
    End If
    If Me.tsTextBox.Value <> "" Then
        ActiveSheet.Range("B7").Value = Me.csTextBox.Value
    End If
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.versionLabel.Caption = "Project Contacts" & vbCrLf & _
        BidGulp.module_name & " v." & BidGulp.module_version
    Me.pmTextBox.Value = ActiveSheet.Range("B3").Value
    Me.tlTextBox.Value = ActiveSheet.Range("B4").Value
    Me.tsTextBox.Value = ActiveSheet.Range("B5").Value
    Me.corTextBox.Value = ActiveSheet.Range("B6").Value
    Me.csTextBox.Value = ActiveSheet.Range("B7").Value
    Me.Show
End Sub
