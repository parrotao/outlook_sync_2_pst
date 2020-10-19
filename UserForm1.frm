VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5985
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   12510
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 
Private Sub CommandButton1_Click()

 Set myNameSpace = Application.GetNamespace("MAPI")
DoEvents
 
 
 For i = 1 To myNameSpace.Folders.Count
 DoEvents
 
   UserForm1.ListBox1.AddItem (myNameSpace.Folders(i).Name)
 Next
 
End Sub

Private Sub CommandButton2_Click()
Dim a As String

UserForm1.TextBox1 = UserForm1.ListBox1.Value
a = UserForm1.TextBox1
UserForm1.ListBox2.Clear

 For i = 1 To myNameSpace.Folders(a).Folders.Count
    UserForm1.ListBox2.AddItem (myNameSpace.Folders(a).Folders(i).Name)
    
    
 Next
End Sub

Private Sub CommandButton3_Click()
Dim a As String

UserForm1.TextBox2 = UserForm1.ListBox1.Value
a = UserForm1.TextBox2
UserForm1.ListBox3.Clear

 For i = 1 To myNameSpace.Folders(a).Folders.Count
    UserForm1.ListBox3.AddItem (myNameSpace.Folders(a).Folders(i).Name)
    
    
 Next
End Sub

Private Sub CommandButton4_Click()

If UserForm1.TextBox1.Value = "" Or UserForm1.ListBox2.Value = "" Or UserForm1.TextBox2.Value = "" Or UserForm1.ListBox3.Value = "" Or Val(UserForm1.TextBox3.Value) <= Val(UserForm1.TextBox4.Value) Then

    MsgBox "input error"
Else
Call MoveItems(UserForm1.TextBox1.Value, UserForm1.ListBox2.Value, UserForm1.TextBox2.Value, UserForm1.ListBox3.Value, UserForm1.TextBox3.Value, UserForm1.TextBox4.Value)

MsgBox "done"
End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub CommandButton5_Click()

If UserForm1.TextBox1.Value = "" Or UserForm1.ListBox2.Value = "" Then

    MsgBox "input error"
Else
Call Adv_MoveItems(UserForm1.TextBox1.Value, UserForm1.ListBox2.Value, "Mailbox", UserForm1.ListBox2.Value)
MsgBox "done"
End If
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub UserForm_Click()

End Sub
