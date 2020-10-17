Attribute VB_Name = "Module1"
Sub MoveItems(ft1, ft2, tt1, tt2, t1, t2)
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 Dim i
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 
 

 
 Set myInbox = myNameSpace.Folders(ft1).Folders(ft2)
 Set myItems = myInbox.Items
 
 Set myDestFolder = myNameSpace.Folders(tt1).Folders(tt2)
 


For i = myItems.Count To 1 Step -1
 Set myItem = myItems(i)
On Error Resume Next
 If Not (myItem.ReceivedTime > Now() - t1 And myItem.ReceivedTime < Now() - t2) Then
    UserForm1.Label1 = myItems.Count
    DoEvents
 Else
    myItem.Move myDestFolder
    DoEvents
    UserForm1.Label1 = myItems.Count

    UserForm1.Label2 = myDestFolder.Items.Count
 End If
On Error GoTo 0

Next


End Sub

Sub Adv_MoveItems(ft1, ft2, tt1, tt2)

 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 Dim i
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 
 

 
 Set myInbox = myNameSpace.Folders(ft1).Folders(ft2)
 Set myItems = myInbox.Items
 
 


For i = myItems.Count To 1 Step -1
 Set myItem = myItems(i)
On Error Resume Next

 If Trim(Year(Now)) = Trim(Year(myItem.ReceivedTime)) Then
    UserForm1.Label1 = myItems.Count
    DoEvents
 Else
    Set myDestFolder = myNameSpace.Folders(tt1 & "_" & Trim(Year(myItem.ReceivedTime))).Folders(tt2)
    myItem.Move myDestFolder
    DoEvents
    UserForm1.Label1 = myItems.Count

    UserForm1.Label2 = myDestFolder.Items.Count
 End If
On Error GoTo 0

Next


End Sub

