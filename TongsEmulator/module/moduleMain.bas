Attribute VB_Name = "moduleMain"
Public MainServer As Server

Public ClientSession As Collection

Public Sub Main()
    formMain.Show
    
    Set MainServer = New Server
    Set ClientSession = New Collection
End Sub
