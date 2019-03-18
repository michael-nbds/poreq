Attribute VB_Name = "Module1"
Attribute VB_Description = "globals"
Option Explicit
'---------------------
DefInt A-Z

Global DB As Database
Global DefaultWorkspace As Workspace
Sub SelectText(ctrIn As Control)
    ctrIn.SelStart = 0
    ctrIn.SelLength = Len(ctrIn.Text)
End Sub
Sub main()
    Set DefaultWorkspace = Workspaces(0)
    Set DB = DefaultWorkspace.OpenDatabase(App.Path & "\wmars.mdb")

    MainFrm.Show 1

End Sub
