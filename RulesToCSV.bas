Attribute VB_Name = "Module3"
'TBA:
'exceptions
'full conditions and actions lists
'compatibility with multiple actions and conditions
'formatting for multiple action paramaters

Sub ListRules()
    Dim oFileSys As Object
    Dim oStore As Outlook.Store
    Dim colRules As Object
    Dim oRule As Outlook.Rule
    Dim sOutput As String

    'On Error Resume Next
    Set oFileSys = CreateObject("Scripting.FileSystemObject")

    If oFileSys.FileExists("C:\Outlook Files\OLrules.csv") Then
        oFileSys.DeleteFile ("C:\Outlook Files\OLrules.csv")
    End If
    Open "D:\Users\tsuma\Documents\Outlook Files\OLrules.csv" For Output As #1
    Set oStore = Application.Session.DefaultStore
    Set colRules = oStore.GetRules
    
    sOutput = """Num"",""Name"",""Action"",""Action Parameters"",""Condition"",""Condition Parameters""" & vbCrLf
    
    For Each oRule In colRules
    sOutput = sOutput & """" & oRule.ExecutionOrder & """,""" & oRule.Name & """"
    
    If (oRule.Actions.MoveToFolder.Enabled = True) Then
        Call AddToOutput(sOutput, "MoveTo", oRule.Actions.MoveToFolder.Folder.FolderPath)
    ElseIf (oRule.Actions.Stop.Enabled = True) Then
        Call AddToOutput(sOutput, "Stop", oRule.Actions.Stop.Enabled)
    End If
    
    If (oRule.Conditions.Subject.Enabled = True) Then
        Call AddToOutput(sOutput, "InSubject", oRule.Conditions.Subject.Text)
    ElseIf (oRule.Conditions.SenderAddress.Enabled = True) Then
        Call AddToOutput(sOutput, "InAddress", oRule.Conditions.SenderAddress.Address)
    ElseIf (oRule.Conditions.From.Enabled = True) Then
        Call AddToOutput(sOutput, "IsAddress", oRule.Conditions.From.Recipients)
    End If
            
    sOutput = sOutput & vbCrLf
            
    Next
    Print #1, sOutput
    Close #1
End Sub
Private Function AddToOutput(sOutput As String, ConditionName As String, ConditionInput)
    Dim iterator
    
    sOutput = sOutput & ",""" & ConditionName & """"
    
    If (IsArray(ConditionInput)) Then
        For Each iterator In ConditionInput
            sOutput = sOutput & ",""" & iterator & """"
        Next
    ElseIf (IsObject(ConditionInput)) Then
        For Each iterator In ConditionInput
            If (Not IsMissing(iterator.Address)) Then
                sOutput = sOutput & ",""" & iterator.Address & """"
            End If
        Next
    Else
        sOutput = sOutput & ",""" & ConditionInput & """"
    End If
    
End Function
