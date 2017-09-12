Attribute VB_Name = "Module1"

Function getFolder(psl)
   getFolder = ""
   If InStr(psl, "Google") > 0 Then getFolder = "Google"
   If InStr(psl, "Windows") > 0 Then getFolder = "Microsoft/Windows"
   If InStr(psl, "OneDrive") > 0 Then getFolder = "Microsoft/OneDrive"
   If InStr(psl, "Sway") > 0 Then getFolder = "Microsoft/Sway"
End Function


Sub toFolder()

   Set myapp = CreateObject("Outlook.Application")
   '��M�g���C
   Set i_Folder = myapp.Session.GetDefaultFolder(6)
   ' ��M�g���C�̓��e���ړ�
   Dim oDest As Outlook.MAPIFolder '�t�H���_�[

   UserForm1.Show vbModeless
   UserForm1.Label2 = i_Folder.Items.Count
   UserForm1.Label1 = 0

   '��M�g���C��S������
   cnt = 1
   For idx = i_Folder.Items.Count To 1 Step -1
      On Error GoTo CONTINUE
      
      psl = i_Folder.Items(idx).SentOnBehalfOfName
      sbj = i_Folder.Items(idx).Subject
      
      fld = getFolder(psl)
      If fld <> "" Then
         If InStr(fld, "/") > 0 Then
            f1 = Split(fld, "/")(0)
            f2 = Split(fld, "/")(1)
            Set oDest = Application.Session.Folders("�l�p Outlook �f�[�^ �t�@�C��").Folders(f1).Folders(f2)
            i_Folder.Items(idx).Move oDest
         Else
            Set oDest = Application.Session.Folders("�l�p Outlook �f�[�^ �t�@�C��").Folders(fld)
            i_Folder.Items(idx).Move oDest
         End If
      End If

CONTINUE:

      
      UserForm1.Label1 = cnt
      DoEvents
      cnt = cnt + 1
      
   Next idx

   Unload UserForm1

End Sub

