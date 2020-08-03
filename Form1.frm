VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghapus File ke Recycle Bin"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MyBool As Boolean
  'Ganti nama file di bawah dengan nama file yang ingin
  'Anda hapus.
  DelToRecycBin ("c:\My Documents\MyFile.Zip")
End Sub

Public Function DelToRecycBin(FileName As String)
Dim FileOperation As SHFILEOPSTRUCT
Dim lReturn As Long
On Error GoTo DelToRecycBin_Err
  With FileOperation
     .wFunc = F0_DELETE
     .pFrom = FileName
     .fFlags = F0F_ALLOWUNDO + F0F_CREATEPROGRESSDLG
  End With
  lReturn = SHFileOperation(FileOperation)
  Exit Function
DelToRecycBin_Err:
  MsgBox Err.Number & Err.Description
End Function

