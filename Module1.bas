Attribute VB_Name = "Module1"
'Jika pilihan 'delete to recycle bin' Windows di-'nonaktif-kan, file akan langsung dihapus secara 'permanen (?)... hati-hati!

Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Boolean
hNameMappings As Long
lpszProgressTitle As String
End Type
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const F0_DELETE = &H3
Public Const F0F_ALLOWUNDO = &H40
Public Const F0F_CREATEPROGRESSDLG As Long = &H0


