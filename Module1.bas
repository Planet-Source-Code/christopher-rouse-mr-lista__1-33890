Attribute VB_Name = "Module1"
Type cus
    genre As String
End Type
Public genreList(126) As cus
Global sd As String
Global artg As String
Global mcfca(50) As Integer
Global titg As String
Global nocdg As Integer
Global picg(50, 15) As String
Global pnum As Integer
Global front(50) As Integer
Global back(50) As Integer
Global fault(50)
Global comments(50) As String
Global sf As Integer
Global sf2 As Integer
Global cwo As Integer
Global ind As Integer
Global iod As String
Global gnod As Integer
Global titles(50) As String
Global artists(50) As String
Global nocds(50) As Integer

Option Explicit
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
        "SHGetPathFromIDListA" (ByVal pidl As Long, _
        ByVal pszPath As String) As Long

Const BIF_RETURNONLYFSDIRS = &H1

Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName  As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Type SHITEMID
   cb As Long
   abID As Byte
End Type

Type ITEMIDLIST
   mkid As SHITEMID
End Type


