Attribute VB_Name = "Mod_Movefrm"
Option Compare Database

Option Explicit

Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As LongPtr, _
ByVal wMsg As Long, _
ByVal wParam As LongPtr, IParam As Any) As Long

Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As LongPtr
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HT_CAPTION = &H2

