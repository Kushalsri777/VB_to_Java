Attribute VB_Name = "modColorDepthAPI"
'******************************************************************************
' Module     : modColorDepthAPI
' Description: Used by Desktop Technology's standard Splash screen
' Procedures : N/A
' Modified   :
'
' --------------------------------------------------
Option Explicit
Option Compare Binary

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

