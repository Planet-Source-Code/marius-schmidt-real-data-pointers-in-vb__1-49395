VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This sample app shows that VB stores its data in the same order as it gets defined

'Well, where could we use this code?
'Did you ever noticed that in D3D there's a lock function?
'it will just return a pointer to some data, but if you want
'to modify this data you'll most likely copy an array of data
'to this location, with this code you could just create a pointer
'to this data so you wont have to create a temporary array just
'for copying the data to the memory location!

'We want to point at these 4 variables
Dim SomeData As Long
Dim SomeOtherData As Long
Dim SomeMoreData As Long
Dim YesItsMoreData As Long

Dim pData() As Long 'Our pointer

Private Sub Form_Load()
ReDim pData(0) As Long 'initialize the pointer

Dim sah As SAFEARRAYHEADER 'Holds array settings
Dim pDataPos As Long


SomeData = 1024
SomeOtherData = 25
SomeMoreData = 99
YesItsMoreData = 1000000


pDataPos = VarPtr(SomeData) 'Get the address of the first variable

'Let pData point to the first variable we defined
RedimArray 4, 4, sah, pDataPos, VarPtrArray(pData())

MsgBox pData(0)
MsgBox pData(1)
MsgBox pData(2)
MsgBox pData(3)

End Sub
