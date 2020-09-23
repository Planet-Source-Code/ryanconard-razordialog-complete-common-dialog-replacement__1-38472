VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl RazorDialog 
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "RazorDialog.ctx":0000
   ScaleHeight     =   495
   ScaleMode       =   0  'User
   ScaleWidth      =   500
   Begin MSComDlg.CommonDialog CD1 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "RazorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'RazorDialog Control - Complete Common Dialog Replacement
'© Copyright 2002 Ryan Conard
'RazorDialog is a Closed Source user control. You may use the OCX
'freely without altering the code.

Event ShowAbout(AhowAbout)
Event SetFont(SetFont)
Event SetBackColor(SetBackColor)
Event SetTextColor(SetTextColor)
Event OpenFileCustomFilter(OpenFileCustomFilter)
Event OpenFileDefaultFilter(OpenFileDefaultFilter)
Event SaveFileCustomFilter(SaveFileCustomFilter)
Event SaveFileDefaultFilter(SaveFileDefaultFilter)

Function ShowAbout()
MsgBox "RazorDialog - (C) Copyright Ryan Conard. - A complete Common Dialog replacement.", vbInformation, "RazorDialog"
End Function

Function SetFont(RichTextBox As Variant)
Dim TextColor As Long
Dim Bold As Boolean
Dim Italic As Boolean
Dim Underline As Boolean
Dim StrikeThru As Boolean
Dim Font As String
Dim Size As Integer
CD1.DialogTitle = "RazorDialog - Font"
CD1.Flags = cdlCFEffects Or cdlCFBoth
CD1.ShowFont
TextColor = CD1.Color
Bold = CD1.FontBold
Italic = CD1.FontItalic
Underline = CD1.FontUnderline
StrikeThru = CD1.FontStrikethru
Font = CD1.FontName
Size = CD1.FontSize
RichTextBox.SelFontName = Font
RichTextBox.SelFontSize = Size
RichTextBox.SelColor = TextColor
If Bold = True Then
RichTextBox.SelBold = Bold
If Italic = True Then
RichTextBox.SelItalic = Italic
If Underline = True Then
RichTextBox.SelUnderline = Underline
If StrikeThru = True Then
RichTextBox.SelStrikeThru = StrikeThru
End If
End If
End If
End If
End Function

Function SetBackColor(RichTextBox As Variant)
Dim SelectedColor As Long
CD1.DialogTitle = "RazorDialog - Color"
CD1.Flags = cdlCCRGBInit
CD1.ShowColor
SelectedColor = CD1.Color
RichTextBox.BackColor = SelectedColor
End Function

Function SetTextColor(RichTextBox As Variant)
Dim SelectedColor As Long
CD1.DialogTitle = "RazorDialog - Color"
CD1.Flags = cdlCCRGBInit
CD1.ShowColor
SelectedColor = CD1.Color
RichTextBox.SelColor = SelectedColor
End Function

Function OpenFileCustomFilter(RichTextBox As Variant, Filter As Variant)
Dim FileName As String
CD1.DialogTitle = "RazorDialog - Open File"
'Example of a basic filter...
' "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
'With the ""
CD1.Filter = Filter
CD1.ShowOpen
FileName = CD1.FileName
RichTextBox.FileName = FileName
End Function

Function SaveFileCustomFilter(RichTextBox As Variant, Filter As Variant)
Dim FileName As String
CD1.DialogTitle = "RazorDialog - Save File"
'Example of a basic filter...
' "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
'With the ""
CD1.Filter = Filter
CD1.ShowSave
FileName = CD1.FileName
RichTextBox.SaveFile FileName
End Function

Function OpenFileDefaultFilter(RichTextBox As Variant)
Dim FileName As String
CD1.DialogTitle = "RazorDialog - Open File"
CD1.Filter = "Text Files (*.TXT)|*.TXT|Word Documents (*.DOC)|*.DOC|All Files (*.*)|*.*"
CD1.ShowOpen
FileName = CD1.FileName
RichTextBox.FileName = FileName
End Function

Function SaveFileDefaultFilter(RichTextBox As Variant)
Dim FileName As String
CD1.DialogTitle = "RazorDialog - Save File"
CD1.Filter = "Text Files (*.TXT)|*.TXT|Word Documents (*.DOC)|*.DOC|All Files (*.*)|*.*"
CD1.ShowSave
FileName = CD1.FileName
RichTextBox.SaveFile FileName
End Function

Private Sub UserControl_Initialize()
UserControl.Height = 500
UserControl.Width = 500
End Sub
