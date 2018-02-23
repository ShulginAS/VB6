VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "TextEditZel"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6090
   Icon            =   "texteditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   1191
      ButtonWidth     =   1085
      ButtonHeight    =   1032
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Del"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2520
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Открыть"
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   34
         ImageHeight     =   33
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "texteditor.frx":16B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "texteditor.frx":17488
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "texteditor.frx":17D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "texteditor.frx":18674
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11033
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"texteditor.frx":18F6A
   End
   Begin VB.Menu file 
      Caption         =   "Файл"
      Begin VB.Menu make 
         Caption         =   "Создать"
      End
      Begin VB.Menu open 
         Caption         =   "Открыть"
      End
      Begin VB.Menu safe 
         Caption         =   "Сохранить "
      End
      Begin VB.Menu safeas 
         Caption         =   "Сохранить как"
      End
      Begin VB.Menu exit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu red 
      Caption         =   "Правка"
      Begin VB.Menu copy 
         Caption         =   "Копировать"
      End
      Begin VB.Menu out 
         Caption         =   "Вырезать"
      End
      Begin VB.Menu ins 
         Caption         =   "Вставить"
      End
      Begin VB.Menu del 
         Caption         =   "Удалить"
      End
   End
   Begin VB.Menu Format 
      Caption         =   "Формат"
      Begin VB.Menu Font 
         Caption         =   "Шрифт"
      End
      Begin VB.Menu Color 
         Caption         =   "Цвет"
         Begin VB.Menu FontColor 
            Caption         =   "Шрифт"
         End
         Begin VB.Menu Background 
            Caption         =   "Фон"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub copy_Click()
Clipboard.Clear 'очистка буфера обмена
Clipboard.SetText Form1.RichTextBox1.SelText 'выделенный текст в буфер обмена
End Sub

Private Sub del_Click()
Form1.RichTextBox1.SelText = "" 'удаление выделенного фрагмента
End Sub





Private Sub ins_Click()
Form1.RichTextBox1.Text = Clipboard.GetText() 'вставка фрагмента текста из буфера обмена
End Sub

Private Sub open_Click()
CommonDialog1.ShowOpen
RichTextBox1.LoadFile CommonDialog1.FileName
End Sub

Private Sub out_Click()
Clipboard.Clear 'очистка буфера обмена
Clipboard.SetText Form1.RichTextBox1.SelText 'выделенный текст в буфер обмена
Form1.RichTextBox1.SelText = "" 'удаление выделенного текста
End Sub

Private Sub safe_Click()

 If CommonDialog1.FileName = "" Then
 safe_Click
 Else
 RichTextBox1.SaveFile CommonDialog1.FileName
 End If
 End Sub

Private Sub safeas_Click()
CommonDialog1.ShowSave
RichTextBox1.SaveFile CommonDialog1.FileName
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case Is = "Copy"
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText
Case Is = "Cut"
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText
Form1.RichTextBox1.SelText = ""
Case Is = "Paste"
Form1.RichTextBox1.SelText = Clipboard.GetText()
Case Is = "Del"
Form1.RichTextBox1.SelText = ""
End Select
End Sub
