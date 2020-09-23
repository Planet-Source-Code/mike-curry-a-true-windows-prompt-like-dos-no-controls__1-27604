VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "MCBBX"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Screen 
      Height          =   780
      ItemData        =   "bix.frx":0000
      Left            =   7200
      List            =   "bix.frx":0002
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialogue 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Courier New"
      FontSize        =   10
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Cursor 
      Interval        =   210
      Left            =   120
      Top             =   120
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Colours 
         Caption         =   "&Colours"
      End
      Begin VB.Menu Font 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Prompt, x, y, CursorFlash, i, promptline, pline, xfactor, yfactor
Public screenwidth, screenheight, mxc, mxl, lastitem

Private Sub Colours_Click()
    Cursor.Enabled = False
    CommonDialogue.ShowColor
    Me.ForeColor = CommonDialogue.Color
    Cursor.Enabled = True
End Sub

Private Sub Font_Click()
    On Error Resume Next
    Cursor.Enabled = False
    CommonDialogue.ShowFont
    Me.Font = CommonDialogue.FontName
    Me.FontSize = CommonDialogue.FontSize
    Me.FontBold = CommonDialogue.FontBold
    Me.FontItalic = CommonDialogue.FontItalic
    Cursor.Enabled = True
    xfactor = Me.TextWidth("X")
    yfactor = Me.TextHeight("X")

End Sub

Private Static Sub Cursor_Timer()
                
    CursorFlash = CursorFlash + 1
    
    If CursorFlash = 2 Then CursorFlash = 0
            
    If CursorFlash = 0 Then
        Me.Line (x, y)-(x + xfactor, y + yfactor), Me.BackColor, BF
        update_xy
    Else
        Me.Line (x, y)-(x + xfactor, y + yfactor), 0, BF
        update_xy
    End If
    
End Sub

Private Static Sub Form_KeyPress(KeyAscii As Integer)
    Cursor.Enabled = False
    
    Dim compare, temp_prompt
    
    If KeyAscii = 13 Then
        
        remove_cursor
        
        x = 0
        y = y + yfactor
        
        If promptline = 0 Then
            temp_prompt = Prompt
        Else
            temp_prompt = ""
        End If
        
        Screen.AddItem temp_prompt & pline
        
        If Int(y / yfactor) >= Int(mxl / yfactor) Then
            
            x = 0: y = 0
            Me.Cls
            
            Screen.RemoveItem 0
            update_xy
        
            For i = 0 To Int(mxl / yfactor)
                    x = 0: update_xy
                    Me.Print Screen.List(i);
                    y = y + yfactor: x = 0
                    update_xy
            Next
            y = mxl - yfactor
        End If
        
        promptline = 0
        pline = ""
        
        
        update_xy
        Me.Print Prompt;
        
        x = xfactor * 2
        update_xy
        
        CursorFlash = 0
        Cursor.Enabled = True
        Exit Sub
    End If
    
    If KeyAscii = 8 Then
        
        compare = xfactor * 2
        If promptline Then compare = 0
        
        If x > compare Then
            x = x - xfactor
            If Len(pline) > 1 Then
                pline = Mid(pline, 1, Len(pline) - 1)
            Else
                pline = ""
            End If
        Else
            If promptline Then

                remove_cursor
                
                y = y - yfactor
                x = mxc - xfactor * 2
                                                                
                Dim Tempi As Integer
                
                Tempi = Int(mxl / yfactor) - 2
                
                'pline = Screen.List(Tempi)
                pline = Screen.List(Int(y / yfactor))
                
                'Screen.RemoveItem (Tempi)
                Screen.RemoveItem (Int(y / yfactor))
                
                promptline = promptline - 1
                If promptline < 0 Then promptline = 0
                
            End If
        End If
        
        update_xy
        Me.Line (x, y)-(x + (xfactor * 2), y + yfactor), Me.BackColor, BF
        
        update_xy
        
        Cursor.Enabled = True
        Exit Sub
    End If
    
    remove_cursor
    update_xy
    Me.Print Chr(KeyAscii);
    x = x + xfactor
    pline = pline + Chr(KeyAscii)
    
    If x + xfactor >= mxc Then
                
        remove_cursor
        
        x = 0
        y = y + yfactor
            
        If promptline = 0 Then
            temp_prompt = Prompt
        Else
            temp_prompt = ""
        End If
             
        Screen.AddItem temp_prompt & pline
        
        pline = ""
        
        If y >= mxl Then
            x = 0: y = 0
            Me.Cls
            
            Screen.RemoveItem 0
            update_xy
        
            For i = 0 To Int(mxl / yfactor)
                    Me.Print Screen.List(i);
                    y = y + yfactor: x = 0
                    update_xy
            Next
            y = mxl - yfactor
        End If

        CursorFlash = 0

        promptline = promptline + 1
    End If
    
    Cursor.Enabled = True
End Sub

Private Sub Form_Load()
    Form1.AutoRedraw = True
    
    Prompt = "->"
    
    CursorFlash = 0
    x = 0: y = 0
    
    mxc = lastxpos()
    mxl = lastypos()
            
    Me.Cls
    Me.ForeColor = 0
    promptline = 0
    pline = ""
    
    xfactor = Me.TextWidth("X")
    yfactor = Me.TextHeight("X")

    Me.Print Prompt;
            
    x = xfactor * 2
    update_xy
    
    Form1.Refresh
        
    Form1.AutoRedraw = False
    
End Sub
Sub remove_cursor()
    Me.Line (x, y)-(x + xfactor, y + yfactor), Me.BackColor, BF
    update_xy
End Sub
Sub update_xy()
    Me.CurrentX = x
    Me.CurrentY = y
End Sub
Function lastxpos()
    Dim char As Variant
            
    If xfactor = 0 Then
        xfactor = Me.TextWidth("X")
        yfactor = Me.TextHeight("X")
    End If
    
    Do
        char = char + xfactor
        If char > screenwidth Then
            char = char - xfactor
            Exit Do
        End If
    Loop
    lastxpos = char
End Function
Function lastypos()
    Dim char As Variant
        
    Do
        char = char + yfactor
        If char > screenheight Then
            char = char - yfactor
            Exit Do
        End If
    Loop
    lastypos = char - (5 * yfactor)
End Function

Private Sub Form_Resize()
    screenwidth = Me.Width
    screenheight = Me.Height
    
    mxc = lastxpos()
    mxl = lastypos()

End Sub
