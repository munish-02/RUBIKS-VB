VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1335
      ScaleWidth      =   975
      TabIndex        =   70
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   975
      TabIndex        =   69
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   12360
      ScaleHeight     =   1455
      ScaleWidth      =   975
      TabIndex        =   68
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   12360
      ScaleHeight     =   1575
      ScaleWidth      =   1215
      TabIndex        =   67
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3600
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton Command101 
      BackColor       =   &H00808000&
      Caption         =   "Solve"
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9360
      Width           =   1215
   End
   Begin VB.PictureBox F2 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   65
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox F3 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   64
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox F4 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   63
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox F5 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   62
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox F6 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   61
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox F7 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   60
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox F8 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   59
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox F9 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   58
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox F1 
      BackColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   57
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox D1 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   56
      Top             =   6360
      Width           =   700
   End
   Begin VB.PictureBox D2 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   55
      Top             =   6360
      Width           =   700
   End
   Begin VB.PictureBox D3 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   54
      Top             =   6360
      Width           =   700
   End
   Begin VB.PictureBox D4 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   53
      Top             =   7080
      Width           =   700
   End
   Begin VB.PictureBox D5 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   52
      Top             =   7080
      Width           =   700
   End
   Begin VB.PictureBox D6 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   51
      Top             =   7080
      Width           =   700
   End
   Begin VB.PictureBox D7 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   50
      Top             =   7800
      Width           =   700
   End
   Begin VB.PictureBox D8 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   49
      Top             =   7800
      Width           =   700
   End
   Begin VB.PictureBox D9 
      BackColor       =   &H00008000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   48
      Top             =   7800
      Width           =   700
   End
   Begin VB.PictureBox R2 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   2640
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   47
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox R3 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   3360
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   46
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox R4 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   1920
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   45
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox R7 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   1920
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   44
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox R5 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   2640
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   43
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox R6 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   3360
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   42
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox R8 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   2640
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   41
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox R9 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   3360
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   40
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox R1 
      BackColor       =   &H000080FF&
      Height          =   700
      Left            =   1920
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   39
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox B2 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   9600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   38
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox B3 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   10320
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   37
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox B4 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   8880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   36
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox B5 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   9600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   35
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox B6 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   10320
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   34
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox B7 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   8880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   33
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox B8 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   9600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   32
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox B9 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   10320
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   31
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox B1 
      BackColor       =   &H0000FFFF&
      Height          =   700
      Left            =   8880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   30
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox U1 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   29
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox U2 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   28
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox U3 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   27
      Top             =   120
      Width           =   700
   End
   Begin VB.PictureBox U4 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   26
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox U5 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   25
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox U6 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   24
      Top             =   840
      Width           =   700
   End
   Begin VB.PictureBox U7 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5160
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   23
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox U8 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   5880
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   22
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox U9 
      BackColor       =   &H00800000&
      Height          =   700
      Left            =   6600
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   21
      Top             =   1560
      Width           =   700
   End
   Begin VB.PictureBox L2 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   10680
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   20
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox L3 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   11400
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   19
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox L4 
      BackColor       =   &H00000080&
      Height          =   705
      Left            =   9960
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   18
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox L5 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   10680
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   17
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox L8 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   10680
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   16
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox L7 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   9960
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   15
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox L6 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   11400
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   14
      Top             =   4200
      Width           =   700
   End
   Begin VB.PictureBox L9 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   11400
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   13
      Top             =   4920
      Width           =   700
   End
   Begin VB.PictureBox L1 
      BackColor       =   &H00000080&
      Height          =   700
      Left            =   9960
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   12
      Top             =   3480
      Width           =   700
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   8280
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   11
      Top             =   8280
      Width           =   1100
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3600
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   10
      Top             =   8520
      Width           =   1100
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   10920
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   8280
      Width           =   1100
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   9600
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   8280
      Width           =   1100
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   8280
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   6120
      Width           =   1100
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   10920
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   6120
      Width           =   1100
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   9600
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   6120
      Width           =   1100
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2280
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   8520
      Width           =   1100
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   960
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   8520
      Width           =   1100
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2280
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   960
      ScaleHeight     =   1935
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   7
      X1              =   5160
      X2              =   6240
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   6
      X1              =   7320
      X2              =   8400
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   5
      X1              =   12120
      X2              =   8400
      Y1              =   5520
      Y2              =   4920
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   4
      X1              =   12120
      X2              =   8280
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   3
      X1              =   5400
      X2              =   1920
      Y1              =   5160
      Y2              =   5520
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   2
      X1              =   6240
      X2              =   1920
      Y1              =   2760
      Y2              =   3480
   End
   Begin VB.Line Line13 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5760
      X2              =   5160
      Y1              =   5640
      Y2              =   8400
   End
   Begin VB.Line Line12 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   7320
      X2              =   7320
      Y1              =   5640
      Y2              =   6360
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   1
      X1              =   8400
      X2              =   7320
      Y1              =   4920
      Y2              =   8520
   End
   Begin VB.Line Line11 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   0
      X1              =   5160
      X2              =   5160
      Y1              =   5640
      Y2              =   6360
   End
   Begin VB.Line Line10 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5160
      X2              =   4080
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line9 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5160
      X2              =   4080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   6240
      X2              =   8880
      Y1              =   2760
      Y2              =   120
   End
   Begin VB.Line Line7 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   8400
      X2              =   11040
      Y1              =   4920
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   7320
      X2              =   9960
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   7320
      X2              =   9960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line5 
      BorderStyle     =   5  'Dash-Dot-Dot
      DrawMode        =   9  'Not Mask Pen
      Index           =   1
      X1              =   7320
      X2              =   7320
      Y1              =   3480
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderStyle     =   5  'Dash-Dot-Dot
      DrawMode        =   9  'Not Mask Pen
      Index           =   0
      X1              =   5160
      X2              =   5160
      Y1              =   3600
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   2
      X1              =   7680
      X2              =   7680
      Y1              =   3240
      Y2              =   5400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   8040
      X2              =   8040
      Y1              =   3000
      Y2              =   5160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   8400
      X2              =   8400
      Y1              =   2760
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   6
      X1              =   7320
      X2              =   8400
      Y1              =   4200
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   5
      X1              =   7320
      X2              =   8400
      Y1              =   4920
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   4
      X1              =   7320
      X2              =   8400
      Y1              =   5640
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   5520
      X2              =   7680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   5880
      X2              =   8040
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   0
      X1              =   6240
      X2              =   8400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   6600
      X2              =   7680
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   7320
      X2              =   8400
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   5880
      X2              =   6960
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   5160
      X2              =   6240
      Y1              =   3480
      Y2              =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Mr()
a = F6.BackColor
b = F5.BackColor
c = F4.BackColor
D = R6.BackColor
e = R5.BackColor
F = R4.BackColor
g = B4.BackColor
h = B5.BackColor
i = B6.BackColor
j = L6.BackColor
k = L5.BackColor
L = L4.BackColor

F6.BackColor = j
F5.BackColor = k
F4.BackColor = L
R6.BackColor = a
R5.BackColor = b
R4.BackColor = c
B4.BackColor = D
B5.BackColor = e
B6.BackColor = F
L6.BackColor = g
L5.BackColor = h
L4.BackColor = i
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Ml()
For abc = 1 To 3
a = F6.BackColor
b = F5.BackColor
c = F4.BackColor
D = R6.BackColor
e = R5.BackColor
F = R4.BackColor
g = B4.BackColor
h = B5.BackColor
i = B6.BackColor
j = L6.BackColor
k = L5.BackColor
L = L4.BackColor

F6.BackColor = j
F5.BackColor = k
F4.BackColor = L
R6.BackColor = a
R5.BackColor = b
R4.BackColor = c
B4.BackColor = D
B5.BackColor = e
B6.BackColor = F
L6.BackColor = g
L5.BackColor = h
L4.BackColor = i
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub

Private Sub Mu()
For abc = 1 To 3
a = F2.BackColor
b = F5.BackColor
c = F8.BackColor
D = D2.BackColor
e = D5.BackColor
F = D8.BackColor
g = B8.BackColor
h = B5.BackColor
i = B2.BackColor
j = U2.BackColor
k = U5.BackColor
L = U8.BackColor

F2.BackColor = j
F5.BackColor = k
F8.BackColor = L
D2.BackColor = a
D5.BackColor = b
D8.BackColor = c
B8.BackColor = D
B5.BackColor = e
B2.BackColor = F
U2.BackColor = g
U5.BackColor = h
U8.BackColor = i
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Md()
a = F2.BackColor
b = F5.BackColor
c = F8.BackColor
D = D2.BackColor
e = D5.BackColor
F = D8.BackColor
g = B8.BackColor
h = B5.BackColor
i = B2.BackColor
j = U2.BackColor
k = U5.BackColor
L = U8.BackColor

F2.BackColor = j
F5.BackColor = k
F8.BackColor = L
D2.BackColor = a
D5.BackColor = b
D8.BackColor = c
B8.BackColor = D
B5.BackColor = e
B2.BackColor = F
U2.BackColor = g
U5.BackColor = h
U8.BackColor = i
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Di()
For abc = 1 To 3
a = F7.BackColor
b = F8.BackColor
c = F9.BackColor
D = L7.BackColor
e = L8.BackColor
F = L9.BackColor
g = B9.BackColor
h = B8.BackColor
i = B7.BackColor
j = R7.BackColor
k = R8.BackColor
L = R9.BackColor

m = D1.BackColor
n = D2.BackColor
o = D3.BackColor
p = D4.BackColor
R = D6.BackColor
s = D7.BackColor
t = D8.BackColor
u = D9.BackColor

F7.BackColor = j
F8.BackColor = k
F9.BackColor = L
L7.BackColor = a
L8.BackColor = b
L9.BackColor = c
B9.BackColor = D
B8.BackColor = e
B7.BackColor = F
R7.BackColor = g
R8.BackColor = h
R9.BackColor = i
D1.BackColor = s
D2.BackColor = p
D3.BackColor = m
D4.BackColor = t
D6.BackColor = n
D7.BackColor = u
D8.BackColor = R
D9.BackColor = o
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Fi()
For abc = 1 To 3
a = U7.BackColor
b = U8.BackColor
c = U9.BackColor
D = L1.BackColor
e = L4.BackColor
F = L7.BackColor
g = D3.BackColor
h = D2.BackColor
i = D1.BackColor
j = R9.BackColor
k = R6.BackColor
L = R3.BackColor

m = F1.BackColor
n = F2.BackColor
o = F3.BackColor
p = F4.BackColor
q = F6.BackColor
R = F7.BackColor
s = F8.BackColor
t = F9.BackColor

U7.BackColor = j
U8.BackColor = k
U9.BackColor = L
L1.BackColor = a
L4.BackColor = b
L7.BackColor = c
D3.BackColor = D
D2.BackColor = e
D1.BackColor = F
R9.BackColor = g
R6.BackColor = h
R3.BackColor = i
F1.BackColor = R
F2.BackColor = p
F3.BackColor = m
F4.BackColor = s
F6.BackColor = n
F7.BackColor = t
F8.BackColor = q
F9.BackColor = o
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub FF()
a = U7.BackColor
b = U8.BackColor
c = U9.BackColor
D = L1.BackColor
e = L4.BackColor
F = L7.BackColor
g = D3.BackColor
h = D2.BackColor
i = D1.BackColor
j = R9.BackColor
k = R6.BackColor
L = R3.BackColor

m = F1.BackColor
n = F2.BackColor
o = F3.BackColor
p = F4.BackColor
q = F6.BackColor
R = F7.BackColor
s = F8.BackColor
t = F9.BackColor

U7.BackColor = j
U8.BackColor = k
U9.BackColor = L
L1.BackColor = a
L4.BackColor = b
L7.BackColor = c
D3.BackColor = D
D2.BackColor = e
D1.BackColor = F
R9.BackColor = g
R6.BackColor = h
R3.BackColor = i
F1.BackColor = R
F2.BackColor = p
F3.BackColor = m
F4.BackColor = s
F6.BackColor = n
F7.BackColor = t
F8.BackColor = q
F9.BackColor = o
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk


End Sub
Private Sub Bi()
a = U1.BackColor
b = U2.BackColor
c = U3.BackColor
D = L3.BackColor
e = L6.BackColor
F = L9.BackColor
g = D9.BackColor
h = D8.BackColor
i = D7.BackColor
j = R7.BackColor
k = R4.BackColor
L = R1.BackColor

m = B1.BackColor
n = B2.BackColor
o = B3.BackColor
p = B4.BackColor
R = B6.BackColor
s = B7.BackColor
t = B8.BackColor
u = B9.BackColor

U1.BackColor = j
U2.BackColor = k
U3.BackColor = L
L3.BackColor = a
L6.BackColor = b
L9.BackColor = c
D9.BackColor = D
D8.BackColor = e
D7.BackColor = F
R7.BackColor = g
R4.BackColor = h
R1.BackColor = i

B1.BackColor = s
B2.BackColor = p
B3.BackColor = m
B4.BackColor = t
B6.BackColor = n
B7.BackColor = u
B8.BackColor = R
B9.BackColor = o
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub BB()
For abc = 1 To 3
a = U1.BackColor
b = U2.BackColor
c = U3.BackColor
D = L3.BackColor
e = L6.BackColor
F = L9.BackColor
g = D9.BackColor
h = D8.BackColor
i = D7.BackColor
j = R7.BackColor
k = R4.BackColor
L = R1.BackColor

m = B1.BackColor
n = B2.BackColor
o = B3.BackColor
p = B4.BackColor
R = B6.BackColor
s = B7.BackColor
t = B8.BackColor
u = B9.BackColor

U1.BackColor = j
U2.BackColor = k
U3.BackColor = L
L3.BackColor = a
L6.BackColor = b
L9.BackColor = c
D9.BackColor = D
D8.BackColor = e
D7.BackColor = F
R7.BackColor = g
R4.BackColor = h
R1.BackColor = i

B1.BackColor = s
B2.BackColor = p
B3.BackColor = m
B4.BackColor = t
B6.BackColor = n
B7.BackColor = u
B8.BackColor = R
B9.BackColor = o
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Ui()
a = F1.BackColor
b = F2.BackColor
c = F3.BackColor
D = L1.BackColor
e = L2.BackColor
F = L3.BackColor
g = B3.BackColor
h = B2.BackColor
i = B1.BackColor
j = R1.BackColor
k = R2.BackColor
L = R3.BackColor

m = U1.BackColor
n = U2.BackColor
o = U3.BackColor
p = U4.BackColor
R = U6.BackColor
s = U7.BackColor
t = U8.BackColor
u = U9.BackColor

F1.BackColor = j
F2.BackColor = k
F3.BackColor = L
L1.BackColor = a
L2.BackColor = b
L3.BackColor = c
B3.BackColor = D
B2.BackColor = e
B1.BackColor = F
R1.BackColor = g
R2.BackColor = h
R3.BackColor = i

U1.BackColor = o
U2.BackColor = R
U3.BackColor = u
U4.BackColor = n
U6.BackColor = t
U7.BackColor = m
U8.BackColor = p
U9.BackColor = s
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub UU()
For abc = 1 To 3
a = F1.BackColor
b = F2.BackColor
c = F3.BackColor
D = L1.BackColor
e = L2.BackColor
F = L3.BackColor
g = B3.BackColor
h = B2.BackColor
i = B1.BackColor
j = R1.BackColor
k = R2.BackColor
L = R3.BackColor

m = U1.BackColor
n = U2.BackColor
o = U3.BackColor
p = U4.BackColor
R = U6.BackColor
s = U7.BackColor
t = U8.BackColor
u = U9.BackColor

F1.BackColor = j
F2.BackColor = k
F3.BackColor = L
L1.BackColor = a
L2.BackColor = b
L3.BackColor = c
B3.BackColor = D
B2.BackColor = e
B1.BackColor = F
R1.BackColor = g
R2.BackColor = h
R3.BackColor = i

U1.BackColor = o
U2.BackColor = R
U3.BackColor = u
U4.BackColor = n
U6.BackColor = t
U7.BackColor = m
U8.BackColor = p
U9.BackColor = s
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub DD()
a = F7.BackColor
b = F8.BackColor
c = F9.BackColor
D = L7.BackColor
e = L8.BackColor
F = L9.BackColor
g = B9.BackColor
h = B8.BackColor
i = B7.BackColor
j = R7.BackColor
k = R8.BackColor
L = R9.BackColor

m = D1.BackColor
n = D2.BackColor
o = D3.BackColor
p = D4.BackColor
R = D6.BackColor
s = D7.BackColor
t = D8.BackColor
u = D9.BackColor

F7.BackColor = j
F8.BackColor = k
F9.BackColor = L
L7.BackColor = a
L8.BackColor = b
L9.BackColor = c
B9.BackColor = D
B8.BackColor = e
B7.BackColor = F
R7.BackColor = g
R8.BackColor = h
R9.BackColor = i
D1.BackColor = s
D2.BackColor = p
D3.BackColor = m
D4.BackColor = t
D6.BackColor = n
D7.BackColor = u
D8.BackColor = R
D9.BackColor = o
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Li()
For abc = 1 To 3
a = F1.BackColor
b = F4.BackColor
c = F7.BackColor
D = D1.BackColor
e = D4.BackColor
F = D7.BackColor
g = B7.BackColor
h = B4.BackColor
i = B1.BackColor
j = U1.BackColor
k = U4.BackColor
L = U7.BackColor
m = R1.BackColor
n = R2.BackColor
o = R3.BackColor
p = R4.BackColor
R = R6.BackColor
s = R7.BackColor
t = R8.BackColor
u = R9.BackColor

F1.BackColor = j
F4.BackColor = k
F7.BackColor = L
D1.BackColor = a
D4.BackColor = b
D7.BackColor = c
B7.BackColor = D
B4.BackColor = e
B1.BackColor = F
U1.BackColor = g
U4.BackColor = h
U7.BackColor = i
R1.BackColor = s
R2.BackColor = p
R3.BackColor = m
R4.BackColor = t
R6.BackColor = n
R7.BackColor = u
R8.BackColor = R
R9.BackColor = o
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub LL()
a = F1.BackColor
b = F4.BackColor
c = F7.BackColor
D = D1.BackColor
e = D4.BackColor
F = D7.BackColor
g = B7.BackColor
h = B4.BackColor
i = B1.BackColor
j = U1.BackColor
k = U4.BackColor
L = U7.BackColor
m = R1.BackColor
n = R2.BackColor
o = R3.BackColor
p = R4.BackColor
R = R6.BackColor
s = R7.BackColor
t = R8.BackColor
u = R9.BackColor

F1.BackColor = j
F4.BackColor = k
F7.BackColor = L
D1.BackColor = a
D4.BackColor = b
D7.BackColor = c
B7.BackColor = D
B4.BackColor = e
B1.BackColor = F
U1.BackColor = g
U4.BackColor = h
U7.BackColor = i
R1.BackColor = s
R2.BackColor = p
R3.BackColor = m
R4.BackColor = t
R6.BackColor = n
R7.BackColor = u
R8.BackColor = R
R9.BackColor = o
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub Ri()
For abc = 1 To 3
a = F9.BackColor
b = F6.BackColor
c = F3.BackColor
D = U9.BackColor
e = U6.BackColor
F = U3.BackColor
g = B3.BackColor
h = B6.BackColor
i = B9.BackColor
j = D9.BackColor
k = D6.BackColor
L = D3.BackColor
m = L1.BackColor
n = L2.BackColor
o = L3.BackColor
p = L4.BackColor
R = L6.BackColor
s = L7.BackColor
t = L8.BackColor
u = L9.BackColor


F9.BackColor = j
F6.BackColor = k
F3.BackColor = L
U9.BackColor = a
U6.BackColor = b
U3.BackColor = c
B3.BackColor = D
B6.BackColor = e
B9.BackColor = F
D9.BackColor = g
D6.BackColor = h
D3.BackColor = i
L1.BackColor = s
L2.BackColor = p
L3.BackColor = m
L4.BackColor = t
L6.BackColor = n
L7.BackColor = u
L8.BackColor = R
L9.BackColor = o
Next abc
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub
Private Sub RR()

a = F9.BackColor
b = F6.BackColor
c = F3.BackColor
D = U9.BackColor
e = U6.BackColor
F = U3.BackColor
g = B3.BackColor
h = B6.BackColor
i = B9.BackColor
j = D9.BackColor
k = D6.BackColor
L = D3.BackColor
m = L1.BackColor
n = L2.BackColor
o = L3.BackColor
p = L4.BackColor
R = L6.BackColor
s = L7.BackColor
t = L8.BackColor
u = L9.BackColor


F9.BackColor = j
F6.BackColor = k
F3.BackColor = L
U9.BackColor = a
U6.BackColor = b
U3.BackColor = c
B3.BackColor = D
B6.BackColor = e
B9.BackColor = F
D9.BackColor = g
D6.BackColor = h
D3.BackColor = i
L1.BackColor = s
L2.BackColor = p
L3.BackColor = m
L4.BackColor = t
L6.BackColor = n
L7.BackColor = u
L8.BackColor = R
L9.BackColor = o
For kkk = 1 To 10 Step 0.0001
muni = muni + i
Next kkk

End Sub



Private Sub Command125_Click()
Menu.Show


End Sub



Private Sub Command101_Click()
For mmm = 1 To 2
For i = 1 To 24
  If U1.BackColor = U2.BackColor And U2.BackColor = U3.BackColor And U4.BackColor = U5.BackColor And U5.BackColor = U6.BackColor And U6.BackColor = U7.BackColor And U7.BackColor = U9.BackColor And U9.BackColor = U8.BackColor And U8.BackColor = U1.BackColor And F5.BackColor = F1.BackColor And F1.BackColor = F3.BackColor And F2.BackColor = F1.BackColor And R3.BackColor = R5.BackColor And R5.BackColor = R2.BackColor And R2.BackColor = R1.BackColor And L1.BackColor = L2.BackColor And L2.BackColor = L3.BackColor And L3.BackColor = L5.BackColor Then
  Exit For
  
   Else
          UUU = U5.BackColor
          FFF = F5.BackColor
          RRR = R5.BackColor
          LLL = L5.BackColor
             For m = 1 To 4
                If UUU = U6.BackColor And FFF = L2.BackColor Then
                   Call Ri
                   Call Di
                   Call Mr
                   Call DD
                   Call RR
                   Call Di
                   Call Mr
                   Call DD
                   Call Fi
                   Call Di
                   Call Ml
                   Call DD
                   Call FF
                   Call Di
                   Call Ml
                   Call DD
                   xyz = xyz & " Ri Di Mr D R Di Mr D fi Di Ml D F Di Ml D "
                   ElseIf UUU = L2.BackColor And FFF = U6.BackColor Then
                   Call Ri
                   Call Ri
                   Call Md
                   Call Ri
                   Call RR
                   Call Di
                   Call Ri
                   Call Mu
                   Call RR
                   xyz = xyz & " Ri Ri Md Ri R Di Ri Mu R "
                   ElseIf FFF = U8.BackColor And UUU = F2.BackColor Then
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " Md Di Di Mu Di Md D Mu "
                   ElseIf FFF = B2.BackColor And UUU = U2.BackColor Then
                   Call Mu
                   Call Di
                   Call Di
                   Call Md
                   Call Di
                   Call Di
                   Call Md
                   Call DD
                   Call DD
                   Call Mu
                   xyz = xyz & " Mu Di Di Md Di Di Di Md D Mu "
                   ElseIf FFF = U2.BackColor And UUU = B2.BackColor Then
                   Call Mu
                   Call Di
                   Call Di
                   Call Md
                   Call Di
                   Call Di
                   Call Md
                   Call DD
                   Call DD
                   Call Mu
                   xyz = xyz & " Mu Di Di Md Di Di Di Md D Mu "
                   ElseIf FFF = R2.BackColor And UUU = U4.BackColor Then
                   Call LL
                   Call Ml
                   Call Li
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   xyz = xyz & " L Ml Li F Mr Fi Mr "
                   ElseIf FFF = U4.BackColor And UUU = R2.BackColor Then
                   Call LL
                   Call Ml
                   Call Li
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " L Ml Li F Mr Fi Mr Md Di Di Mu Di Md D Mu "
                   ElseIf FFF = F8.BackColor And UUU = D2.BackColor Then
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   xyz = xyz & " Md Di Di Mu "
                   ElseIf FFF = D2.BackColor And UUU = F8.BackColor Then
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " Di Md D Mu "
                   ElseIf FFF = L4.BackColor And UUU = F6.BackColor Then
                   Call Mr
                   Call Mr
                   Call Fi
                   Call Ml
                   Call FF
                   Call Ml
                   xyz = xyz & " Mr Mr Fi Ml F Ml "
                   ElseIf FFF = F6.BackColor And UUU = L4.BackColor Then
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   xyz = xyz & " Ml F Mr Fi "
                   ElseIf FFF = F4.BackColor And UUU = R6.BackColor Then
                   Call Mr
                   Call Fi
                   Call Ml
                   Call FF
                   xyz = xyz & " Mr Fi Ml F "
                   ElseIf FFF = R6.BackColor And UUU = F4.BackColor Then
                   Call Ml
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   xyz = xyz & " Ml Ml F Mr Fi Mr "
               
                   ElseIf FFF = R3.BackColor And UUU = F1.BackColor And RRR = U7.BackColor Then
                     Call LL
                     Call Di
                     Call Li
                     Call DD
                     Call LL
                     Call Di
                     Call Li
                     xyz = xyz & " L Di Li D L Di Di "
                    
                     ElseIf UUU = R3.BackColor And FFF = U7.BackColor And RRR = F1.BackColor Then
                     Call Fi
                     Call DD
                     Call FF
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Fi D F L D Li "
                     ElseIf UUU = F7.BackColor And FFF = D1.BackColor And RRR = R9.BackColor Then
                     Call Fi
                     Call Di
                     Call FF
                     xyz = xyz & " Fi Di F "
                     ElseIf UUU = D1.BackColor And FFF = R9.BackColor And RRR = F9.BackColor Then
                     Call LL
                     Call Di
                     Call Di
                     Call Li
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " L Di Di Li Di L D Li "
                     ElseIf UUU = F3.BackColor And FFF = U9.BackColor And RRR = L1.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Ri D R Di L D Li "
                     ElseIf UUU = F9.BackColor And FFF = L7.BackColor Then
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Di L D Li "
                     ElseIf FFF = F9.BackColor And UUU = D3.BackColor Then
                     Call Ri
                     Call Di
                     Call Di
                     Call Fi
                     Call Di
                     Call FF
                     xyz = xyz & " Ri Di Di Fi Di F "
                     ElseIf L1.BackColor = FFF And F3.BackColor = RRR Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf UUU = L1.BackColor And FFF = F3.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf FFF = U9.BackColor And UUU = F3.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf UUU = F3.BackColor And RRR = L3.BackColor Or B3.BackColor = UUU And RRR = F3.BackColor Or FFF = F3.BackColor And UUU = L3.BackColor Then
                     Call Bi
                     Call DD
                     Call BB
                     xyz = xyz & " Bi D B "
                     ElseIf UUU = U1.BackColor And FFF = R1.BackColor Or U1.BackColor = FFF And B1.BackColor = UUU Or U1.BackColor = RRR And R1.BackColor = UUU Then
                     Call BB
                     Call Di
                     Call Bi
                     xyz = xyz & " B Di Bi"
                     ElseIf U7.BackColor = RRR And R3.BackColor = FFF Or UUU = R3.BackColor And F1.BackColor = RRR Then
                     Call LL
                     Call Di
                     Call Li
                     xyz = xyz & " L Di Li "
                     
                   End If
                Call Di
                Call Mr
                xyz = xyz & " Di Mr "
                
            Next m
        End If
        Call UU
        Call Di
        Call Mr
        xyz = xyz & " U Di Mr "

Next i
Call RR
Call RR
Call Li
Call Li
Call Md
Call Md
xyz = xyz & " Invert teh cube "
For i = 1 To 60
If F4.BackColor = F5.BackColor And F5.BackColor = F6.BackColor And L4.BackColor = L5.BackColor And L5.BackColor = L6.BackColor And R4.BackColor = R5.BackColor And R5.BackColor = R6.BackColor And B4.BackColor = B5.BackColor And B5.BackColor = B6.BackColor Then
Exit For
Else
   For j = 1 To 4


  
   If F2.BackColor = F5.BackColor And U8.BackColor = L5.BackColor Then
   xyz = xyz & "U R Ui Ri Ui Fi U F"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   ElseIf F2.BackColor = L5.BackColor And U8.BackColor = F5.BackColor Then
   Call Di
   Call Mr
   xyz = xyz & " Di Mr "
   xyz = xyz & " Ui Li U L U F Ui Fi"
   Call Ui
   Call Li
   Call UU
   Call LL
   Call UU
   Call FF
   Call Ui
   Call Fi
   Call DD
   Call Ml
   xyz = xyz & "  D Ml "
   ElseIf F5.BackColor = L4.BackColor And L5.BackColor = F6.BackColor Then
   xyz = xyz & "U R Ui Ri Ui Fi U F"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   ElseIf F5.BackColor = B6.BackColor And L5.BackColor = L6.BackColor Or L5.BackColor = B6.BackColor And F5.BackColor = L6.BackColor Then
   Call UU
   Call Di
   Call Mr
   xyz = xyz & " U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   ElseIf F5.BackColor = R4.BackColor And L5.BackColor = B4.BackColor Or R4.BackColor = L5.BackColor And F5.BackColor = B4.BackColor Then
   Call UU
   Call Di
   Call Mr
   Call UU
   Call Di
   Call Mr
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   Call Ui
   Call DD
   Call Ml
   
   xyz = xyz & " U Di Mr U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml Ui DD Ml "
   ElseIf F5.BackColor = R6.BackColor And R5.BackColor = F4.BackColor Then
    Call UU
   Call Di
   Call Mr
    Call UU
   Call Di
   Call Mr
    Call UU
   Call Di
   Call Mr
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   Call Ui
   Call DD
   xyz = xyz & " U Di Mr U Di Mr U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml Ui DD Ml Ui DD Ml "
   Call Ml
  End If
   Call UU
   xyz = xyz & " U  "
   
   Next j
   End If
   Call UU
   Call Di
   Call Mr
   xyz = xyz & " U Di Mr "

   
   Next i
      For i = 1 To 4
   For j = 1 To 6
   If U2.BackColor = U5.BackColor And U5.BackColor = U8.BackColor And U5.BackColor = U4.BackColor And U5.BackColor = U6.BackColor Then
   Exit For
   
   ElseIf U2.BackColor <> U5.BackColor And U4.BackColor <> U5.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
  
   ElseIf U2.BackColor = U5.BackColor And U5.BackColor = U4.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   qqq = 1
   xyz = xyz & " F R U Ri Ui R U Ri Ui Fi "
   ElseIf U5.BackColor = U6.BackColor And U6.BackColor = U4.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
   qqq = 1
   ElseIf j = 5 And U2.BackColor <> U5.BackColor And U4.BackColor <> U5.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
   End If
   Call Ui
   xyz = xyz & " Ui "
   Next j
  Next i
   For i = 1 To 10
 If U5.BackColor = U1.BackColor And U5.BackColor = U2.BackColor And U5.BackColor = U7.BackColor And U7.BackColor = U9.BackColor Then
 Exit For
 Else
 For j = 1 To 50
 If U5.BackColor = U1.BackColor And U5.BackColor = U2.BackColor And U5.BackColor = U7.BackColor And U7.BackColor = U9.BackColor Then
 Exit For
 ElseIf U5.BackColor = U7.BackColor And U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And R1.BackColor = U5.BackColor And R3.BackColor = U5.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And F1.BackColor = U5.BackColor And F3.BackColor = U5.BackColor Then
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And L1.BackColor = U5.BackColor And L3.BackColor = U5.BackColor Then
 Call UU
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U U R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And B1.BackColor = U5.BackColor And B3.BackColor = U5.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = F1.BackColor And U5.BackColor = F3.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & "  R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = R1.BackColor And U5.BackColor = R3.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U1.BackColor = U5.BackColor And R3.BackColor = U5.BackColor And L1.BackColor = U5.BackColor Then
 
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 ElseIf U1.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = B3.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U1.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = L3.BackColor Then
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 End If
 Call UU
 Next j
 Call UU
 Call Di
 Call Mr
 xyz = xyz & " U Di Mr "
 End If
 Next i
 If U5.BackColor <> U1.BackColor Or U5.BackColor <> U2.BackColor Or U5.BackColor <> U3.BackColor Or U5.BackColor <> U6.BackColor Or U5.BackColor <> U4.BackColor Or U5.BackColor <> U7.BackColor Or U5.BackColor <> U8.BackColor Or U5.BackColor <> U9.BackColor Then
 For i = 1 To 24
  If U1.BackColor = U2.BackColor And U2.BackColor = U3.BackColor And U4.BackColor = U5.BackColor And U5.BackColor = U6.BackColor And U6.BackColor = U7.BackColor And U7.BackColor = U9.BackColor And U9.BackColor = U8.BackColor And U8.BackColor = U1.BackColor And F5.BackColor = F1.BackColor And F1.BackColor = F3.BackColor And F2.BackColor = F1.BackColor And R3.BackColor = R5.BackColor And R5.BackColor = R2.BackColor And R2.BackColor = R1.BackColor And L1.BackColor = L2.BackColor And L2.BackColor = L3.BackColor And L3.BackColor = L5.BackColor Then
  Exit For
  
   Else
          UUU = U5.BackColor
          FFF = F5.BackColor
          RRR = R5.BackColor
          LLL = L5.BackColor
             For m = 1 To 4
                If UUU = U6.BackColor And FFF = L2.BackColor Then
                   Call Ri
                   Call Di
                   Call Mr
                   Call DD
                   Call RR
                   Call Di
                   Call Mr
                   Call DD
                   Call Fi
                   Call Di
                   Call Ml
                   Call DD
                   Call FF
                   Call Di
                   Call Ml
                   Call DD
                   xyz = xyz & " Ri Di Mr D R Di Mr D fi Di Ml D F Di Ml D "
                   ElseIf UUU = L2.BackColor And FFF = U6.BackColor Then
                   Call Ri
                   Call Ri
                   Call Md
                   Call Ri
                   Call RR
                   Call Di
                   Call Ri
                   Call Mu
                   Call RR
                   xyz = xyz & " Ri Ri Md Ri R Di Ri Mu R "
                   ElseIf FFF = U8.BackColor And UUU = F2.BackColor Then
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " Md Di Di Mu Di Md D Mu "
                   ElseIf FFF = B2.BackColor And UUU = U2.BackColor Then
                   Call Mu
                   Call Di
                   Call Di
                   Call Md
                   Call Di
                   Call Di
                   Call Md
                   Call DD
                   Call DD
                   Call Mu
                   xyz = xyz & " Mu Di Di Md Di Di Di Md D Mu "
                   ElseIf FFF = U2.BackColor And UUU = B2.BackColor Then
                   Call Mu
                   Call Di
                   Call Di
                   Call Md
                   Call Di
                   Call Di
                   Call Md
                   Call DD
                   Call DD
                   Call Mu
                   xyz = xyz & " Mu Di Di Md Di Di Di Md D Mu "
                   ElseIf FFF = R2.BackColor And UUU = U4.BackColor Then
                   Call LL
                   Call Ml
                   Call Li
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   xyz = xyz & " L Ml Li F Mr Fi Mr "
                   ElseIf FFF = U4.BackColor And UUU = R2.BackColor Then
                   Call LL
                   Call Ml
                   Call Li
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " L Ml Li F Mr Fi Mr Md Di Di Mu Di Md D Mu "
                   ElseIf FFF = F8.BackColor And UUU = D2.BackColor Then
                   Call Md
                   Call Di
                   Call Di
                   Call Mu
                   xyz = xyz & " Md Di Di Mu "
                   ElseIf FFF = D2.BackColor And UUU = F8.BackColor Then
                   Call Di
                   Call Md
                   Call DD
                   Call Mu
                   xyz = xyz & " Di Md D Mu "
                   ElseIf FFF = L4.BackColor And UUU = F6.BackColor Then
                   Call Mr
                   Call Mr
                   Call Fi
                   Call Ml
                   Call FF
                   Call Ml
                   xyz = xyz & " Mr Mr Fi Ml F Ml "
                   ElseIf FFF = F6.BackColor And UUU = L4.BackColor Then
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   xyz = xyz & " Ml F Mr Fi "
                   ElseIf FFF = F4.BackColor And UUU = R6.BackColor Then
                   Call Mr
                   Call Fi
                   Call Ml
                   Call FF
                   xyz = xyz & " Mr Fi Ml F "
                   ElseIf FFF = R6.BackColor And UUU = F4.BackColor Then
                   Call Ml
                   Call Ml
                   Call FF
                   Call Mr
                   Call Fi
                   Call Mr
                   xyz = xyz & " Ml Ml F Mr Fi Mr "
               
                   ElseIf FFF = R3.BackColor And UUU = F1.BackColor And RRR = U7.BackColor Then
                     Call LL
                     Call Di
                     Call Li
                     Call DD
                     Call LL
                     Call Di
                     Call Li
                     xyz = xyz & " L Di Li D L Di Di "
                    
                     ElseIf UUU = R3.BackColor And FFF = U7.BackColor And RRR = F1.BackColor Then
                     Call Fi
                     Call DD
                     Call FF
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Fi D F L D Li "
                     ElseIf UUU = F7.BackColor And FFF = D1.BackColor And RRR = R9.BackColor Then
                     Call Fi
                     Call Di
                     Call FF
                     xyz = xyz & " Fi Di F "
                     ElseIf UUU = D1.BackColor And FFF = R9.BackColor And RRR = F9.BackColor Then
                     Call LL
                     Call Di
                     Call Di
                     Call Li
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " L Di Di Li Di L D Li "
                     ElseIf UUU = F3.BackColor And FFF = U9.BackColor And RRR = L1.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Ri D R Di L D Li "
                     ElseIf UUU = F9.BackColor And FFF = L7.BackColor Then
                     Call Di
                     Call LL
                     Call DD
                     Call Li
                     xyz = xyz & " Di L D Li "
                     ElseIf FFF = F9.BackColor And UUU = D3.BackColor Then
                     Call Ri
                     Call Di
                     Call Di
                     Call Fi
                     Call Di
                     Call FF
                     xyz = xyz & " Ri Di Di Fi Di F "
                     ElseIf L1.BackColor = FFF And F3.BackColor = RRR Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf UUU = L1.BackColor And FFF = F3.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf FFF = U9.BackColor And UUU = F3.BackColor Then
                     Call Ri
                     Call DD
                     Call RR
                     xyz = xyz & " Ri D R "
                     ElseIf UUU = F3.BackColor And RRR = L3.BackColor Or B3.BackColor = UUU And RRR = F3.BackColor Or FFF = F3.BackColor And UUU = L3.BackColor Then
                     Call Bi
                     Call DD
                     Call BB
                     xyz = xyz & " Bi D B "
                     ElseIf UUU = U1.BackColor And FFF = R1.BackColor Or U1.BackColor = FFF And B1.BackColor = UUU Or U1.BackColor = RRR And R1.BackColor = UUU Then
                     Call BB
                     Call Di
                     Call Bi
                     xyz = xyz & " B Di Bi"
                     ElseIf U7.BackColor = RRR And R3.BackColor = FFF Or UUU = R3.BackColor And F1.BackColor = RRR Then
                     Call LL
                     Call Di
                     Call Li
                     xyz = xyz & " L Di Li "
                     
                   End If
                Call Di
                Call Mr
                xyz = xyz & " Di Mr "
                
            Next m
        End If
        Call UU
        Call Di
        Call Mr
        xyz = xyz & " U Di Mr "

Next i
Call RR
Call RR
Call Li
Call Li
Call Md
Call Md
xyz = xyz & " Invert teh cube "
For i = 1 To 60
If F4.BackColor = F5.BackColor And F5.BackColor = F6.BackColor And L4.BackColor = L5.BackColor And L5.BackColor = L6.BackColor And R4.BackColor = R5.BackColor And R5.BackColor = R6.BackColor And B4.BackColor = B5.BackColor And B5.BackColor = B6.BackColor Then
Exit For
Else
   For j = 1 To 4


  
   If F2.BackColor = F5.BackColor And U8.BackColor = L5.BackColor Then
   xyz = xyz & "U R Ui Ri Ui Fi U F"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   ElseIf F2.BackColor = L5.BackColor And U8.BackColor = F5.BackColor Then
   Call Di
   Call Mr
   xyz = xyz & " Di Mr "
   xyz = xyz & " Ui Li U L U F Ui Fi"
   Call Ui
   Call Li
   Call UU
   Call LL
   Call UU
   Call FF
   Call Ui
   Call Fi
   Call DD
   Call Ml
   xyz = xyz & "  D Ml "
   ElseIf F5.BackColor = L4.BackColor And L5.BackColor = F6.BackColor Then
   xyz = xyz & "U R Ui Ri Ui Fi U F"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   ElseIf F5.BackColor = B6.BackColor And L5.BackColor = L6.BackColor Or L5.BackColor = B6.BackColor And F5.BackColor = L6.BackColor Then
   Call UU
   Call Di
   Call Mr
   xyz = xyz & " U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml"
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   ElseIf F5.BackColor = R4.BackColor And L5.BackColor = B4.BackColor Or R4.BackColor = L5.BackColor And F5.BackColor = B4.BackColor Then
   Call UU
   Call Di
   Call Mr
   Call UU
   Call Di
   Call Mr
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   Call Ui
   Call DD
   Call Ml
   
   xyz = xyz & " U Di Mr U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml Ui DD Ml "
   ElseIf F5.BackColor = R6.BackColor And R5.BackColor = F4.BackColor Then
    Call UU
   Call Di
   Call Mr
    Call UU
   Call Di
   Call Mr
    Call UU
   Call Di
   Call Mr
   Call UU
   Call RR
   Call Ui
   Call Ri
   Call Ui
   Call Fi
   Call UU
   Call FF
   Call Ui
   Call DD
   Call Ml
   Call Ui
   Call DD
   xyz = xyz & " U Di Mr U Di Mr U Di Mr U R Ui Ri Ui Fi U F Ui DD Ml Ui DD Ml Ui DD Ml "
   Call Ml
  End If
   Call UU
   xyz = xyz & " U  "
   
   Next j
   End If
   Call UU
   Call Di
   Call Mr
   xyz = xyz & " U Di Mr "

   
   Next i
      For i = 1 To 4
   For j = 1 To 6
   If U2.BackColor = U5.BackColor And U5.BackColor = U8.BackColor And U5.BackColor = U4.BackColor And U5.BackColor = U6.BackColor Then
   Exit For
   
   ElseIf U2.BackColor <> U5.BackColor And U4.BackColor <> U5.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
  
   ElseIf U2.BackColor = U5.BackColor And U5.BackColor = U4.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   qqq = 1
   xyz = xyz & " F R U Ri Ui R U Ri Ui Fi "
   ElseIf U5.BackColor = U6.BackColor And U6.BackColor = U4.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
   qqq = 1
   ElseIf j = 5 And U2.BackColor <> U5.BackColor And U4.BackColor <> U5.BackColor And U5.BackColor <> U6.BackColor And U5.BackColor <> U8.BackColor Then
   Call FF
   Call RR
   Call UU
   Call Ri
   Call Ui
   Call Fi
   xyz = xyz & " F R U Ri Ui Fi "
   End If
   Call Ui
   xyz = xyz & " Ui "
   Next j
  Next i
   For i = 1 To 10
 If U5.BackColor = U1.BackColor And U5.BackColor = U2.BackColor And U5.BackColor = U7.BackColor And U7.BackColor = U9.BackColor Then
 Exit For
 Else
 For j = 1 To 50
 If U5.BackColor = U1.BackColor And U5.BackColor = U2.BackColor And U5.BackColor = U7.BackColor And U7.BackColor = U9.BackColor Then
 Exit For
 ElseIf U5.BackColor = U7.BackColor And U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And R1.BackColor = U5.BackColor And R3.BackColor = U5.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And F1.BackColor = U5.BackColor And F3.BackColor = U5.BackColor Then
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And L1.BackColor = U5.BackColor And L3.BackColor = U5.BackColor Then
 Call UU
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U U R U Ri U R U U Ri "
 ElseIf U5.BackColor <> U1.BackColor And U5.BackColor <> U3.BackColor And U5.BackColor <> U7.BackColor And U5.BackColor <> U9.BackColor And B1.BackColor = U5.BackColor And B3.BackColor = U5.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = F1.BackColor And U5.BackColor = F3.BackColor Then
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & "  R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = R1.BackColor And U5.BackColor = R3.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U3.BackColor = U5.BackColor And U1.BackColor = U5.BackColor And R3.BackColor = U5.BackColor And L1.BackColor = U5.BackColor Then
 
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 ElseIf U1.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = B3.BackColor Then
 Call Ui
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " Ui R U Ri U R U U Ri "
 ElseIf U1.BackColor = U5.BackColor And U5.BackColor = U9.BackColor And U5.BackColor = L3.BackColor Then
 Call UU
 Call RR
 Call UU
 Call Ri
 Call UU
 Call RR
 Call UU
 Call UU
 Call Ri
 xyz = xyz & " U R U Ri U R U U Ri "
 End If
 Call UU
 Next j
 Call UU
 Call Di
 Call Mr
 xyz = xyz & " U Di Mr "
 End If
 Next i
 End If
 If B1.BackColor <> B3.BackColor And B1.BackColor <> B5.BackColor Then
For j = 1 To 8
If B1.BackColor = B3.BackColor And B3.BackColor = B5.BackColor Then
Exit For
Else
For i = 1 To 4
If B1.BackColor = B3.BackColor And B3.BackColor = B5.BackColor Then
Exit For
Else
Call UU
xyz = xyz & " U "
End If

Next i

End If
Call UU
Call Di
Call Mr
xyz = xyz & "U Di Mr "
Next j

End If

If B1.BackColor <> B3.BackColor Or B1.BackColor <> B5.BackColor Then
For j = 1 To 8
If B1.BackColor = B3.BackColor And B3.BackColor = B5.BackColor Then
Exit For
Else
For i = 1 To 4
If B1.BackColor = B3.BackColor And B3.BackColor = B5.BackColor Then
Exit For
Else
Call UU
xyz = xyz & " U "
End If

Next i

End If
Call UU
Call Di
Call Mr
xyz = xyz & "U Di Mr "
Next j
End If
If L1.BackColor <> L3.BackColor Or L1.BackColor <> L5.BackColor Or F1.BackColor <> F3.BackColor Or F1.BackColor <> F5.BackColor Or R1.BackColor <> R3.BackColor Or R5.BackColor <> R1.BackColor Then
xyz = xyz & " Ri F Ri B B R Fi Ri B B Ri Ri Ui "
Call Ri
Call FF
Call Ri
Call BB
Call BB
Call RR
Call Fi
Call Ri
Call BB
Call BB
Call Ri
Call Ri
Call Ui
End If
For i = 1 To 4
If B2.BackColor = B5.BackColor Then
qwe = 1
Exit For
Else
Call UU
Call Di
Call Mr
xyz = xyz & " U Di Mr "
End If
Next i
If qwe <> 1 Then
Call FF
Call FF
Call Ui
Call Ri
Call LL
Call FF
Call FF
Call RR
Call Li
Call Ui
Call FF
Call FF
xyz = xyz & " F F Ui Ri L F F R Li Ui F F "
End If
For i = 1 To 4
If B2.BackColor = B5.BackColor Then
qwe = 1
Exit For
Else
Call UU
Call Di
Call Mr
xyz = xyz & " U Di Mr "
End If
Next i
If F2.BackColor = L5.BackColor Then
Call FF
Call FF
Call Ui
Call Ri
Call LL
Call FF
Call FF
Call RR
Call Li
Call Ui
Call FF
Call FF
xyz = xyz & " F F Ui Ri L F F R Li Ui F F "
ElseIf F2.BackColor = R5.BackColor Then
Call FF
Call FF
Call UU
Call Ri
Call LL
Call FF
Call FF
Call RR
Call Li
Call UU
Call FF
Call FF
xyz = xyz & " F F UU Ri L F F R Li UU F F "
End If
 
 
 

Next mmm
mmm = MsgBox(xyz, vbInformation, "Solution")
End Sub

Private Sub Form_Load()

aF1 = F1.BackColor
aF2 = F2.BackColor
aF3 = F3.BackColor
aF4 = F4.BackColor
aF5 = F5.BackColor
aF6 = F6.BackColor
aF7 = F7.BackColor
aF8 = F8.BackColor
aF9 = F9.BackColor
aR1 = R1.BackColor
aR2 = R2.BackColor
aR3 = R3.BackColor
aR4 = R4.BackColor
aR6 = R6.BackColor
aR7 = R7.BackColor
aR8 = R8.BackColor
aR9 = R9.BackColor
aD1 = D1.BackColor
aD2 = D2.BackColor
aD3 = D3.BackColor
aD4 = D4.BackColor
aD5 = D5.BackColor
aD6 = D6.BackColor
aD7 = D7.BackColor
aD8 = D8.BackColor
aD9 = D9.BackColor
aL1 = L1.BackColor
aL2 = L2.BackColor
aL3 = L3.BackColor
aL4 = L4.BackColor
aL5 = L5.BackColor
aL6 = L6.BackColor
aL7 = L7.BackColor
aL8 = L8.BackColor
aL9 = L9.BackColor
aU1 = U1.BackColor
aU2 = U2.BackColor
aU3 = U3.BackColor
aU4 = U4.BackColor
aU5 = U5.BackColor
aU6 = U6.BackColor
aU7 = U7.BackColor
aU8 = U8.BackColor
aU9 = U9.BackColor
aB1 = B1.BackColor
aB2 = B2.BackColor
aB3 = B3.BackColor
aB4 = B4.BackColor
aB5 = B5.BackColor
aB6 = B6.BackColor
aB7 = B7.BackColor
aB8 = B8.BackColor
aB9 = B9.BackColor


End Sub

Private Sub Picture1_Click()

RR
End Sub

Private Sub Picture10_Click()
Fi
End Sub

Private Sub Picture11_Click()
UU

End Sub

Private Sub Picture12_Click()
Ui
End Sub


Private Sub Picture13_Click()
Mr
End Sub

Private Sub Picture14_Click()
Ml
End Sub

Private Sub Picture15_Click()
Md
End Sub

Private Sub Picture16_Click()
Mu

End Sub

Private Sub Picture2_Click()

Ri

End Sub

Private Sub Picture3_Click()
LL
End Sub

Private Sub Picture4_Click()
Li

End Sub

Private Sub Picture5_Click()
BB

End Sub

Private Sub Picture6_Click()
Bi
End Sub

Private Sub Picture7_Click()
DD


End Sub

Private Sub Picture8_Click()
Di
End Sub

Private Sub Picture9_Click()
FF


End Sub

Private Sub Text1_Click()
Md
End Sub

Private Sub Text2_Click()
Mu

End Sub

Private Sub Text3_Click()
Mr

End Sub

Private Sub Text4_Click()
Ml

End Sub


