VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "��������̨�� v1.0"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   1620
      ScaleWidth      =   5520
      TabIndex        =   0
      Top             =   0
      Width           =   5550
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '����esc�˳�����
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = Picture1.Width
    Me.Height = Picture1.Height '���ô�����
    Me.KeyPreview = True '���������а����¼���ִ��me�İ����¼�
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub
