VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SmcSplashScreenForm 
   Caption         =   "SMC Dry Cleaning"
   ClientHeight    =   6972
   ClientLeft      =   100
   ClientTop       =   460
   ClientWidth     =   7320
   OleObjectBlob   =   "SplashScreenForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SmcSplashScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SmcDryCleaningForm As String
Private Sub UserForm_Activate()
    
    SmcSplashScreenForm.LoadingLbl.Caption = "Welcome"
    SmcSplashScreenForm.CopyRightLbl.Caption = "Created by Marius Dragan on 07/10/2018."
    Application.Wait (Now + TimeValue("00:00:03"))
    SmcSplashScreenForm.LoadingLbl.Caption = "Loading Data..."
    SmcSplashScreenForm.CopyRightLbl.Caption = "Created by Marius Dragan on 07/10/2018."
    SmcSplashScreenForm.Repaint
    Application.Wait (Now + TimeValue("00:00:03"))
    SmcSplashScreenForm.LoadingLbl.Caption = "Creating Forms..."
    SmcSplashScreenForm.CopyRightLbl.Caption = "Copyright © 2018."
    SmcSplashScreenForm.Repaint
    Application.Wait (Now + TimeValue("00:00:03"))
    SmcSplashScreenForm.LoadingLbl.Caption = "Opening..."
    SmcSplashScreenForm.CopyRightLbl.Caption = "All rights reserved."
    SmcSplashScreenForm.Repaint
    Application.Wait (Now + TimeValue("00:00:02"))
    Unload SmcSplashScreenForm
    Application.Wait (Now + TimeValue("00:00:01"))
    DryCleaningForm.Show vbModeless
End Sub

Private Sub UserForm_Initialize()
    'Hide workbook when launched
    SmcDryCleaningForm = ThisWorkbook.Name
    wsAdmin.Visible = xlSheetVeryHidden
    
    If Workbooks.Count > 1 Then
            Windows(SmcDryCleaningForm).Visible = False
        Else
            Application.Visible = False
        End If
End Sub
