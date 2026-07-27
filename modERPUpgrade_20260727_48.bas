Attribute VB_Name = "modERPUpgrade_20260727_48"
Option Explicit
Option Compare Database

'=============================================================================
' modERPUpgrade_20260727_48
' QC Tab Button Completion
'
' Completes all empty and unlinked buttons in the QC left-navigation panel,
' applies the same visual design (colours, font, size) used by the MIR
' Register form buttons, and wires every button to its target form.
'
' New forms created by this upgrade
'   frmQCDashboard     – QC Dashboard (stub, ready for content)
'   frmITPRegister     – Inspection & Test Plan Register
'   frmNCRRegister     – Non-Conformance Report Register
'   frmTestReports     – Test Reports Register
'   frmQCReports       – QC Reports Hub
'
' Buttons completed in frmQCLeftMenu
'   btnQCDashboard     – opens frmQCDashboard
'   btnITPRegister     – opens frmITPRegister
'   btnMIR             – already linked; style harmonised
'   btnWIR             – connected by upgrade 47; style harmonised
'   btnNCR             – opens frmNCRRegister
'   btnTestReports     – opens frmTestReports
'   btnQCReports       – opens frmQCReports
'
' Version  : 2026.07.27.48
' Author   : G. Hayek
' Date     : 27-Jul-2026
'=============================================================================

Private Const UPGRADE_CODE      As String  = "2026.07.27.48"
Private Const UPGRADE_NAME      As String  = "Completed QC tab buttons with MIR-style design and navigation links"
Private Const QC_MENU_FORM      As String  = "frmQCLeftMenu"
Private Const MIR_FORM          As String  = "frmMIRRegister"
Private Const REF_BUTTON        As String  = "btnMIR"        ' Style source in frmQCLeftMenu
Private Const STANDARD_FONT     As String  = "Segoe UI"      ' ERP-standard font for all stub forms
Private Const FORM_DEFAULT_WIDTH As Long   = 9500            ' 9500 twips = 9500/1440 ≈ 6.6 inches
Private Const VBE_MAX_COLUMN    As Long    = 9999            ' Sentinel: search to end-of-line in CodeModule.Find

' ─── Entry points ────────────────────────────────────────────────────────────

Public Sub RunUpgrade_20260727_48()
    On Error GoTo EH

    If UpgradeAlreadyApplied48(UPGRADE_CODE) Then
        MsgBox "Upgrade " & UPGRADE_CODE & " is already applied.", _
               vbInformation, "ERP Upgrade"
        Exit Sub
    End If

    If Not FormExists48(QC_MENU_FORM) Then
        MsgBox "Cannot find """ & QC_MENU_FORM & """. " & _
               "Ensure the QC left-menu form exists before running this upgrade.", _
               vbCritical, "ERP Upgrade"
        Exit Sub
    End If

    Application.Echo False
    DoCmd.Hourglass True

    '-- 1. Create placeholder forms for unimplemented QC features
    CreateQCPlaceholderForms48

    '-- 2. Style + wire every navigation button in the QC left menu
    StyleAndLinkQCButtons48

    '-- 3. Record this upgrade in tblERPUpgradeHistory
    RecordUpgrade48

    DoCmd.Hourglass False
    Application.Echo True

    MsgBox "Upgrade " & UPGRADE_CODE & " completed." & vbCrLf & vbCrLf & _
           UPGRADE_NAME, vbInformation, "ERP Upgrade"
    Exit Sub

EH:
    DoCmd.Hourglass False
    Application.Echo True
    MsgBox "Upgrade " & UPGRADE_CODE & " failed at step """ & _
           Err.Source & """: " & Err.Description, vbCritical, "ERP Upgrade"
End Sub

Public Sub RollbackUpgrade_20260727_48()
    ' Removes the five placeholder forms added by this upgrade.
    ' Button-property changes inside frmQCLeftMenu are not reversed
    ' because the visual improvements are always desirable.
    On Error Resume Next
    Dim frms(4) As String
    frms(0) = "frmQCDashboard"
    frms(1) = "frmITPRegister"
    frms(2) = "frmNCRRegister"
    frms(3) = "frmTestReports"
    frms(4) = "frmQCReports"

    Dim i As Integer
    For i = 0 To 4
        If FormExists48(frms(i)) Then
            DoCmd.DeleteObject acForm, frms(i)
        End If
    Next i

    On Error GoTo 0
    MsgBox "Rollback for " & UPGRADE_CODE & " complete. " & _
           "Placeholder QC forms have been removed.", vbInformation, "ERP Upgrade"
End Sub

' ─── Step 1 – Placeholder forms ──────────────────────────────────────────────

Private Sub CreateQCPlaceholderForms48()
    ' Creates a clean, consistently-styled stub form for each QC feature
    ' that does not yet have a full register form.

    CreateStubForm48 "frmQCDashboard",  "QC Dashboard"
    CreateStubForm48 "frmITPRegister",  "Inspection & Test Plan Register"
    CreateStubForm48 "frmNCRRegister",  "Non-Conformance Report Register"
    CreateStubForm48 "frmTestReports",  "Test Reports"
    CreateStubForm48 "frmQCReports",    "QC Reports"
End Sub

Private Sub CreateStubForm48(ByVal formName As String, ByVal formTitle As String)
    ' Creates a placeholder form with a header label and a centred
    ' "Under Construction" notice.  Skips creation if the form already exists.

    If FormExists48(formName) Then Exit Sub

    Dim frm As Form
    Dim ctl As Control

    '-- Create form shell
    Set frm = CreateForm()
    frm.Caption    = formTitle
    frm.ScrollBars = 0                   ' Neither
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.BorderStyle = 1                  ' Thin
    frm.Width       = FORM_DEFAULT_WIDTH

    ' ---- Header section ----
    frm.Section(acHeader).Visible = True
    frm.Section(acHeader).Height  = 600

    Set ctl = CreateControl(frm.Name, acLabel, acHeader, , , 120, 80, 9260, 440)
    ctl.Name      = "lblFormTitle"
    ctl.Caption   = formTitle
    ctl.FontName  = STANDARD_FONT
    ctl.FontSize  = 14
    ctl.FontBold  = True
    ctl.ForeColor = RGB(31, 73, 125)     ' ERP dark blue
    ctl.BackStyle = 0                    ' Transparent
    ctl.TextAlign = 1                    ' Left

    ' ---- Detail section ----
    frm.Section(acDetail).Height = 3600

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 500, 1400, 8500, 520)
    ctl.Name      = "lblUnderConstruction"
    ctl.Caption   = "This module is currently under development. It will be available in a future upgrade."
    ctl.FontName  = STANDARD_FONT
    ctl.FontSize  = 11
    ctl.ForeColor = RGB(128, 128, 128)
    ctl.BackStyle = 0
    ctl.TextAlign = 2                    ' Center

    '-- Save and close
    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes

    '-- Rename from auto-generated name to the requested name
    DoCmd.Rename formName, acForm, frm.Name
End Sub

' ─── Step 2 – Style and wire QC navigation buttons ───────────────────────────

Private Sub StyleAndLinkQCButtons48()
    '-- Open the QC left-menu form in Design view so we can read/write properties
    DoCmd.OpenForm QC_MENU_FORM, acDesign

    Dim qcFrm As Form
    Set qcFrm = Forms(QC_MENU_FORM)

    '-- Use btnMIR (already complete) as the visual reference
    If Not ControlExists48(qcFrm, REF_BUTTON) Then
        DoCmd.Close acForm, QC_MENU_FORM, acSaveNo
        Err.Raise vbObjectError + 1001, "StyleAndLinkQCButtons48", _
                  "Reference button """ & REF_BUTTON & """ not found in " & QC_MENU_FORM
    End If

    Dim refBtn As CommandButton
    Set refBtn = qcFrm.Controls(REF_BUTTON)

    '-- Table of all 7 QC navigation buttons [control name, caption, target form]
    Dim btnCount As Integer
    btnCount = 7

    Dim names(6)   As String
    Dim caps(6)    As String
    Dim targets(6) As String

    names(0) = "btnQCDashboard" : caps(0) = "QC Dashboard"                : targets(0) = "frmQCDashboard"
    names(1) = "btnITPRegister" : caps(1) = "ITP Register"                : targets(1) = "frmITPRegister"
    names(2) = "btnMIR"         : caps(2) = "Material Inspection Request" : targets(2) = MIR_FORM
    names(3) = "btnWIR"         : caps(3) = "Work Inspection Request"     : targets(3) = "frmWIR"
    names(4) = "btnNCR"         : caps(4) = "Non-Conformance Report"      : targets(4) = "frmNCRRegister"
    names(5) = "btnTestReports" : caps(5) = "Test Reports"                : targets(5) = "frmTestReports"
    names(6) = "btnQCReports"   : caps(6) = "QC Reports"                  : targets(6) = "frmQCReports"

    Dim i As Integer

    '-- Apply style and caption to every QC navigation button
    For i = 0 To btnCount - 1
        If ControlExists48(qcFrm, names(i)) Then
            ApplyQCButtonStyle48 qcFrm, names(i), refBtn, caps(i), targets(i)
        End If
    Next i

    '-- Save and close the form
    DoCmd.Save acForm, QC_MENU_FORM
    DoCmd.Close acForm, QC_MENU_FORM, acSaveYes

    '-- Inject the click-event Sub bodies into the form's class module
    InjectQCClickHandlers48
End Sub

Private Sub ApplyQCButtonStyle48(ByVal frm As Form, _
                                  ByVal btnName As String, _
                                  ByVal refBtn As CommandButton, _
                                  ByVal caption As String, _
                                  ByVal targetForm As String)
    ' Copies visual properties from refBtn onto the named button,
    ' sets the caption, and marks the OnClick as an event procedure.

    Dim btn As CommandButton
    Set btn = frm.Controls(btnName)

    '-- Visual appearance (copied from reference / MIR-style)
    btn.BackColor       = refBtn.BackColor
    btn.ForeColor       = refBtn.ForeColor
    btn.FontName        = refBtn.FontName
    btn.FontSize        = refBtn.FontSize
    btn.FontBold        = refBtn.FontBold
    btn.FontItalic      = refBtn.FontItalic
    btn.FontUnderline   = refBtn.FontUnderline
    btn.BorderColor     = refBtn.BorderColor
    btn.BorderStyle     = refBtn.BorderStyle
    btn.BorderWidth     = refBtn.BorderWidth
    btn.BackStyle       = refBtn.BackStyle
    btn.SpecialEffect   = refBtn.SpecialEffect
    btn.Width           = refBtn.Width
    btn.Height          = refBtn.Height
    btn.TextAlign       = refBtn.TextAlign

    '-- Caption: set only when the button is unnamed / empty
    If Trim(btn.Caption) = "" _
    Or LCase(Trim(btn.Caption)) = LCase(btnName) Then
        btn.Caption = caption
    End If

    '-- Wire the click event (will be backed by code injected below)
    btn.OnClick = "[Event Procedure]"
End Sub

' ─── Step 2b – Inject click handlers via VBE ─────────────────────────────────

Private Sub InjectQCClickHandlers48()
    ' Writes the missing click-handler Subs into Form_frmQCLeftMenu's
    ' class module.  Requires "Trust access to the VBA project object model"
    ' (File > Options > Trust Center > Macro Settings).

    Const COMPONENT_NAME As String = "Form_frmQCLeftMenu"

    On Error GoTo VBEAccessDenied

    '-- Locate the VBProject for the current database rather than assuming index 1,
    '   which may not be correct when add-ins or multiple databases are loaded.
    Dim vbProj   As Object   ' VBIDE.VBProject
    Dim vbComp   As Object   ' VBIDE.VBComponent
    Dim codemod  As Object   ' VBIDE.CodeModule
    Dim prj      As Object

    For Each prj In Application.VBE.VBProjects
        If LCase(prj.Filename) = LCase(CurrentProject.FullName) Then
            Set vbProj = prj
            Exit For
        End If
    Next prj

    If vbProj Is Nothing Then
        Err.Raise vbObjectError + 1002, "InjectQCClickHandlers48", _
                  "Could not locate the VBProject for the current database."
    End If

    Set vbComp  = vbProj.VBComponents(COMPONENT_NAME)
    Set codemod = vbComp.CodeModule

    '-- All 7 click-handler Subs to inject  [sub name, target form name]
    Dim subNames(6)   As String
    Dim formNames(6)  As String

    subNames(0) = "btnQCDashboard_Click" : formNames(0) = "frmQCDashboard"
    subNames(1) = "btnITPRegister_Click" : formNames(1) = "frmITPRegister"
    subNames(2) = "btnMIR_Click"         : formNames(2) = MIR_FORM
    subNames(3) = "btnWIR_Click"         : formNames(3) = "frmWIR"
    subNames(4) = "btnNCR_Click"         : formNames(4) = "frmNCRRegister"
    subNames(5) = "btnTestReports_Click" : formNames(5) = "frmTestReports"
    subNames(6) = "btnQCReports_Click"   : formNames(6) = "frmQCReports"

    Dim i As Integer
    For i = 0 To 6
        InjectOpenFormHandler48 codemod, subNames(i), formNames(i)
    Next i

    Exit Sub

VBEAccessDenied:
    ' If VBE access is restricted fall back to function-call event expressions.
    ' This approach does not require VBE trust but does leave [Event Procedure]
    ' stubs in place – the developer should replace them with the generated
    ' function calls shown in the message below, or enable VBE trust and re-run.
    MsgBox "VBE access is not enabled." & vbCrLf & vbCrLf & _
           "To complete the click-event wiring automatically, enable " & _
           """Trust access to the VBA project object model"" in " & _
           "File > Options > Trust Center > Macro Settings, then re-run " & _
           "this upgrade." & vbCrLf & vbCrLf & _
           "Alternatively, add the following Subs manually to " & _
           "Form_frmQCLeftMenu:" & vbCrLf & vbCrLf & _
           BuildManualHandlerList48(), _
           vbInformation, "ERP Upgrade – Manual Step Required"
End Sub

Private Sub InjectOpenFormHandler48(ByVal codemod As Object, _
                                     ByVal subName As String, _
                                     ByVal targetForm As String)
    ' Adds a click-handler Sub if it does not already exist in the module.
    ' CodeModule.Find returns Boolean (True if found); search the full module.

    Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    startLine = 1 : startCol = 1
    endLine = codemod.CountOfLines : endCol = VBE_MAX_COLUMN

    If codemod.Find(subName, startLine, startCol, endLine, endCol) Then
        Exit Sub   ' Already present – do not duplicate
    End If

    Dim code As String
    code = BuildOpenFormHandlerCode48(subName, targetForm)

    codemod.InsertLines codemod.CountOfLines + 1, code
End Sub

Private Function BuildOpenFormHandlerCode48(ByVal subName As String, _
                                              ByVal targetForm As String) As String
    ' Returns the VBA source for a single navigation click-handler Sub.
    ' Used by both InjectOpenFormHandler48 (live injection) and
    ' BuildManualHandlerList48 (manual fallback message), so both paths
    ' always produce identical handler logic.
    Dim nl As String
    nl = vbCrLf
    BuildOpenFormHandlerCode48 = _
        "Private Sub " & subName & "()" & nl & _
        "    If (SysCmd(acSysCmdGetObjectState, acForm, """ & targetForm & """) And acObjStateOpen) <> 0 Then" & nl & _
        "        DoCmd.SelectObject acForm, """ & targetForm & """" & nl & _
        "    Else" & nl & _
        "        DoCmd.OpenForm """ & targetForm & """" & nl & _
        "    End If" & nl & _
        "End Sub"
End Function

Private Function BuildManualHandlerList48() As String
    ' Returns a readable list of handler stubs for the manual-step message.
    ' Each stub is generated by BuildOpenFormHandlerCode48, which is also used
    ' by InjectOpenFormHandler48, guaranteeing the two paths stay in sync.
    Dim nl As String
    nl = vbCrLf

    Dim subs(6)  As String
    Dim forms(6) As String
    subs(0) = "btnQCDashboard_Click" : forms(0) = "frmQCDashboard"
    subs(1) = "btnITPRegister_Click" : forms(1) = "frmITPRegister"
    subs(2) = "btnMIR_Click"         : forms(2) = MIR_FORM
    subs(3) = "btnWIR_Click"         : forms(3) = "frmWIR"
    subs(4) = "btnNCR_Click"         : forms(4) = "frmNCRRegister"
    subs(5) = "btnTestReports_Click" : forms(5) = "frmTestReports"
    subs(6) = "btnQCReports_Click"   : forms(6) = "frmQCReports"

    Dim i As Integer
    For i = 0 To 6
        If i = 0 Then
            BuildManualHandlerList48 = BuildOpenFormHandlerCode48(subs(i), forms(i))
        Else
            BuildManualHandlerList48 = BuildManualHandlerList48 & nl & nl & _
                                       BuildOpenFormHandlerCode48(subs(i), forms(i))
        End If
    Next i
End Function

' ─── Step 3 – Record upgrade ─────────────────────────────────────────────────

Private Sub RecordUpgrade48()
    Dim sql As String
    sql = "INSERT INTO tblERPUpgradeHistory " & _
          "(UpgradeVersion, AppliedOn, AppliedBy, FrontEndName, Notes) " & _
          "VALUES (" & _
          "'" & UPGRADE_CODE & "', " & _
          "Now(), " & _
          "'" & Replace(CurrentUser(), "'", "''") & "', " & _
          "'" & Replace(CurrentProject.Name, "'", "''") & "', " & _
          "'" & Replace(UPGRADE_NAME, "'", "''") & "')"
    CurrentDb.Execute sql, dbFailOnError
End Sub

' ─── Helpers ─────────────────────────────────────────────────────────────────

Private Function UpgradeAlreadyApplied48(ByVal upgradeCode As String) As Boolean
    ' Uses Replace to guard against single-quote injection in the version string.
    Dim safecode As String
    safecode = Replace(upgradeCode, "'", "''")
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT COUNT(*) AS n FROM tblERPUpgradeHistory " & _
        "WHERE UpgradeVersion='" & safecode & "'", _
        dbOpenSnapshot)
    UpgradeAlreadyApplied48 = (rs!n > 0)
    rs.Close
End Function

Private Function FormExists48(ByVal formName As String) As Boolean
    On Error Resume Next
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If LCase(obj.Name) = LCase(formName) Then
            FormExists48 = True
            Exit Function
        End If
    Next obj
    FormExists48 = False
End Function

Private Function ControlExists48(ByVal frm As Form, _
                                  ByVal controlName As String) As Boolean
    On Error Resume Next
    Dim ctl As Control
    Set ctl = frm.Controls(controlName)
    ControlExists48 = (Err.Number = 0)
    On Error GoTo 0
End Function
