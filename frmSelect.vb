Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class frmSelect

    ' Fills the combobox from the database when a radio button is checked
    Private Sub optSSOP_CheckedChanged(sender As Object, e As EventArgs) Handles optSSOP.CheckedChanged, optSCOP.CheckedChanged, chkPayload.CheckedChanged, chkAOCS.CheckedChanged, chkEPS.CheckedChanged, chkSCS.CheckedChanged, chkSYS.CheckedChanged, chkTCR.CheckedChanged

        If optSSOP.Checked = True Then
            procFilter("%SSOP%")
        ElseIf optSCOP.Checked = True Then
            procFilter("%SCOP%")
        End If

    End Sub

    ' Initialise the form
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        Dim chk As CheckBox
        optSCOP.Checked = False
        optSSOP.Checked = True
        cboSelect.Text = ""
        cboSelect.Items.Clear()

        For Each chk In Me.Controls.OfType(Of CheckBox)

            If chk.Checked = True Then
                chk.Checked = False
            End If
        Next

    End Sub

    ' Creates the form for the procedure selected
    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click

        Dim strARESID As String = ""

        If cboSelect.Text <> "" Then
            SQLa = "SELECT * FROM ARES_Procedure_List"

            CallAccess()
            RSa.MoveFirst()
            Do Until RSa.EOF
                If cboSelect.Text = RSa.Fields("ARES_ID").Value Then
                    strARESID = RSa.Fields("ARES_ID").Value
                End If
                RSa.MoveNext()
            Loop
            CloseAccess()


            If strARESID <> "" Then
                Dim frmProcedure As New frmProcedure()
                frmProcedure.Text = cboSelect.Text
                If frmProcedure.Visible = False Then
                    frmProcedure.MdiParent = frmMDI
                    frmProcedure.Show()
                    If frmProcedure.Visible = True Then
                    End If
                End If
                Me.Close()
            Else
                MsgBox("Procedure " & strARESID & " doesn't exist in the database.", MsgBoxStyle.OkOnly, "Eshailsat Information")
            End If
        End If

    End Sub

    ' Checks which checkbox/radio button is chosen and builds a query based on it
    Private Sub procFilter(ByVal type As String)

        Dim strChk As String = ""
        Dim chk As CheckBox
        Dim strSubsystem As String = ""
        Dim listSubsystem As List(Of [String]) = New List(Of String)

        SQLa = "SELECT * FROM ARES_Procedure_List WHERE ARES_ID LIKE '" & type & "'"

        cboSelect.Text = ""

        ' Loops through the checkbox controls
        For Each chk In Me.Controls.OfType(Of CheckBox)

            If chk.Checked = True Then
                listSubsystem.Add("Subsystem = '" & Microsoft.VisualBasic.Mid(chk.Text, 1) & "'")
            End If
        Next

        ' Only executes when at least one checkbox is selected
        If listSubsystem.Count > 0 Then
            strSubsystem = [String].Join(" OR ", listSubsystem.ToArray())
            SQLa = SQLa & "AND (" & strSubsystem & ")"
        End If

        CallAccess()
        cboSelect.Items.Clear()

        Do Until RSa.EOF()
            cboSelect.Items.Add(RSa.Fields("ARES_ID").Value)
            RSa.MoveNext()
        Loop
        CloseAccess()

    End Sub


End Class