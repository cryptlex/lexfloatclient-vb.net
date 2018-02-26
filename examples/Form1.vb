Imports FloatSample.Cryptlex

Public Class Form1
    Dim floatClient As LexFloatClient
    Public Sub New()
        InitializeComponent()
        Dim status As Integer
        status = LexFloatClient.SetProductFile("Product.dat")
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error setting product file: " + status.ToString()
            Return
        End If
    End Sub

    Private Sub LicenceRefreshCallback(status As UInteger)
        Select Case status
            Case LexFloatClient.LF_E_LICENSE_EXPIRED
                Me.statusLabel.Text = "The lease expired before it could be renewed."
                Exit Select
            Case LexFloatClient.LF_E_LICENSE_EXPIRED_INET
                Me.statusLabel.Text = "The lease expired due to network connection failure."
                Exit Select
            Case LexFloatClient.LF_E_SERVER_TIME
                Me.statusLabel.Text = "The lease expired because Server System time was modified."
                Exit Select
            Case LexFloatClient.LF_E_TIME
                Me.statusLabel.Text = "The lease expired because Client System time was modified."
                Exit Select
            Case Else
                Me.statusLabel.Text = "The lease expired due to some other reason."
                Exit Select
        End Select
    End Sub

    Private Sub leaseBtn_Click(sender As Object, e As EventArgs) Handles leaseBtn.Click
        If floatClient IsNot Nothing AndAlso floatClient.HasLicense() = LexFloatClient.LF_OK Then
            Return
        End If
        Dim status As Integer
        floatClient = New LexFloatClient()
        status = floatClient.SetVersionGUID("59A44CE9-5415-8CF3-BD54-EA73A64E9A1B")
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error setting version GUID: " + status.ToString()
            Return
        End If
        status = floatClient.SetFloatServer("localhost", 8090)
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error Setting Host Address: " + status.ToString()
            Return
        End If
        status = floatClient.SetLicenseCallback(AddressOf LicenceRefreshCallback)
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error Setting Callback Function: " + status.ToString()
            Return
        End If
        status = floatClient.RequestLicense()
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error Requesting License: " + status.ToString()
            Return
        End If
        Me.statusLabel.Text = "License leased successfully!"
    End Sub

    Private Sub dropBtn_Click(sender As Object, e As EventArgs) Handles dropBtn.Click
        If floatClient Is Nothing Then
            Return
        End If
        Dim status As Integer
        status = floatClient.DropLicense()
        If status <> LexFloatClient.LF_OK Then
            Me.statusLabel.Text = "Error Dropping License: " + status.ToString()
            Return
        End If
        Me.statusLabel.Text = "License dropped successfully!"
    End Sub
End Class
