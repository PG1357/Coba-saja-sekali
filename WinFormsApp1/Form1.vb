Public Class Form1
    Dim hasil As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.LoadMSComm1.PortName = "COM5"
        MSComm1.BaudRate = 9600
        MSComm1.Parity = IO.Ports.Parity.None
        MSComm1.DataBits = 8
        MSComm1.StopBits = IO.Ports.StopBits.One
        MSComm1.Open()
        Timer1.Enabled = True
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MSComm1.Close()
        Timer1.Enabled = False
        End
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click MSComm1.WriteLine(Trim(TextBox1.Text))
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As System.Object, ByVal e As
System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
If e.KeyChar = Chr(13) Then
            MSComm1.WriteLine(Trim(TextBox1.Text))
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Timer1.Tick
hasil = MSComm1.ReadExisting()
        If Len(hasil) > 0 Then
            Me.TextBox2.Text = hasil
            ListBox1.Items.Add(hasil)
            Me.TextBox1.Text = ""
            Me.TextBox2.Refresh()
            Me.ListBox1.Refresh()
        End If
    End Sub
End Class
