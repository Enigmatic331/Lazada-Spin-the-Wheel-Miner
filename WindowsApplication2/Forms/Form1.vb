Public Class Form1
   Public Sub SetMessage(ByVal strMessage As String)
      Label1.Text = strMessage
      Label1.Refresh()
   End Sub

   Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
      Me.Text = System.Diagnostics.Process.GetCurrentProcess().ProcessName
   End Sub
End Class