Public Class Form1
    Private _booTacheEnCours As Boolean = False
    Private unTimer As System.Timers.Timer
    Private DiapoNumber As Integer = 1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Lancer le timer pour lancer la visionneuse
        If _booTacheEnCours = False Then
            unTimer = New System.Timers.Timer()
            unTimer.Interval = 500

            ' Déclaration de l'événement déclenché à chaque fin du timer
            AddHandler unTimer.Elapsed, AddressOf OnTimedEvent

            unTimer.AutoReset = True

            ' Démarrage
            unTimer.Enabled = True
            Button1.Text = "Arrêter visionneuse"
            _booTacheEnCours = True
        Else
            unTimer.Stop()
            Button1.Text = "Démarrer visionneuse"
            _booTacheEnCours = False

        End If
    End Sub
    Private Sub OnTimedEvent(source As Object, e As System.Timers.ElapsedEventArgs)
        Try
            'On affiche l'heure de l'événement
            Select Case DiapoNumber
                Case 1
                    PictureBox1.Image = My.Resources.Pic1
                Case 2
                    PictureBox1.Image = My.Resources.Pic2
                Case 3
                    PictureBox1.Image = My.Resources.Pic3
            End Select
            If DiapoNumber < 3 Then
                DiapoNumber += 1
            Else
                DiapoNumber = 1
            End If

        Catch ex As Exception
            Debug.Print("Exception:" & ex.Message)
        End Try


    End Sub
End Class
