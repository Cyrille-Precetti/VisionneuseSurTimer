Imports System.IO

Public Class Form1
    Private _booTacheEnCours As Boolean = False
    Private unTimer As System.Timers.Timer
    Private DiapoNumber As Integer = 1
    'Liste des types d'image
    Private ExtensionsFichiersImages() As String = {".gif", ".GIF", ".jpg", ".JPG", ".bmp", ".png", ".tiff", ".TIFF"}
    Private ListeDiapos As List(Of Bitmap) 'ImageList

    Private Sub ChargementImagesDepuisDossier(ByRef CheminDossier As String)
        'Création de la liste d'image ImageList
        ListeDiapos = New List(Of Bitmap) ' New ImageList
        'Récupération des fichiers depuis le dossier en question
        Dim FichiersDansDir() As String = IO.Directory.GetFiles(CheminDossier)
        Try
            'Pour chaque fichier on vérifie la validité de l'extension et on le charge dans l'ImageList
            For Each FichierImage As String In FichiersDansDir
                Dim ExtensionFichier As String = Path.GetExtension(FichierImage)
                If ExtensionsFichiersImages.Contains(ExtensionFichier) = True Then
                    Dim diapo As System.Drawing.Image = bitmap.FromFile(FichierImage)
                    ListeDiapos.Add(diapo) 'ListeDiapos.Images.Add(diapo)
                End If
            Next
        Catch ex As Exception
            Debug.Print("Erreur " & ex.Message & " type: " & ex.ToString)
        End Try


    End Sub
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
            'Select Case DiapoNumber
            '    Case 1
            '        PictureBox1.Image = My.Resources.Pic1
            '    Case 2
            '        PictureBox1.Image = My.Resources.Pic2
            '    Case 3
            '        PictureBox1.Image = My.Resources.Pic3
            'End Select

            If DiapoNumber < ListeDiapos.Count - 1 Then
                PictureBox1.Image = ListeDiapos(DiapoNumber) 'ListeDiapos.Images(DiapoNumber)
                DiapoNumber += 1
            Else
                PictureBox1.Image = ListeDiapos(0) 'ListeDiapos.Images(1)
                DiapoNumber = 1
            End If

        Catch ex As Exception
            Debug.Print("Exception:" & ex.Message)
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'C:\Users\Cyrille\Documents\Mes images\Birthday2014
        Call ChargementImagesDepuisDossier("C:\Users\Cyrille\Documents\Mes images\Birthday2014")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MessageBox.Show("Hello" & TextBox1.Text.ToString)
    End Sub
End Class
