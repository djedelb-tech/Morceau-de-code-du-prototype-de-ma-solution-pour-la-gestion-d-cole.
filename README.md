Imports System.Data
Imports System.DirectoryServices.ActiveDirectory
Imports System.Web
Imports Microsoft.Data.SqlClient
Imports Microsoft.SqlServer

' --- Appel au système lié à l'impression et IO pour télécharger les fichiers ---

Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms.Design


' --- Importats system.Data fait appel au système de données et le Microsoft.Data.SqlClient au sql server donc requêtes et tout ---
Public Class Academie

    ' --- Chaîne de connexion de l'interface à la base ---

    Dim conSql As New SqlConnection("Server=DESKTOP-LVSR9EN\BDSERVER;DataBase=BQAS;Trusted_Connection=True;TrustServerCertificate=True;")

    ' --- code début impression ---

    Dim AperçuInscrits As New PrintPreviewDialog
    Dim ListeInscrits As New List(Of Eleve)

    ' --- Variables pour la pagination ---

    Dim currentIndex As Integer = 0
    Dim NumeroPage As Integer = 1

    Private Sub ImprimerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImprimerToolStripMenuItem.Click
        ChargerDonnees()
        PrintDocument1.DefaultPageSettings.Landscape = True
        AperçuInscrits.Document = PrintDocument1
        AperçuInscrits.ShowDialog()

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim y As Integer = 50
        Dim TitreImpression As New Font("Segoe UI", 14, FontStyle.Bold)
        Dim CorpsImpression As New Font("Segoe UI", 13, FontStyle.Regular)


        ' --- Colonnes ---

        Dim xNom As Integer = 50
        Dim xPostnom As Integer = 150
        Dim xPrenom As Integer = 250
        Dim xSexe As Integer = 350
        Dim xDateNaissance As Integer = 450
        Dim xLieuNaissance As Integer = 550
        Dim xNiveau As Integer = 650
        Dim xClasse As Integer = 750
        Dim xSection As Integer = 850
        Dim xOptionE As Integer = 950
        Dim xAdresse As Integer = 1050
        Dim xTelephone As Integer = 1150
        Dim xDateInscription As Integer = 1250

        Dim LargeurTableau As Integer = 650

        ' --- TitreImpression ---

        e.Graphics.DrawString("LISTE DES INSCRITS", TitreImpression, Brushes.Black, 50, y)
        y += 40

        ' --- En-têtes ---

        e.Graphics.DrawString("Nom", CorpsImpression, Brushes.Black, xNom, y)
        e.Graphics.DrawString("Postnom", CorpsImpression, Brushes.Black, xPostnom, y)
        e.Graphics.DrawString("Prénom", CorpsImpression, Brushes.Black, xPrenom, y)
        e.Graphics.DrawString("Sexe", CorpsImpression, Brushes.Black, xSexe, y)
        e.Graphics.DrawString("Naissance", CorpsImpression, Brushes.Black, xDateNaissance, y)
        e.Graphics.DrawString("Lieu", CorpsImpression, Brushes.Black, xLieuNaissance, y)
        e.Graphics.DrawString("Niveau", CorpsImpression, Brushes.Black, xNiveau, y)
        e.Graphics.DrawString("Classe", CorpsImpression, Brushes.Black, xClasse, y)
        e.Graphics.DrawString("Section", CorpsImpression, Brushes.Black, xSection, y)
        e.Graphics.DrawString("Option", CorpsImpression, Brushes.Black, xOptionE, y)
        e.Graphics.DrawString("Adresse", CorpsImpression, Brushes.Black, xAdresse, y)
        e.Graphics.DrawString("Téléphone", CorpsImpression, Brushes.Black, xTelephone, y)
        e.Graphics.DrawString("Inscription", CorpsImpression, Brushes.Black, xDateInscription, y)

        y += 40

        ' Ligne sous les en-têtes
        y += 20
        e.Graphics.DrawLine(Pens.Black, xNom, y, largeurTableau, y)

        ' --- Corps ---

        ' For Each eleve In ListeInscrits
        'e.Graphics.DrawString(eleve, CorpsImpression, Brushes.Black, 50, y)
        'y += 25

        ' Corps
        Try
            While currentIndex < ListeInscrits.Count
                Dim Eleve = ListeInscrits(currentIndex)
                y += 20
                e.Graphics.DrawString(Eleve.Nom, CorpsImpression, Brushes.Black, xNom, y)
                e.Graphics.DrawString(Eleve.Postnom, CorpsImpression, Brushes.Black, xPostnom, y)
                e.Graphics.DrawString(Eleve.Prenom, CorpsImpression, Brushes.Black, xPrenom, y)
                e.Graphics.DrawString(Eleve.Sexe, CorpsImpression, Brushes.Black, xSexe, y)
                e.Graphics.DrawString(Eleve.DateNaissance, CorpsImpression, Brushes.Black, xDateNaissance, y)
                e.Graphics.DrawString(Eleve.LieuNaissance, CorpsImpression, Brushes.Black, xLieuNaissance, y)
                e.Graphics.DrawString(Eleve.Niveau, CorpsImpression, Brushes.Black, xNiveau, y)
                e.Graphics.DrawString(Eleve.Classe, CorpsImpression, Brushes.Black, xClasse, y)
                e.Graphics.DrawString(Eleve.Section, CorpsImpression, Brushes.Black, xSection, y)
                e.Graphics.DrawString(Eleve.OptionE, CorpsImpression, Brushes.Black, xOptionE, y)
                e.Graphics.DrawString(Eleve.Adresse, CorpsImpression, Brushes.Black, xAdresse, y)
                e.Graphics.DrawString(Eleve.Telephone, CorpsImpression, Brushes.Black, xTelephone, y)
                e.Graphics.DrawString(Eleve.DateInscription, CorpsImpression, Brushes.Black, xDateInscription, y)

                currentIndex += 1

                ' Saut de page
                If y > e.PageBounds.Height - 100 Then
                    e.HasMorePages = True
                    NumeroPage += 1
                    Return
                End If
            End While
            ' Next

            ' --- Bordures verticales ---

            e.Graphics.DrawLine(Pens.Black, xNom, y, LargeurTableau, y)       ' Bordure gauche
            e.Graphics.DrawLine(Pens.Black, LargeurTableau, 70, LargeurTableau, y + 15) ' Bordure droite

            ' --- Pied de page ---

            Dim PolicePiedPgae As New Font("Segoe UI", 9, FontStyle.Italic)
            Dim TextePiedPage As String = "Page " & NumeroPage & " - Imprimé le " & DateTime.Now.ToString("dd/MM/yyyy HH:mm")
            e.Graphics.DrawString(TextePiedPage, PolicePiedPgae, Brushes.Gray, 50, e.PageBounds.Height - 50)
            e.HasMorePages = False

            ' --- Gestion du saut de page ---

            If currentIndex < ListeInscrits.Count Then
                e.HasMorePages = True
                NumeroPage += 1
            Else
                e.HasMorePages = False
                currentIndex = 0
                NumeroPage = 1
            End If
        Catch ex As Exception
            MessageBox.Show("Une erreur est survenue", ex.Message)
        End Try

    End Sub

    ' -- Code pour charger de données depuis la Base de données ---

    Sub ChargerDonnees()
        ListeInscrits.Clear()
        Dim cmd As New SqlCommand("SELECT Nom, Postnom, Prenom, Sexe, DateNaissance, LieuNaissance, Adresse, Niveau, Classe, Section, OptionE, Telephone, DateInscription from INSCRIPTION", conSql)
        conSql.Open()

        ' --- code pour lire les données ---

        Dim Lecture As SqlDataReader = cmd.ExecuteReader
        Try
            While Lecture.Read()
                Dim eleve As New Eleve With {
                    .Nom = Lecture("Nom").ToString(),
                    .Postnom = Lecture("Postnom").ToString(),
                    .Prenom = Lecture("Prenom").ToString(),
                    .Sexe = Lecture("Sexe").ToString(),
                    .DateNaissance = Lecture("DateNaissance").ToString(),
                    .LieuNaissance = Lecture("LieuNaissance").ToString(),
                    .Adresse = Lecture("Adresse").ToString(),
                    .Niveau = Lecture("Niveau").ToString(),
                    .Classe = Lecture("Classe").ToString(),
                    .Section = Lecture("Section").ToString(),
                    .OptionE = Lecture("OptionE").ToString(),
                    .Telephone = Lecture("Telephone").ToString(),
                    .DateInscription = Lecture("DateInscription").ToString()
                }
                ListeInscrits.Add(eleve)
            End While
        Catch ex As Exception
            MessageBox.Show("Une erreur est survenue", ex.Message)
        Finally
            conSql.Close()
        End Try

    End Sub
    ' Classe Eleve
    Public Class Eleve
        Public Property Nom As String
        Public Property Postnom As String
        Public Property Prenom As String
        Public Property Sexe As String
        Public Property DateNaissance As String
        Public Property LieuNaissance As String
        Public Property Adresse As String
        Public Property Niveau As String
        Public Property Classe As String
        Public Property Section As String
        Public Property OptionE As String
        Public Property Telephone As String
        Public Property DateInscription As String
    End Class


    Private Sub Academie_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' --- Code crée juste pour teste la connexion avec la base de données. ---
        'Try
        'conSql.Open()
        'MessageBox.Show("Connexion réussie ")
        'conSql.Close()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        ' --- End Try ---

        ComboBox1.Focus()
        ToolStripButton2.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' --- Code pour annuler les actions en cours ---

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.Text = ""
        ComboBox1.Focus()
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
    End Sub

    Private Sub FermerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FermerToolStripMenuItem.Click

        ' --- Fermeture ---

        DialogResult = MessageBox.Show("Voulez-vous vraiment quitter ?", "Fermeture", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
        If DialogResult = vbYes Then
            End
        End If
    End Sub

    Private Sub FermerLaFenêtreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FermerLaFenêtreToolStripMenuItem.Click
        DialogResult = MessageBox.Show("Voulez-vous vraiment quitter ?", "Fermeture", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
        If DialogResult = vbYes Then
            End
        End If
    End Sub

    Private Sub FermerToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FermerToolStripMenuItem1.Click
        DialogResult = MessageBox.Show("Voulez-vous vraiment quitter ?", "Fermeture", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
        If DialogResult = vbYes Then
            End
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        ' --- Changement d'une interface à une autre ---

        Me.Hide()
        Acceuil.Show()
    End Sub

    Private Sub ComboBox1_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox1.DrawItem

        ' --- Evenement DrawItem permet d'afficher les continus combobox lorsque les propriétés d'affichage de celui ont été personnalisé ---

        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox1.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub

    Private Sub ComboBox2_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox2.DrawItem
        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox2.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub

    Private Sub ComboBox3_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox3.DrawItem
        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox3.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub


    Private Sub ComboBox4_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox4.DrawItem
        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox4.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub

    Private Sub ComboBox5_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox5.DrawItem
        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox5.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub

    Private Sub ComboBox6_drawitem(sender As Object, e As DrawItemEventArgs) Handles ComboBox6.DrawItem
        If e.Index < 0 Then Exit Sub
        e.DrawBackground()
        Dim Brush As Brush = Brushes.LightBlue
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Brush = Brushes.White
        End If
        e.Graphics.DrawString(ComboBox6.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        e.DrawFocusRectangle()
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Maternelle" Or ComboBox2.Text = "Primaire" Or ComboBox2.Text = "Secondaire" Then
            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
        Else
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True
        End If
    End Sub

    ' --- Choix de classe selon le niveau ---
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox2.Text = "Maternelle" And ComboBox3.Text = "4ème" Or ComboBox3.Text = "5ème" Or ComboBox3.Text = "6ème" Then
            MessageBox.Show("Rassurez-vous d'avoir choisi la bonne classe pour le cycle de Maternelle", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ComboBox3.Text = ""
            ComboBox3.Focus()
        End If
    End Sub

    ' --- Code d'enregistrement élève liant l'interface à la base ---

    Private connectionString As String = "Server=DESKTOP-LVSR9EN\BDSERVER;DataBase=BQAS;Trusted_Connection=True;TrustServerCertificate=True;"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim action As String = ComboBox1.Text.Trim()

        Select Case action
            Case "Inscription"
                ' --- code d'insertion ---
                Dim Nom As String = TextBox1.Text
                Dim Postnom As String = TextBox2.Text
                Dim Prenom As String = TextBox3.Text
                Dim Sexe As String = ComboBox6.Text
                Dim DateNaissance As Date = DateTimePicker1.Text
                Dim LieuNaissance As String = TextBox4.Text
                Dim Adresse As String = TextBox5.Text
                Dim Telephone As String = TextBox6.Text
                Dim Niveau As String = ComboBox2.Text
                Dim Classe As String = ComboBox3.Text
                Dim Section As String = ComboBox4.Text
                Dim OptionE As String = ComboBox5.Text
                Dim DateInscription As Date = DateTimePicker2.Text

                Dim Query As String = "INSERT INTO INSCRIPTION (Nom, Postnom, Prenom, Sexe, DateNaissance, LieuNaissance, Adresse, Telephone, Niveau, Classe, Section, OptionE, DateInscription) values (@Nom, @Postnom, @Prenom, @Sexe, @DateNaissance, @LieuNaissance, @Adresse, @Telephone, @Niveau, @Classe, @Section, @OptionE, @DateInscription)"

                Using cmd As New SqlCommand(Query, conSql)
                    cmd.Parameters.AddWithValue("@Nom", TextBox1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Postnom", TextBox2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Prenom", TextBox3.Text.Trim())
                    cmd.Parameters.AddWithValue("@Sexe", ComboBox6.Text.Trim())
                    cmd.Parameters.AddWithValue("@DateNaissance", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@LieuNaissance", TextBox4.Text.Trim())
                    cmd.Parameters.AddWithValue("@Adresse", TextBox5.Text.Trim())
                    cmd.Parameters.AddWithValue("@Telephone", TextBox6.Text.Trim)
                    cmd.Parameters.AddWithValue("@DateInscription", DateTimePicker2.Text)
                    cmd.Parameters.AddWithValue("@Niveau", ComboBox2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Classe", ComboBox3.Text.Trim())
                    cmd.Parameters.AddWithValue("@Section", ComboBox4.Text.Trim())
                    cmd.Parameters.AddWithValue("@OptionE", ComboBox5.Text.Trim())

                    Try
                        conSql.Open()
                        cmd.ExecuteNonQuery()
                        MessageBox.Show("Elève enregistré avec succès", "Inscription", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        ComboBox1.Text = ""
                        ComboBox1.Focus()
                        ComboBox2.Text = ""
                        ComboBox3.Text = ""
                        ComboBox4.Text = ""
                        ComboBox5.Text = ""
                        ComboBox6.Text = ""
                        ComboBox4.Enabled = True
                        ComboBox5.Enabled = True
                    Catch ex As Exception
                        MessageBox.Show("Erreur", ex.Message)
                    End Try
                End Using
                conSql.Close()

            Case "Modification"

                ' --- Exemple de mise à jour (à adapter selon ta clé primaire) ---

                Dim Query As String = "UPDATE INSCRIPTION SET Nom=@Nom, Postnom=@Postnom, Prenom=@Prenom, Sexe=@Sexe, DateNaissance=@DateNaissance, LieuNaissance=@LieuNaissance, Adresse=@Adresse, Telephone=@Telephone, Niveau=@Niveau, Classe=@Classe, Section=@Section, OptionE=@OptionE, DateInscription=@DateInscription WHERE Id=@Id"

                ' --- Ajoute ici tes paramètres et ta logique ---
                Using cmd As New SqlCommand(Query, conSql)
                    cmd.Parameters.AddWithValue("@Nom", TextBox1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Postnom", TextBox2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Prenom", TextBox3.Text.Trim())
                    cmd.Parameters.AddWithValue("@Sexe", ComboBox6.Text.Trim())
                    cmd.Parameters.AddWithValue("@DateNaissance", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@LieuNaissance", TextBox4.Text.Trim())
                    cmd.Parameters.AddWithValue("@Adresse", TextBox5.Text.Trim())
                    cmd.Parameters.AddWithValue("@Telephone", TextBox6.Text.Trim)
                    cmd.Parameters.AddWithValue("@DateInscription", DateTimePicker2.Text)
                    cmd.Parameters.AddWithValue("@Niveau", ComboBox2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Classe", ComboBox3.Text.Trim())
                    cmd.Parameters.AddWithValue("@Section", ComboBox4.Text.Trim())
                    cmd.Parameters.AddWithValue("@OptionE", ComboBox5.Text.Trim())

                    Try
                        conSql.Open()
                        cmd.ExecuteNonQuery()
                        MessageBox.Show("Elève modifié avec succès", "Modification", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        ComboBox1.Text = ""
                        ComboBox1.Focus()
                        ComboBox2.Text = ""
                        ComboBox3.Text = ""
                        ComboBox4.Text = ""
                        ComboBox5.Text = ""
                        ComboBox6.Text = ""
                        ComboBox4.Enabled = True
                        ComboBox5.Enabled = True
                    Catch ex As Exception
                        MessageBox.Show("Erreur", ex.Message)
                    End Try
                End Using
                conSql.Close()

            Case "Réinscription"
            ' --- Exemple : soit un nouvel INSERT, soit un UPDATE selon ta logique métier ---

            Case "Suppression"
                Dim Query As String = "DELETE FROM INSCRIPTION WHERE Id=@Id"
                ' --- Ajoute ici tes paramètres et ta logique ---

            Case Else
                MessageBox.Show("Veuillez choisir une action valide.", "Action", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Select
    End Sub
End Class
