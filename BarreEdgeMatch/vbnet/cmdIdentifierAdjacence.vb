Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display

'**
'Nom de la composante : cmdIdentifierAdjacence.vb
'
'''<summary>
''' Commande qui permet d'identifier les points d'adjacence et les erreurs aux limites communes de découpage.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 21 juillet 2011
'''</remarks>
''' 
Public Class cmdIdentifierAdjacence
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        Try
            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                gpApp = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Rendre active la commande
                    Enabled = True
                    'Définir le document
                    gpMxDoc = CType(gpApp.Document, IMxDocument)

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing      'Interface qui permet de changer le curseur de la souris.
        Dim pTrackCancel As ITrackCancel = Nothing      'Interface qui permet d'annuler la sélection avec la touche ESC.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Vider la liste des points d'adjacence et d'erreur
            m_MenuEdgeMatch.trePoints.Nodes.Clear()
            m_MenuEdgeMatch.treErreurs.Nodes.Clear()

            'Initialiser la map et la liste des classes visible
            'Call m_MenuEdgeMatch.InitialiserClasses()

            'Permettre d'annuler la sélection avec la touche ESC
            pTrackCancel = New CancelTracker
            pTrackCancel.CancelOnKeyPress = True
            pTrackCancel.CancelOnClick = False
            pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar

            'Transformer les tolérances en géographique au besoin
            Call TransformerTolerances(m_LimiteDecoupage.Envelope)

            'Indentifier tous les points d'adjacence
            Call IdentifierAdjacenceDecoupage(pTrackCancel)

            'Remplir le formulaires contenant les points d'adjacence
            pTrackCancel.Progressor.Message = "Remplir le formulaire des points d'adjacence"
            Call m_MenuEdgeMatch.RemplirFormulairePointAdjacence()

            'Remplir le formulaires contenant les erreurs de précision
            pTrackCancel.Progressor.Message = "Remplir le formulaire des erreurs de précision"
            Call m_MenuEdgeMatch.RemplirFormulaireErreurPrecision()

            'Remplir le formulaires contenant les erreurs d'adjacence
            pTrackCancel.Progressor.Message = "Remplir le formulaire des erreurs d'adjacence"
            Call m_MenuEdgeMatch.RemplirFormulaireErreurAdjacence()

            'Remplir le formulaires contenant les erreurs d'attribut
            pTrackCancel.Progressor.Message = "Remplir le formulaire des erreurs d'attribut"
            Call m_MenuEdgeMatch.RemplirFormulaireErreurAttribut()

            'Initialiser le nombre de points d.Adjacence
            m_MenuEdgeMatch.lblNbPointAdjacence.Text = "Nombre de points d'Adjacence: " & m_ListePointAdjacence.GeometryCount.ToString
            'Initialiser le nombre d'erreurs
            m_MenuEdgeMatch.lblNbErreur.Text = "Nombre total d'erreurs: " & m_ErreurFeature.Count.ToString

            'Montrer toutes l'information pour tous points
            If m_MenuEdgeMatch.chkOuvrirPoints.Checked And m_MenuEdgeMatch.trePoints.Nodes.Count > 0 Then m_MenuEdgeMatch.trePoints.ExpandAll()

            'Montrer toutes l'information pour toutes les erreurs
            If m_MenuEdgeMatch.chkOuvrirErreurs.Checked And m_MenuEdgeMatch.treErreurs.Nodes.Count > 0 Then m_MenuEdgeMatch.treErreurs.ExpandAll()

            'Afficher la page des erreurs
            m_MenuEdgeMatch.tabEdgeMatch.SelectedIndex = 3

            'Dessiner la géométrie et le texte des erreurs de précision, d'adjacence et d'attribut
            pTrackCancel.Progressor.Message = "Dessiner les erreurs"
            Call DessinerErreurs()
            pTrackCancel.Progressor.Message = ""

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            pTrackCancel = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si la limite commune est présente
        If m_LimiteDecoupage Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si la limite commune est vide
            If m_LimiteDecoupage.IsEmpty Then
                'Désactiver la commande
                Enabled = False
            Else
                'Activer la commande
                Enabled = True
            End If
        End If
    End Sub
End Class
