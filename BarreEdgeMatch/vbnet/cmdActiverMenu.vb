Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms

'**
'Nom de la composante : cmdActiverMenu.vb
'
'''<summary>
''' Commande qui permet d'activer le menu du EdgeMatch afin de définir les classes et les paramètres d'affichage.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 4 juillet 2011
'''</remarks>
''' 
Public Class cmdActiverMenu
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    'Déclarer les variables de travail
    Private Shared _DockWindow As ESRI.ArcGIS.Framework.IDockableWindow
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        Try
            'Définir les variables de travail
            Dim windowID As UID = New UIDClass

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

                    'Créer un nouveau menu
                    windowID.Value = "MPO_BarreEdgeMatch_dckMenuEdgeMatch"
                    _DockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

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
        Try
            'Sortir si le menu n'est pas créé
            If _DockWindow Is Nothing Then Return

            'Activer ou désactiver le menu
            _DockWindow.Show((Not _DockWindow.IsVisible()))
            Checked = _DockWindow.IsVisible()

            'Vérifier si le menu est visible
            If _DockWindow.IsVisible() Then
                'Initialiser les AddHandler
                Call m_MenuEdgeMatch.InitHandler()

                'si le menu n'est pas visible
            Else
                'Détruire les AddHandler
                Call m_MenuEdgeMatch.DeleteHandler()
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            Enabled = My.ArcMap.Application IsNot Nothing
        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub
End Class
