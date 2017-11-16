Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

'**
'Nom de la composante : cmdActiverMenu.vb
'
'''<summary>
'''Permet d'activer un menu pour afficher et modifier les paramètres de la barre de sélection.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
'''</remarks>
''' 
Public Class cmdActiverMenu
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private Shared _DockWindow As ESRI.ArcGIS.Framework.IDockableWindow

    Public Sub New()
        Try
            'Définir les variables de travail
            Dim windowID As UID = New UIDClass

            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                m_Application = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Rendre active la commande
                    Enabled = True
                    'Définir le document
                    m_MxDocument = CType(m_Application.Document, IMxDocument)

                    'Créer un nouveau menu
                    windowID.Value = "MPO_BarreSelection_dckMenuSelection"
                    _DockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

                    'Initialiser le menu des paramètres de sélection
                    Call m_MenuParametresSelection.Init()

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
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
                Call m_MenuParametresSelection.InitHandler()

                'si le menu n'est pas visible
            Else
                'Détruire les AddHandler
                Call m_MenuParametresSelection.DeleteHandler()
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si aucun Layer n'est présent
        If m_MxDocument.FocusMap.LayerCount = 0 Then
            Enabled = False
        Else
            Enabled = True
        End If
    End Sub
End Class
