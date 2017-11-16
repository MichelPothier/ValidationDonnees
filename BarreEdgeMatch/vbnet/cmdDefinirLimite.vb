Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms

'**
'Nom de la composante : cmdDefinirLimite.vb
'
'''<summary>
''' Commande qui permet de définir et afficher les limites communes entre les éléments de la classe découpage 
''' qui sont présents dans la fenêtre graphique.
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
Public Class cmdDefinirLimite
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
        Dim pMouseCursor As IMouseCursor = Nothing

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Permet de définir les limites de découpage commune
            Call DefinirLimiteDecoupage()

            'Vérifier si la limite commune est présente
            If m_LimiteDecoupage IsNot Nothing Then
                'Vérifier si la limite commune n'est pas vide 
                If Not m_LimiteDecoupage.IsEmpty Then
                    'Rafraîchir la vue active
                    m_MxDocument.ActiveView.Refresh()

                    'Permet de vider la mémoire sur les évènements
                    System.Windows.Forms.Application.DoEvents()

                    'Dessiner la limite commune
                    Call bDessinerGeometrie(m_LimiteDecoupage, False)
                End If
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si le Layer de découpage est présent
        If m_FeatureLayerDecoupage Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si la classe de découpage est présent
            If m_FeatureLayerDecoupage.FeatureClass Is Nothing Then
                'Désactiver la commande
                Enabled = False
            Else
                'Activer la commande
                Enabled = True
            End If
        End If
    End Sub
End Class
