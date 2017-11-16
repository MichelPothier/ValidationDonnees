Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry

'**
'Nom de la composante : cmdDecouperLimiteFenetre.vb
'
'''<summary>
''' Commande qui permet de découper et afficher les limites communes selon l'enveloppe correspondant à la fenêtre graphique.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 7 novembre 2013
'''</remarks>
''' 
Public Class cmdDecouperLimiteFenetre
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
        Dim pMouseCursor As IMouseCursor = Nothing      'Interface qui permet de changer l'image du curseur
        Dim pTopoOp As ITopologicalOperator = Nothing   'Interface qui permet de découper la limite selon un polygon

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Permet de découper les limites de découpage commune selon l'enveloppe de la fenêtre graphique
            pTopoOp = CType(m_LimiteDecoupage, ITopologicalOperator)
            pTopoOp.Clip(gpMxDoc.ActiveView.Extent)

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
            pTopoOp = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si la limite est présente
        If m_LimiteDecoupage Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si la limite est présente
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
