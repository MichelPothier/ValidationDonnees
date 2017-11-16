Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms

'**
'Nom de la composante : cmdDessinerLimite.vb
'
'''<summary>
''' Commande qui permet de dessiner les limites communes de découpage utilisé par le EdgeMatch.
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
Public Class cmdDessinerLimite
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
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface qui permet de créer un buffer

        Try
            'Dessiner la limite commune
            Call bDessinerGeometrie(m_LimiteDecoupage, False)

            'Interface pour créer le buffer
            pTopoOp = CType(m_LimiteDecoupage, ITopologicalOperator2)

            'Dessiner le buffer de la limite commune
            Call bDessinerGeometrie(pTopoOp.Buffer(m_TolRecherche), False)

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
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
