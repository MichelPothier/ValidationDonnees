Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Data.OleDb
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry

'**
'Nom de la composante : cboFeatureLayer.vb
'
'''<summary>
''' Commande qui permet de définir le FeatureLayer dans lequel les éléments seront sélectionnés selon les paramètres spécifiés.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Public Class cboFeatureLayer
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    'Déclarer les variables de travail
    Dim gpApp As IApplication = Nothing             'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing            'Interface ESRI contenant un document ArcMap

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

                    'Conserver le document
                    m_MxDocument = gpMxDoc

                    'Remplir le ComboBox
                    Call RemplirComboBox()

                    'Conserver le lien avec le combo box
                    m_cboFeatureLayer = Me

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overloads Overrides Sub OnFocus(ByVal focus As Boolean)
        Try
            'Remplir le ComboBox
            Call RemplirComboBox()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overloads Overrides Sub OnSelChange(ByVal cookie As Integer)
        Try
            'Si aucun n'est sélectionné
            If cookie = -1 Then
                'On ne fait rien
                Exit Sub
            End If

            'Définir le FeatureLayer de sélection. 
            m_FeatureLayer = m_MapLayer.ExtraireFeatureLayerByName(CStr(Me.GetItem(cookie).Tag), True)

            'Remplir la liste des attributs
            m_cboAttributGroupe.initialiser()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Me.Enabled = True
    End Sub

#Region "Routines et fonctions privées"
    '''<summary>
    ''' Routine qui permet de remplir le ComboBox à partir des noms de FeatureLayer contenus dans la Map active.
    '''</summary>
    ''' 
    Public Sub RemplirComboBox()
        'Déclarer la variables de travail
        Dim qFeatureLayerColl As Collection = Nothing   'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant une classe de données
        Dim sNom As String = Nothing                    'Nom du FeatureLayer sélectionné
        Dim iDefaut As Integer = Nothing                'Index par défaut
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Initialiser le nom du FeatureLayer
            sNom = ""
            'Vérifier si le FeatureLayer est valide
            If m_FeatureLayer IsNot Nothing Then
                'Conserver le FeatureLayer sélectionné
                sNom = m_FeatureLayer.Name
            End If
            'Effacer tous les FeatureLayer existants
            Me.Clear()
            m_FeatureLayer = Nothing
            'Définir l'objet pour extraire les FeatureLayer
            m_MapLayer = New clsGererMapLayer(gpMxDoc.FocusMap, True)
            'Définir la liste des FeatureLayer
            qFeatureLayerColl = m_MapLayer.FeatureLayerCollection
            'Vérifier si les FeatureLayers sont présents
            If qFeatureLayerColl IsNot Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Ajouter le FeatureLayer
                    iDefaut = Me.Add(pFeatureLayer.Name, pFeatureLayer.Name)
                    'Vérifier si le nom du FeatureLayer correspond à celui sélectionné
                    If pFeatureLayer.Name = sNom Then
                        'Sélectionné la valeur par défaut
                        Me.Select(iDefaut)
                    End If
                Next
            End If

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            sNom = Nothing
            iDefaut = Nothing
            i = Nothing
        End Try
    End Sub
#End Region
End Class
