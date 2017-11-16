'Nom de la composante : clsGererMapLayer.vb
'
'''<summary>
''' Librairie de Classe qui permet de manipuler les différents types de Layer contenu dans une Map.
'''</summary>
'''
'''<remarks>
'''Cette librairie est utilisable pour les outils interactifs ou Batch dans ArcMap (ArcGis de ESRI).
'''
'''Auteur : Michel Pothier
'''Date : 6 Mai 2011
'''</remarks>
''' 
Public Class clsGererMapLayer
    'Déclarer les variables globales
    '''<summary>Interface contenant la Map à Gérer.</summary>
    Private gpMap As IMap
    '''<summary>Numéro du code spécifique de l'incohérence.</summary>
    Private gnCodeSpecifique As Integer

#Region "Propriétés"
    '''<summary>
    '''Propriété qui permet de définir et retourner la Map traitée.
    '''</summary>
    ''' 
    Public Property Map() As IMap
        Get
            Map = gpMap
        End Get
        Set(ByVal value As IMap)
            gpMap = value
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    '''Routine qui permet d'initialiser la classe.
    '''</summary>
    '''
    Public Sub New(ByVal pMap As IMap)
        'Définir les valeur par défaut
        gpMap = pMap
    End Sub

    '''<summary>
    '''Routine qui permet de vider la mémoire des objets de la classe.
    '''</summary>
    '''
    Protected Overrides Sub finalize()
        gpMap = Nothing
    End Sub

    '''<summary>
    ''' Fonction qui permet d'extraire la collection des FeatureLayers présents dans la Map.
    ''' On peut indiquer si on veut aussi extraire les FeatureLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="bNonVisible"> Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non selon ce qui est demandé.</return>
    ''' 
    Public Function CollectionFeatureLayer(ByVal bNonVisible As Boolean) As Collection
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing            'Interface contenant un Groupe de Layers
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant un FeatureLayer
        Dim qFeatureLayerColl As Collection = Nothing       'Collection de FeatureLayer
        Dim i As Integer = Nothing                          'Compteur

        'Définir la collection de FeatureLayer vide
        CollectionFeatureLayer = New Collection

        Try
            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si on tient on doit extraire le Layer même s'il n'est pas visible
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un FeatureLayer
                    If TypeOf pLayer Is IFeatureLayer Then
                        'Définir le FeatureLayer
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Vérifier la présence de la FeatureClass
                        If Not pFeatureLayer.FeatureClass Is Nothing Then
                            'Ajouter un nouveau FeatureLayer dans la collection
                            CollectionFeatureLayer.Add(pFeatureLayer)
                        End If

                        'Vérifier les autres Layer dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer = CType(pLayer, IGroupLayer)

                        'Trouver les autres FeatureLayer dans un GroupLayer
                        Call CollectionFeatureLayerGroup(CollectionFeatureLayer, pGroupLayer, bNonVisible)
                    End If
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
            pFeatureLayer = Nothing
            qFeatureLayerColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire le GroupLayer dans lequel le Layer recherché est présent.
    '''</summary>
    '''
    '''<param name="pLayerRechercher">Interface contenant le Layer à rechercher dans la Map active.</param>
    '''<param name="nPosition">Position su Layer dans le GroupLayer.</param>
    ''' 
    '''<returns>"Collection" contenant les "IGroupLayer" recherchés.</returns>
    ''' 
    Public Function GroupLayer(ByVal pLayerRechercher As ILayer, ByRef nPosition As Integer) As IGroupLayer
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing              'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing    'Interface contenant un GroupLayer
        Dim i As Integer = Nothing                  'Compteur

        'Définir la valeur par défaut
        GroupLayer = Nothing

        Try
            'Vérifier si le Layer est valide
            If pLayerRechercher Is Nothing Then Return Nothing

            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer IsNot pLayerRechercher Then
                    'Vérifier les autres Layer dans un GroupLayer
                    If TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer = CType(pLayer, IGroupLayer)

                        'Trouver les autres GroupLayer dans un GroupLayer
                        GroupLayer = GroupLayerGroup(pGroupLayer, pLayerRechercher, nPosition)
                    End If

                    'Sortir de la boucle si le GroupLayer a été trouvé
                    If Not GroupLayer Is Nothing Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si le FeatureLayer est visible ou non dans la IMap.
    '''</summary>
    ''' 
    '''<param name="pLayerRechercher"> Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="bPresent"> Contient l'indication si le Layer à rechercher est présent dans la Map.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non.</return>
    ''' 
    Public Function EstVisible(ByVal pLayerRechercher As ILayer, ByVal bPresent As Boolean) As Boolean
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing              'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing    'Interface contenant un Groupe de Layers
        Dim i As Integer = Nothing                  'Compteur

        Try
            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer Is pLayerRechercher Then
                    'Retourner l'indication s'il est visible ou non
                    EstVisible = pLayer.Visible

                    'Sortir de la recherche
                    Exit For

                    'Si ce n'est pas le Layer recherché et que c'est un GroupLayer
                ElseIf TypeOf pLayer Is IGroupLayer Then
                    'Définir le GroupLayer
                    pGroupLayer = CType(pLayer, IGroupLayer)

                    'Retourner l'indication s'il est visible ou non
                    EstVisible = EstVisibleGroup(pGroupLayer, pLayer, bPresent)

                    'Sortir si le Layer est présent dans le GroupLayer
                    If bPresent Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet d'extraire la collection des FeatureLayers contenus dans un GroupLayer.
    ''' On peut indiquer si on veut aussi extraire les FeatureLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="qFeatureLayerColl">Collection des FeatureLayer.</param>
    '''<param name="pGroupLayer">Interface ESRI contenant un group de Layers.</param>
    '''<param name="bNonVisible">Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    ''' 
    Private Sub CollectionFeatureLayerGroup(ByRef qFeatureLayerColl As Collection, ByVal pGroupLayer As IGroupLayer, ByVal bNonVisible As Boolean)
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un GroupLayer
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant un FeatureLayer
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer
        Dim i As Integer = Nothing                          'Compteur

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si on tient compte du selectable
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un FeatureLayer
                    If TypeOf pLayer Is IFeatureLayer Then
                        'Définir le FeatureLayer
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Vérifier la présence de la FeatureClass
                        If Not pFeatureLayer.FeatureClass Is Nothing Then
                            'Ajouter un nouveau FeatureLayer dans la collection
                            qFeatureLayerColl.Add(pFeatureLayer)
                        End If

                        'Vérifier les autres noms dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer2 = CType(pLayer, IGroupLayer)

                        'Trouver les autres FeatureLayer dans un GroupLayer
                        Call CollectionFeatureLayerGroup(qFeatureLayerColl, pGroupLayer2, bNonVisible)
                    End If
                End If
            Next i

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer2 = Nothing
            pFeatureLayer = Nothing
            pCompositeLayer = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la collection des GroupLayers contenus dans un GroupLayer.
    '''</summary>
    ''' 
    '''<param name="pGroupLayer">Interface ESRI contenant un groupe de Layers.</param>
    '''<param name="pLayerRechercher">Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="nPosition">Position su Layer dans le GroupLayer.</param>
    ''' 
    Private Function GroupLayerGroup(ByVal pGroupLayer As IGroupLayer, ByVal pLayerRechercher As ILayer, ByRef nPosition As Integer) As IGroupLayer
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un GroupLayer
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer
        Dim i As Integer = Nothing                          'Compteur

        'Initialiser les variables de travail
        GroupLayerGroup = Nothing

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayerRechercher Is pLayer Then
                    'Retourner le Groupe du Layer recherché
                    GroupLayerGroup = pGroupLayer
                    'Définir la position du Layer dans le GroupLayer
                    nPosition = i
                    'Sortir
                    Exit For
                Else
                    'Vérifier les autres noms dans un GroupLayer
                    If TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer2 = CType(pLayer, IGroupLayer)

                        'Trouver les autres GroupLayer dans un GroupLayer
                        GroupLayerGroup = GroupLayerGroup(pGroupLayer2, pLayerRechercher, nPosition)
                    End If

                    'Sortir
                    If Not GroupLayerGroup Is Nothing Then Exit For
                End If
            Next i

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer2 = Nothing
            pCompositeLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si le FeatureLayer est visible ou non dans la IMap.
    '''</summary>
    ''' 
    '''<param name="pGroupLayer">Interface ESRI contenant un group de Layers.</param>
    '''<param name="pLayerRechercher"> Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="bPresent"> Contient l'indication si le Layer à rechercher est présent dans la Map.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non.</return>
    ''' 
    Private Function EstVisibleGroup(ByVal pGroupLayer As IGroupLayer, ByVal pLayerRechercher As ILayer, ByVal bPresent As Boolean) As Boolean
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un Groupe de Layers
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer
        Dim i As Integer = Nothing                          'Compteur

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer Is pLayerRechercher Then
                    'Retourner l'indication s'il est visible ou non
                    EstVisibleGroup = pLayer.Visible

                    'Sortir de la recherche
                    Exit For

                    'Si ce n'est pas le Layer recherché et que c'est un GroupLayer
                ElseIf TypeOf pLayer Is IGroupLayer Then
                    'Définir le GroupLayer
                    pGroupLayer2 = CType(pLayer, IGroupLayer)

                    'Retourner l'indication s'il est visible ou non
                    EstVisibleGroup = EstVisibleGroup(pGroupLayer2, pLayer, bPresent)

                    'Sortir si le Layer est présent dans le GroupLayer
                    If bPresent Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
            pCompositeLayer = Nothing
        End Try
    End Function
#End Region
End Class