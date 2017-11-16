'**
'Nom de la composante : modBarreLayerSelection.vb 
'
'''<summary>
'''Librairies de routines utilisée pour manipuler les divers sélections et listes d'identifiants d'éléments
'''présentent dans les FeatureLayers visibles.
'''</summary>
'''
'''<remarks>
'''Auteur : Michel Pothier
'''</remarks>
''' 
Friend Module modBarreLayerSelection

    '''<summary>
    '''Routine qui permet de créer une nouvelle liste d'identifiants d'éléments dans les FeatureLayers visibles
    '''pour lesquels il y a une sélection d'éléments.
    '''</summary>
    ''' 
    '''<param name="pApplication"> Interface ESRI contenant l'application ArcMap.</param>
    '''<param name="bDeleteLayer"> Indique si on doit détruire le Layer d'origine.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub CreerListeIdentifiant(ByVal pApplication As IApplication, Optional ByVal bDeleteLayer As Boolean = False)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing      'Interface contenant les paramètres d'affichage d'un Layer
        Dim pNewGeoFeatureLayer As IGeoFeatureLayer = Nothing   'Interface contenant les nouveaux paramètres d'affichage d'un Layer
        Dim pRemoveLayersOperation As IRemoveLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour le retrait d'un FeatureLayer
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour l'ajout d'un FeatureLayer
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim pGroupLayer As IGroupLayer = Nothing                'Interface contenant le GroupLayer qui contient le FeatureLayer traité
        Dim nPos As Integer = 0                                 'Position du Layer dans le GroupLayer
        Dim i As Integer = Nothing                              'Compteur
        Dim j As Integer = 0                                    'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
            oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
            qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

            'Traiter tous les FeatureLayers
            For i = 1 To qFeatureLayerColl.Count
                'Définir le FeatureLayer à traiter
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                'Vérifier si des éléments sont sélectionnés dans le layer
                If pFeatureSel.SelectionSet.Count > 0 Then
                    'Extraire le Group du FeatureLayer
                    pGroupLayer = oGererMapLayer.GroupLayer(pFeatureLayer, nPos)
                    'Interface pour créer un FeatureLayer selon la sélection
                    pFLDef = CType(pFeatureLayer, IFeatureLayerDefinition)
                    pNewFeatureLayer = pFLDef.CreateSelectionLayer(pFeatureLayer.Name, True, vbNullString, pFLDef.DefinitionExpression)
                    'Conserver la représentation graphique
                    pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
                    pNewGeoFeatureLayer = CType(pNewFeatureLayer, IGeoFeatureLayer)
                    pNewGeoFeatureLayer.Renderer = pGeoFeatureLayer.Renderer
                    'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                    pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                    pAddLayersOperation.AddLayer(pNewFeatureLayer)
                    pAddLayersOperation.ArrangeLayers = False
                    pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                    pAddLayersOperation.SetDestinationInfo(nPos, pMxDoc.FocusMap, pGroupLayer)
                    pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))
                    'Vérifier si on doit détruire l'ancien Layer
                    If bDeleteLayer Then
                        'Détruire l'ancien Layer
                        pRemoveLayersOperation = CType(New RemoveLayersOperation, IRemoveLayersOperation)
                        pRemoveLayersOperation.AddLayerInfo(pFeatureLayer, pMxDoc.FocusMap, pGroupLayer)
                        pMxDoc.OperationStack.Do(CType(pRemoveLayersOperation, IOperation))
                    End If
                End If
            Next i

            'Vérifier si un traitement a été effectué
            If j > 0 Then
                'Rafraîchir les éléments sélectionnés et la liste des Layers
                pMxDoc.ActiveView.Refresh()
                pMxDoc.ContentsView(0).Refresh(0)
            Else
                'Message d'avertissement
                MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pFLDef = Nothing
            pGeoFeatureLayer = Nothing
            pNewGeoFeatureLayer = Nothing
            pAddLayersOperation = Nothing
            pRemoveLayersOperation = Nothing
            oGererMapLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de retirer les éléments sélectionnés de la liste d'identifiants d'éléments contenue 
    '''dans chacun des FeatureLayers visibles.
    '''</summary>
    ''' 
    '''<param name="pApplication"> Interface ESRI contenant l'application ArcMap.</param>
    '''<param name="bDeleteLayer"> Indique si on doit détruire le Layer d'origine.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub RetirerListeIdentifiant(ByVal pApplication As IApplication, Optional ByVal bDeleteLayer As Boolean = False)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing      'Interface contenant les paramètres d'affichage d'un Layer
        Dim pNewGeoFeatureLayer As IGeoFeatureLayer = Nothing   'Interface contenant les nouveaux paramètres d'affichage d'un Layer
        Dim pRemoveLayersOperation As IRemoveLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour le retrait d'un FeatureLayer
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour l'ajout d'un FeatureLayer
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim pGroupLayer As IGroupLayer = Nothing                'Interface contenant le GroupLayer qui contient le FeatureLayer traité
        Dim nPos As Integer = 0                                 'Position du Layer dans le GroupLayer
        Dim i As Integer = Nothing                              'Compteur
        Dim j As Integer = 0                                    'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
            oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
            qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

            'Traiter tous les FeatureLayers
            For i = 1 To qFeatureLayerColl.Count
                'Définir le FeatureLayer à traiter
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                'Vérifier si des éléments sont sélectionnés dans le Layer
                If pFeatureSel.SelectionSet.Count > 0 Then
                    'Compter les traitements effectués
                    j = j + 1
                    'Inverser la sélection de façon à retirer les éléments sélectionnés de la liste d'identifiants d'éléments
                    pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultXOR, False)
                    'Vérifier si des éléments sont sélectionnés dans le Layer
                    If pFeatureSel.SelectionSet.Count > 0 Then
                        'Extraire le Group du FeatureLayer
                        pGroupLayer = oGererMapLayer.GroupLayer(pFeatureLayer, nPos)
                        'Interface pour créer un FeatureLayer selon la sélection
                        pFLDef = CType(pFeatureLayer, IFeatureLayerDefinition)
                        pNewFeatureLayer = pFLDef.CreateSelectionLayer(pFeatureLayer.Name, True, vbNullString, pFLDef.DefinitionExpression)
                        'Conserver la représentation graphique
                        pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
                        pNewGeoFeatureLayer = CType(pNewFeatureLayer, IGeoFeatureLayer)
                        pNewGeoFeatureLayer.Renderer = pGeoFeatureLayer.Renderer
                        'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                        pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                        pAddLayersOperation.AddLayer(pNewFeatureLayer)
                        pAddLayersOperation.ArrangeLayers = False
                        pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                        pAddLayersOperation.SetDestinationInfo(nPos, pMxDoc.FocusMap, pGroupLayer)
                        pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))
                        'Vérifier s'il faut détruire l'ancien Layer
                        If bDeleteLayer Then
                            'Détruire l'ancien Layer
                            pRemoveLayersOperation = CType(New RemoveLayersOperation, IRemoveLayersOperation)
                            pRemoveLayersOperation.AddLayerInfo(pFeatureLayer, pMxDoc.FocusMap, pGroupLayer)
                            pMxDoc.OperationStack.Do(CType(pRemoveLayersOperation, IOperation))
                        End If
                    End If
                End If
            Next i

            'Vérifier si un traitement a été effectué
            If j > 0 Then
                'Rafraîchir les éléments sélectionnés et la liste des Layers
                pMxDoc.ActiveView.Refresh()
                pMxDoc.ContentsView(0).Refresh(0)
            Else
                'Message d'avertissement
                MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pFLDef = Nothing
            pGeoFeatureLayer = Nothing
            pNewGeoFeatureLayer = Nothing
            pAddLayersOperation = Nothing
            pRemoveLayersOperation = Nothing
            oGererMapLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de détruire la liste d'identifiants d'éléments contenue 
    '''dans chacun des FeatureLayers visibles.
    '''</summary>
    ''' 
    '''<param name="pApplication"> Interface ESRI contenant l'application ArcMap.</param> 
    '''<param name="bDeleteLayer"> Indique si on doit détruire le Layer d'origine.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub DetruireListeIdentifiant(ByVal pApplication As IApplication, Optional ByVal bDeleteLayer As Boolean = False)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pResultSet As ISelectionSet = Nothing               'Interface contenant la nouvelle sélection désirée
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pNewFLDef As IFeatureLayerDefinition = Nothing      'Interface utilisé pour créer un nouveau Layer de sélection
        Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing      'Interface contenant les paramètres d'affichage d'un Layer
        Dim pNewGeoFeatureLayer As IGeoFeatureLayer = Nothing   'Interface contenant les nouveaux paramètres d'affichage d'un Layer
        Dim pRemoveLayersOperation As IRemoveLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour le retrait d'un FeatureLayer
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo pour l'ajout d'un FeatureLayer
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim pGroupLayer As IGroupLayer = Nothing                'Interface contenant le GroupLayer qui contient le FeatureLayer traité
        Dim nPos As Integer = 0                                 'Position du Layer dans le GroupLayer
        Dim i As Integer = Nothing                              'Compteur
        Dim j As Integer = 0                                    'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
            oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
            qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

            'Traiter tous les FeatureLayers
            For i = 1 To qFeatureLayerColl.Count
                'Définir le FeatureLayer à traiter
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                'Interface pour extraire un SelectionSet
                pFLDef = CType(pFeatureLayer, IFeatureLayerDefinition)
                pResultSet = pFLDef.DefinitionSelectionSet
                'Vérifier la présence d'une liste d'identifiants dans le Layer
                If Not pResultSet Is Nothing Then
                    'Extraire le Group du FeatureLayer
                    pGroupLayer = oGererMapLayer.GroupLayer(pFeatureLayer, nPos)
                    'Ajouter le nouveau FeatureLayer sans sélection
                    pNewFeatureLayer = New FeatureLayer
                    pNewFeatureLayer.Name = pFeatureLayer.Name
                    pNewFeatureLayer.FeatureClass = pFeatureLayer.FeatureClass
                    'Ajouter la sélection
                    pFeatureSel = CType(pNewFeatureLayer, IFeatureSelection)
                    pFeatureSel.SelectionSet = pResultSet
                    'Conserver la requête attributive
                    pNewFLDef = CType(pNewFeatureLayer, IFeatureLayerDefinition)
                    pNewFLDef.DefinitionExpression = pFLDef.DefinitionExpression
                    'Conserver la représentation graphique
                    pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
                    pNewGeoFeatureLayer = CType(pNewFeatureLayer, IGeoFeatureLayer)
                    pNewGeoFeatureLayer.Renderer = pGeoFeatureLayer.Renderer
                    'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                    pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                    pAddLayersOperation.AddLayer(pNewFeatureLayer)
                    pAddLayersOperation.ArrangeLayers = False
                    pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                    pAddLayersOperation.SetDestinationInfo(nPos, pMxDoc.FocusMap, pGroupLayer)
                    pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))
                    'Vérifier si on doit détruire l'ancien Layer
                    If bDeleteLayer Then
                        'Détruire l'ancien Layer
                        pRemoveLayersOperation = CType(New RemoveLayersOperation, IRemoveLayersOperation)
                        pRemoveLayersOperation.AddLayerInfo(pFeatureLayer, pMxDoc.FocusMap, pGroupLayer)
                        pMxDoc.OperationStack.Do(CType(pRemoveLayersOperation, IOperation))
                    End If
                End If
            Next i

            'Vérifier si un traitement a été effectué
            If j > 0 Then
                'Rafraîchir les éléments sélectionnés et la liste des Layers
                pMxDoc.ActiveView.Refresh()
                pMxDoc.ContentsView(0).Refresh(0)
            Else
                'Message d'avertissement
                MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pResultSet = Nothing
            pFLDef = Nothing
            pNewFLDef = Nothing
            pGeoFeatureLayer = Nothing
            pNewGeoFeatureLayer = Nothing
            pAddLayersOperation = Nothing
            pRemoveLayersOperation = Nothing
            oGererMapLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de faire une union entre la première sélection d'éléments et toutes les autres sélections 
    '''d'éléments des FeatureLayers visibles qui possèdent la même FeatureClass.
    '''</summary>
    ''' 
    '''<param name="pApplication "> Interface ESRI contenant l'application ArcMap.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub UnionSelection(ByVal pApplication As IApplication)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pNewFeatureSel As IFeatureSelection = Nothing       'Interface utilisé pour extraire les nouveau éléments sélectionnés
        Dim pNewSelectionSet As ISelectionSet = Nothing         'Interface utilisé pour définir les nouveau éléments sélectionnés
        Dim pResultSet As ISelectionSet = Nothing               'Interface contenant le résultat de la sélection désiré
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim sResultat As String = Nothing                       'Contient le nom du Layer avec l'action effectuée
        Dim i As Integer = Nothing                              'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Vérifier si au moins un élément est sélectionné
            If pMxDoc.FocusMap.SelectionCount > 0 Then
                'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
                oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
                qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

                'Traiter tous les FeatureLayers
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer à traiter
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                    'Vérifier si des éléments sont sélectionnés dans le layer
                    If pFeatureSel.SelectionSet.Count > 0 Then
                        'Définir le nouveau FeatureLayer
                        If pNewFeatureLayer Is Nothing Then
                            'Demander le nom du FeatureLayer à créer
                            sResultat = pFeatureLayer.Name & "_union"
                            'Créer le FeatureLayer
                            pNewFeatureLayer = New FeatureLayer
                            pNewFeatureLayer.Name = sResultat
                            pNewFeatureLayer.FeatureClass = pFeatureLayer.FeatureClass
                            pNewFeatureSel = CType(pNewFeatureLayer, IFeatureSelection)
                            pNewSelectionSet = pFeatureSel.SelectionSet
                            pNewFeatureSel.SelectionSet = pFeatureSel.SelectionSet

                            'Vérifier s'il s'agit de la même classe
                        ElseIf pNewFeatureLayer.FeatureClass.AliasName = pFeatureLayer.FeatureClass.AliasName Then
                            'Ajouter la sélection des autres FeatureLayers
                            pNewSelectionSet.Combine(pFeatureSel.SelectionSet, esriSetOperation.esriSetUnion, pResultSet)
                            pNewFeatureSel.SelectionSet = pResultSet
                        Else
                            'Message d'avertissement
                            MsgBox("ATTENTION : La FeatureClass de <" & pFeatureLayer.Name _
                            & "> n'est pas la même que celle du premier FeatureLayer <" _
                            & pNewFeatureLayer.FeatureClass.AliasName & ">")
                        End If
                    End If
                Next i

                'Vérifier si le nouveau FeatureLayer est a été créé
                If Not pNewFeatureLayer Is Nothing Then
                    'Interface pour créer un FeatureLayer selon la sélection
                    pFLDef = CType(pNewFeatureLayer, IFeatureLayerDefinition)
                    pNewFeatureLayer = pFLDef.CreateSelectionLayer(pNewFeatureLayer.Name, True, vbNullString, vbNullString)

                    'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                    pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                    pAddLayersOperation.AddLayer(pNewFeatureLayer)
                    pAddLayersOperation.ArrangeLayers = False
                    pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                    pAddLayersOperation.SetDestinationInfo(0, pMxDoc.FocusMap, Nothing)
                    pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))

                    'Rafraîchir les éléments sélectionnés et la liste des Layers
                    pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                    pMxDoc.ContentsView(0).Refresh(0)
                Else
                    'Message d'avertissement
                    MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pNewFeatureSel = Nothing
            pNewSelectionSet = Nothing
            pResultSet = Nothing
            pFLDef = Nothing
            pAddLayersOperation = Nothing
            oGererMapLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de faire une intersection entre la première sélection d'éléments et toutes les autres sélections 
    '''d'éléments des FeatureLayers visibles qui possèdent la même FeatureClass.
    '''</summary>
    ''' 
    '''<param name="pApplication"> Interface ESRI contenant l'application ArcMap.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub IntersectSelection(ByVal pApplication As IApplication)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pNewFeatureSel As IFeatureSelection = Nothing       'Interface utilisé pour extraire les nouveau éléments sélectionnés
        Dim pNewSelectionSet As ISelectionSet = Nothing         'Interface utilisé pour définir les nouveau éléments sélectionnés
        Dim pResultSet As ISelectionSet = Nothing               'Interface contenant le résultat de la sélection désiré
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim sResultat As String = Nothing                       'Contient le nom du Layer avec l'action effectuée
        Dim i As Integer = Nothing                              'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Vérifier si au moins un élément est sélectionné
            If pMxDoc.FocusMap.SelectionCount > 0 Then
                'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
                oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
                qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

                'Traiter tous les FeatureLayers
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer à traiter
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                    'Vérifier si des éléments sont sélectionnés dans le layer
                    If pFeatureSel.SelectionSet.Count > 0 Then
                        'Définir le nouveau FeatureLayer
                        If pNewFeatureLayer Is Nothing Then
                            'Demander le nom du FeatureLayer à créer
                            sResultat = pFeatureLayer.Name & "_Intersect"
                            'Créer le FeatureLayer
                            pNewFeatureLayer = New FeatureLayer
                            pNewFeatureLayer.Name = sResultat
                            pNewFeatureLayer.FeatureClass = pFeatureLayer.FeatureClass
                            pNewFeatureSel = CType(pNewFeatureLayer, IFeatureSelection)
                            pNewSelectionSet = pFeatureSel.SelectionSet
                            pNewFeatureSel.SelectionSet = pFeatureSel.SelectionSet

                            'Vérifier s'il s'agit de la même classe
                        ElseIf pNewFeatureLayer.FeatureClass.AliasName = pFeatureLayer.FeatureClass.AliasName Then
                            'Ajouter la sélection des autres FeatureLayers
                            pNewSelectionSet.Combine(pFeatureSel.SelectionSet, esriSetOperation.esriSetIntersection, pResultSet)
                            pNewFeatureSel.SelectionSet = pResultSet
                        Else
                            'Message d'avertissement
                            MsgBox("ATTENTION : La FeatureClass de <" & pFeatureLayer.Name _
                            & "> n'est pas la même que celle du premier FeatureLayer <" _
                            & pNewFeatureLayer.FeatureClass.AliasName & ">", )
                        End If
                    End If
                Next i

                'Vérifier si le nouveau FeatureLayer est a été créé
                If Not pNewFeatureLayer Is Nothing Then
                    'Interface pour créer un FeatureLayer selon la sélection
                    pFLDef = CType(pNewFeatureLayer, IFeatureLayerDefinition)
                    pNewFeatureLayer = pFLDef.CreateSelectionLayer(pNewFeatureLayer.Name, True, vbNullString, vbNullString)

                    'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                    pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                    pAddLayersOperation.AddLayer(pNewFeatureLayer)
                    pAddLayersOperation.ArrangeLayers = False
                    pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                    pAddLayersOperation.SetDestinationInfo(0, pMxDoc.FocusMap, Nothing)
                    pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))

                    'Rafraîchir les éléments sélectionnés et la liste des Layers
                    pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                    pMxDoc.ContentsView(0).Refresh(0)
                Else
                    'Message d'avertissement
                    MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pNewFeatureSel = Nothing
            pNewSelectionSet = Nothing
            pResultSet = Nothing
            pFLDef = Nothing
            pAddLayersOperation = Nothing
            oGererMapLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de faire une différence entre la première sélection d'éléments et toutes les autres sélections 
    '''d'éléments des FeatureLayers visibles qui possèdent la même FeatureClass.    
    '''</summary>
    ''' 
    '''<param name="pApplication"> Interface ESRI contenant l'application ArcMap.</param>
    ''' 
    '''<remarks>
    '''Auteur : Michel Pothier
    '''</remarks>
    Public Sub DifferenceSelection(ByVal pApplication As IApplication)
        'Déclarer les variables de travail
        Dim pMxDoc As IMxDocument = Nothing                     'Interface contenant le document ArcMap
        Dim qFeatureLayerColl As Collection = Nothing           'Collection des FeatureLayers à traiter
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le FeatureLayer traité
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire les éléments sélectionnés
        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pNewFeatureSel As IFeatureSelection = Nothing       'Interface utilisé pour extraire les nouveau éléments sélectionnés
        Dim pNewSelectionSet As ISelectionSet = Nothing         'Interface utilisé pour définir les nouveau éléments sélectionnés
        Dim pResultSet As ISelectionSet = Nothing               'Interface contenant le résultat de la sélection désiré
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection
        Dim pAddLayersOperation As IAddLayersOperation = Nothing 'Interface utilisé pour traiter le Undo/Redo
        Dim oGererMapLayer As clsGererMapLayer = Nothing        'Objet qui permet de gérer les Layers dans une Map
        Dim sResultat As String = Nothing                       'Contient le nom du Layer avec l'action effectuée
        Dim i As Integer = Nothing                              'Compteur

        Try
            'Interface d'accès au données dans le document
            pMxDoc = CType(pApplication.Document, IMxDocument)

            'Vérifier si au moins un élément est sélectionné
            If pMxDoc.FocusMap.SelectionCount > 0 Then
                'Remplir la liste des noms de FeatureLayers à traiter qui sont visibles
                oGererMapLayer = New clsGererMapLayer(pMxDoc.FocusMap)
                qFeatureLayerColl = oGererMapLayer.CollectionFeatureLayer(False)

                'Traiter tous les FeatureLayers
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer à traiter
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                    'Vérifier si des éléments sont sélectionnés dans le layer
                    If pFeatureSel.SelectionSet.Count > 0 Then
                        'Vérifier si le nouveau FeatureLayer est présent
                        If pNewFeatureLayer Is Nothing Then
                            'Demander le nom du FeatureLayer à créer
                            sResultat = pFeatureLayer.Name & "_difference"
                            'Créer le nouveau FeatureLayer
                            pNewFeatureLayer = New FeatureLayer
                            pNewFeatureLayer.Name = sResultat
                            pNewFeatureLayer.FeatureClass = pFeatureLayer.FeatureClass
                            pNewFeatureSel = CType(pNewFeatureLayer, IFeatureSelection)
                            pNewSelectionSet = pFeatureSel.SelectionSet
                            pNewFeatureSel.SelectionSet = pFeatureSel.SelectionSet

                            'Vérifier s'i s'agit de la même classe
                        ElseIf pNewFeatureLayer.FeatureClass.AliasName = pFeatureLayer.FeatureClass.AliasName Then
                            'Ajouter la sélection des autres FeatureLayers
                            pNewSelectionSet.Combine(pFeatureSel.SelectionSet, esriSetOperation.esriSetDifference, pResultSet)
                            pNewFeatureSel.SelectionSet = pResultSet
                        Else
                            'Afficher un message d'information
                            MsgBox("ATTENTION : La FeatureClass de <" & pFeatureLayer.Name _
                            & "> n'est pas la même que celle du premier FeatureLayer <" _
                            & pNewFeatureLayer.FeatureClass.AliasName & ">")
                        End If
                    End If
                Next i

                'Vérifier si le nouveau FeatureLayer est a été créé
                If Not pNewFeatureLayer Is Nothing Then
                    'Interface pour créer un FeatureLayer selon la sélection
                    pFLDef = CType(pNewFeatureLayer, IFeatureLayerDefinition)
                    pNewFeatureLayer = pFLDef.CreateSelectionLayer(pNewFeatureLayer.Name, True, vbNullString, vbNullString)

                    'Ajouter le nouveau Layer et le mettre dans le UnDo/ReDo
                    pAddLayersOperation = CType(New AddLayersOperation, IAddLayersOperation)
                    pAddLayersOperation.AddLayer(pNewFeatureLayer)
                    pAddLayersOperation.ArrangeLayers = False
                    pAddLayersOperation.Name = "ADD" & pNewFeatureLayer.Name
                    pAddLayersOperation.SetDestinationInfo(0, pMxDoc.FocusMap, Nothing)
                    pMxDoc.OperationStack.Do(CType(pAddLayersOperation, IOperation))

                    'Rafraîchir les éléments sélectionnés et la liste des Layers
                    pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                    pMxDoc.ContentsView(0).Refresh(0)
                Else
                    'Message d'avertissement
                    MsgBox("Attention : Les éléments sélectionnés ne font pas parties des FeatureLayers visibles")
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox(erreur.ToString)
        Finally
            'Vider la mémoire
            pMxDoc = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pNewFeatureLayer = Nothing
            pNewFeatureSel = Nothing
            pNewSelectionSet = Nothing
            pResultSet = Nothing
            pFLDef = Nothing
            pAddLayersOperation = Nothing
            oGererMapLayer = Nothing
        End Try
    End Sub

End Module
