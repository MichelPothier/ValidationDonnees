Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsDecoupageElement.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont l’attribut et la géométrie de découpage respecte ou non
''' l'élément de la classe de découpage spécifiée.
''' 
''' La classe permet de traiter un seul attribut qui contient la valeur de l'identifiant de découpage à valider.
''' 
''' Note : L'attribut spécifié correspond au nom de l'attribut présent dans la classe de découpage et qui contient la géométrie
'''        et l'identifiant de découpage à traiter.     
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 2 juin 2015
'''</remarks>
''' 
Public Class clsDecoupageElement
    Inherits clsValeurAttribut

    '''<summary>Contient la position de l'attribut de découpage dans la classe de découpage.</summary>
    Protected giPosAttributDecoupage As Integer = -1
    '''<summary>Contient l'identifiant de l'élément de découpage.</summary>
    Protected gsIdentifiantDecoupage As String = ""

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "DATASET_NAME"
            Expression = "DATASET_NAME"

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sNomAttribut"> Nom de l'attribut à traiter.</param>
    '''<param name="sExpression"> Expression régulière à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String,
                   Optional ByVal bLimite As Boolean = False, Optional ByVal bUnique As Boolean = False)
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        giPosAttributDecoupage = Nothing
        gsIdentifiantDecoupage = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.finalize()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de retourner le nom de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides ReadOnly Property Nom() As String
        Get
            Nom = "DecoupageElement"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides Property Parametres() As String
        Get
            'Retourner la valeur des paramètres
            Parametres = gsNomAttribut & " " & gsExpression
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

            Try
                'Extraire les paramètres
                params = value.Split(CChar(" "))
                'Vérifier si les deux paramètres sont présents
                If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION")

                'Définir les valeurs par défaut
                gsNomAttribut = params(0)
                gsExpression = params(1)

            Catch ex As Exception
                'Retourner l'erreur
                Throw ex
            Finally
                'Vider la mémoire
                params = Nothing
            End Try
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Public Overloads Overrides Function ListeParametres() As Collection
        'Déclarer les variables de travail
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface contenant le résultat de la recherche.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing              'Interface contenant le premier élément de la classeà traiter.

        'Définir la liste des paramètres par défaut
        ListeParametres = New Collection
        'Définir le premier paramètre
        ListeParametres.Add(gsNomAttribut + " " + gsNomAttribut)

        Try
            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Vérifier si FeatureLayer est valide
                If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                    'Interface pour extraire et sélectionner les éléments
                    pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

                    'Vérifier si au moins un élément est sélectionné
                    If pFeatureSel.SelectionSet.Count = 0 Then
                        'Interface pour extraire le premier élément
                        pFeatureCursor = gpFeatureLayerSelection.Search(Nothing, False)
                    Else
                        'Interfaces pour extraire le premier élément sélectionné
                        pFeatureSel.SelectionSet.Search(Nothing, False, pCursor)
                        pFeatureCursor = CType(pCursor, IFeatureCursor)
                    End If

                    'Extraire l'élément de découpage
                    pFeature = pFeatureCursor.NextFeature

                    'Vérifier si un élémet est trouvé
                    If pFeature IsNot Nothing Then
                        'Extriare la position de l'attribut par défaut
                        giAttribut = gpFeatureLayerSelection.FeatureClass.Fields.FindField(gsNomAttribut)
                        'Vérifier si l'attribut par défaut est présent
                        If giAttribut > -1 Then
                            'Définir l'expression par défaut
                            gsExpression = gsNomAttribut & "='" & pFeature.Value(giAttribut).ToString & "'"
                            'Définir le paramètre avec identifiant
                            ListeParametres.Add(gsNomAttribut + " " + gsExpression)
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si la FeatureClass est valide.
    ''' 
    '''<return>Boolean qui indique si la FeatureClass est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function FeatureClassValide() As Boolean
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureClassValide = False
            gsMessage = "ERREUR : La FeatureClass est invalide."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'Vérifier si la FeatureClass est de type Polyline et Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'La contrainte est valide
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide"
                Else
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Point, MultiPoint, Polyline ou Polygon."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'expression utilisée pour extraire l'élément de découpage est valide.
    ''' 
    '''<return>Boolean qui indique si l'expression est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function ExpressionValide() As Boolean
        'Déclarer les variables de travail
        Dim pQueryFilter As IQueryFilter = Nothing      'Interface contenant la requête attributive.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface contenant le résultat de la recherche.

        'Définir la valeur par défaut
        ExpressionValide = False
        gsMessage = "ERREUR : L'expression est invalide"

        Try
            'Vérifier si l'identifiant de découpage est présent
            If Len(gsExpression) > 0 Then
                'Vérifier la présence de la classe de découpage
                If gpFeatureLayerDecoupage IsNot Nothing Then
                    'Vérifier si la FeatureClass est valide
                    If gpFeatureLayerDecoupage.FeatureClass IsNot Nothing Then
                        'vérifier si la FeatureClass de découpage est de type Polygon
                        If gpFeatureLayerDecoupage.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                            'Vérifier si l'identifiant de découpage est spécifié
                            If gsExpression.Contains("=") Then
                                'Extraire la position de l'attribut de découpage
                                giPosAttributDecoupage = gpFeatureLayerDecoupage.FeatureClass.Fields.FindField(gsExpression.Split(CChar("="))(0))

                                'Vérifier si l'attribut est absent dans la classe de découpage
                                If giPosAttributDecoupage = -1 Then
                                    gsMessage = "ERREUR : L'attribut de découpage est absent dans la classe de découpage."

                                    'Si l'attribut est présent dans la classe de découpage
                                Else
                                    'Définir l'identifiant de découpage
                                    gsIdentifiantDecoupage = gsExpression.Split(CChar("="))(1)
                                    gsIdentifiantDecoupage = gsIdentifiantDecoupage.Replace("'", "")

                                    'Définir la requête attributive pour extraire l'élément de découpage
                                    pQueryFilter = New QueryFilter
                                    pQueryFilter.WhereClause = gsExpression

                                    'Exécuter la recherche de la requête attributive
                                    pFeatureCursor = gpFeatureLayerDecoupage.Search(pQueryFilter, False)

                                    'Extraire l'élément de découpage
                                    FeatureDecoupage = pFeatureCursor.NextFeature

                                    'Vérifier si l'élément de découpage est présent
                                    If gpFeatureDecoupage IsNot Nothing Then
                                        'La contrainte est valide
                                        ExpressionValide = True
                                        gsMessage = "La contrainte est valide"
                                    Else
                                        gsMessage = "ERREUR : L'identifiant de découpage est absent de la classe de découpage."
                                    End If
                                End If

                                'Si l'identifiant de découpage n'est pas spécifié
                            Else
                                'Extraire la position de l'attribut de découpage
                                giPosAttributDecoupage = gpFeatureLayerDecoupage.FeatureClass.Fields.FindField(gsExpression)

                                'Vérifier si l'attribut est absent dans la classe de découpage
                                If giPosAttributDecoupage = -1 Then
                                    gsMessage = "ERREUR : L'attribut de découpage est absent dans la classe de découpage."

                                    'Si l'attribut est présent dans la classe de découpage
                                Else
                                    'Vérifier si l'élément de découpage est spécifié
                                    If gpFeatureDecoupage IsNot Nothing Then
                                        'Définir l'identifiant de découpage
                                        gsIdentifiantDecoupage = gpFeatureDecoupage.Value(giPosAttributDecoupage).ToString

                                        'Vérifier si l'identifiant n'est pas vide
                                        If gsIdentifiantDecoupage.Length > 0 Then
                                            'La contrainte est valide
                                            ExpressionValide = True
                                            gsMessage = "La contrainte est valide"

                                            'Si l'identifiant est vide
                                        Else
                                            gsMessage = "ERREUR : L'identifiant de la classe de découpage est vide."
                                        End If

                                        'si l'élément de découpage est spécifié
                                    Else
                                        'La contrainte est valide
                                        ExpressionValide = True
                                        gsMessage = "La contrainte est valide"
                                    End If
                                End If
                            End If

                        Else
                            gsMessage = "ERREUR : Le FeatureClass de découpage n'est pas de type Polygon."
                        End If

                    Else
                        gsMessage = "ERREUR : Le FeatureClass de découpage est invalide."
                    End If

                Else
                    gsMessage = "ERREUR : Le FeatureLayer de découpage n'est pas spécifié correctement."
                End If

            Else
                gsMessage = "ERREUR : L'attribut et l'identifiant de la classe de découpage sont absents."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pQueryFilter = Nothing
            pFeatureCursor = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'identifiant ou la géométrie de découpage sont invalides.
    '''
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function Selectionner(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                        'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface utilisé pour extraire les éléments à traiter.
        Dim pQueryFilterDef As IQueryFilterDefinition = Nothing 'Interface utilisé pour indiquer que l'on veut trier.

        Try
            'Sortir si la contrainte est invalide
            If Me.EstValide() = False Then Err.Raise(1, , Me.Message)

            'Définir la géométrie par défaut
            Selectionner = New GeometryBag
            Selectionner.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Conserver le type de sélection à effectuer
            gbEnleverSelection = bEnleverSelection

            ''Interface pour extraire et sélectionner les éléments
            'pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            ''Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            'pSelectionSet = pFeatureSel.SelectionSet

            ''Vérifier si des éléments sont sélectionnés
            'If pSelectionSet.Count = 0 Then
            '    'Sélectionnées tous les éléments du FeatureLayer
            '    pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            '    pSelectionSet = pFeatureSel.SelectionSet
            'End If

            ''Enlever la sélection
            'pFeatureSel.Clear()

            ''Créer une nouvelle requete vide
            'pQueryFilterDef = New QueryFilter
            ''Indiquer la méthode pour trier
            'pQueryFilterDef.PostfixClause = "ORDER BY " & gsNomAttribut
            ''Interfaces pour extraire les éléments sélectionnés
            'pSelectionSet.Search(CType(pQueryFilterDef, IQueryFilter), False, pCursor)
            'pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Traiter le FeatureLayer selon un découpage spécifique
            Selectionner = TraiterDecoupageEstInclus(pTrackCancel, bEnleverSelection)

            ''Vérifier si l'élément de découpage est spécifié
            'If gpFeatureDecoupage IsNot Nothing Then
            '    'Traiter le FeatureLayer selon un découpage spécifique
            '    Selectionner = TraiterDecoupageElementIdentifiant(pFeatureCursor, pTrackCancel, bEnleverSelection)

            '    'Si l'élément de découpage n'est pas spécifié
            'Else
            '    'Traiter le FeatureLayer selon le découpage des élément traités
            '    Selectionner = TraiterDecoupageElement(pFeatureCursor, pTrackCancel, bEnleverSelection)
            'End If

            'Vérifier si la classe d'erreur est présente
            If gpFeatureCursorErreur IsNot Nothing Then
                'Conserver toutes les modifications
                gpFeatureCursorErreur.Flush()
                'Release the update cursor to remove the lock on the input data.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(gpFeatureCursorErreur)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pQueryFilterDef = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur EST_INCLUS VRAI
    ''' avec ses éléments en relation.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments qui respectent ou non les relations spatiales.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterDecoupageEstInclus(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.
        Dim sIdDecSel(0) As String 'Vecteur des positions d'identifiants de découapge pour la classe de sélection.
        Dim sIdDecRel(0) As String 'Vecteur des positions d'identifiants de découapge pour la classe en relation.

        Try
            'Définir la géométrie par défaut
            TraiterDecoupageEstInclus = New GeometryBag
            TraiterDecoupageEstInclus.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterDecoupageEstInclus, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrieAttribut(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, sIdDecSel, giAttribut)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments de découpage (" & gpFeatureLayerDecoupage.Name & ") ..."
            'Lire les éléments en relation
            LireGeometrieAttribut(gpFeatureLayerDecoupage, pTrackCancel, pGeomRelColl, iOidRel, sIdDecRel, giPosAttributDecoupage)

            'Afficher le message de traitement de la relation spatiale
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale EST_INCLUS (" & gpFeatureLayerSelection.Name & "/" & gpFeatureLayerDecoupage.Name & ") ..."
            'Interface pour traiter la relation spatiale
            pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
            'Exécuter la recherche et retourner le résultat de la relation spatiale
            pRelResult = pRelOpNxM.Within(CType(pGeomRelColl, IGeometryBag))

            'Récupération de la mémoire disponible
            pRelOpNxM = Nothing
            GC.Collect()

            'Afficher le message de traitement d'extraction des OIDs trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & "/" & gpFeatureLayerDecoupage.Name & ") ..."
            'Extraire les Oids d'éléments trouvés
            ExtraireListeOid(pRelResult, iOidSel, iOidRel, iOidAdd)

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, iOidAdd, sIdDecSel, sIdDecRel, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
            sIdDecSel = Nothing
            sIdDecRel = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de lire les géométries et les OIDs des éléments d'un FeatureLayer.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomColl"> Interface contenant les géométries des éléments lus.</param>
    '''<param name="iOid"> Vecteur des OIDs d'éléments lus.</param>
    '''<param name="iIdDec"> Vecteur des positions d'identifiants de déciupage.</param>
    '''<param name="iPosAttribut"> Position de l'attribut de découpage.</param>
    ''' 
    '''</summary>
    '''
    Private Sub LireGeometrieAttribut(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel, ByRef pGeomColl As IGeometryCollection,
                                      ByRef iOid() As Integer, ByRef sIdDec() As String, ByVal iPosAttribut As Integer)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément lu.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour extraire la limite de la géométrie.
        Dim pQueryFilterDef As IQueryFilterDefinition = Nothing 'Interface utilisé pour indiquer que l'on veut trier.

        Try
            'Créer un nouveau Bag vide
            pGeometry = New GeometryBag

            'Définir la référence spatiale
            pGeometry.SpatialReference = pFeatureLayer.AreaOfInterest.SpatialReference
            pGeometry.SnapToSpatialReference()

            'Interface pour ajouter les géométries dans le Bag
            pGeomColl = CType(pGeometry, IGeometryCollection)

            'Interface pour sélectionner les éléments
            pFeatureSel = CType(pFeatureLayer, IFeatureSelection)

            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If

            'Interface pour extraire les éléments sélectionnés
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Augmenter le vecteur des Oid selon le nombre d'éléments
            ReDim Preserve iOid(pSelectionSet.Count)
            ReDim Preserve sIdDec(pSelectionSet.Count)

            'Créer une nouvelle requete vide
            pQueryFilterDef = New QueryFilter
            'Indiquer la méthode pour trier
            pQueryFilterDef.PostfixClause = "ORDER BY " & pFeatureLayer.FeatureClass.Fields.Field(iPosAttribut).Name

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(CType(pQueryFilterDef, IQueryFilter), False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments du FeatureLayer
            For i = 0 To pSelectionSet.Count - 1
                'Vérifier si l'élément est présent
                If pFeature IsNot Nothing Then
                    'Ajouter la position de l'identifiant
                    sIdDec(i) = pFeature.Value(iPosAttribut).ToString

                    'Définir la géométrie à traiter
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie à traiter
                    pGeometry.Project(pFeatureLayer.AreaOfInterest.SpatialReference)

                    'Ajouter la géométrie dans le Bag
                    pGeomColl.AddGeometry(pGeometry)

                    'Ajouter le OID de l'élément avec sa séquence 
                    iOid(i) = pFeature.OID

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                End If

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Enlever la sélection des éléments
            pFeatureSel.Clear()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
            pQueryFilterDef = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation spatiale.
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="iOidAdd"> Vecteur des OIDs de la classe de découpage trouvés.</param>
    ''' 
    '''</summary>
    Private Sub ExtraireListeOid(ByVal pRelResult As IRelationResult, _
                                 ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByRef iOidAdd() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Indiquer que le OID respecte la relation en conservant le OID de l'élément de découpage en relation
                'iOidAdd(iSel) = iOidRel(iRel)
                iOidAdd(iSel) = iRel + 1
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte la relation spatiale VRAI du masque à 9 intersections ou du SCL.
    ''' 
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on veut enlever les éléments trouvés de la sélection.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pSelectionSet">Interface contenant les éléments sélectionnés.</param>
    '''<param name="pGeomResColl">GéométryBag contenant les géométries en erreur.</param>
    ''' 
    '''</summary>
    '''
    Private Sub SelectionnerElementErreur(ByVal pGeomSelColl As IGeometryCollection, ByVal pGeomRelColl As IGeometryCollection, _
                                          ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal iOidAdd() As Integer, _
                                          ByVal sIdDecSel() As String, ByVal sIdDecRel() As String, _
                                          ByVal bEnleverSelection As Boolean, ByRef pTrackCancel As ITrackCancel, ByRef pSelectionSet As ISelectionSet, _
                                          ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la relation spatiale.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour définir la géométrie en erreur.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface utilisé pour extraire le nombre de composantes.
        Dim pGDBridge As IGeoDatabaseBridge2 = Nothing      'Interface pour ajouter une liste de OID dans le SelectionSet.
        Dim sMessageAttribut As String = ""         'Contient le message pour l'identifiant de découpage.
        Dim sMessageGeometrie As String = ""        'Contient le message pour la géométrie.
        Dim bSucces As Boolean = True               'Indiquer si toutes les conditions sont respectés.
        Dim iSecDec As Integer = 0  'Contient le numéro de séquence du découpage traité.
        Dim iOid(0) As Integer      'Liste des OIDs à sélectionner.
        Dim iNbOid As Integer = 0   'Compteur de OIDs à sélectionner.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si un seul découpage est présent
            If pGeomRelColl.GeometryCount = 1 Then
                'Définir le numéro de séquence de l'élément de découpage
                iSecDec = 0

                'Définir la géométrie de découpage
                gpPolygoneDecoupage = CType(pGeomRelColl.Geometry(0), IPolygon)
            End If

            'Traiter tous les éléments
            For i = 0 To pGeomSelColl.GeometryCount - 1
                'Indiquer que toutes les conditions sont respectés par défaut
                bSucces = True
                'Définir l'identifiant de découpage de l'élément traité
                gsIdentifiantDecoupage = sIdDecSel(i)
                'Définir le message de l'identifiant de découpage valide par défaut
                sMessageAttribut = " #L'identifiant de découpage est valide : " & gsIdentifiantDecoupage
                'Définir le message de la géométrie valide par défaut
                sMessageGeometrie = " #La géométrie de l'élément est à l'intérieur du polygone de découpage"

                'Définir la géométrie de l'élément
                pGeometry = pGeomSelColl.Geometry(i)

                'Vérifier si aucun OID de découpage n'a été trouvé
                If iOidAdd(i) = 0 Then
                    'Aucun numéro de séquence de découpage trouvé par défaut
                    iSecDec = -1

                    'Trouver le polygone de découpage
                    For j = 0 To pGeomRelColl.GeometryCount - 1
                        'Vérifier si l'identifiant correspond
                        If gsIdentifiantDecoupage = sIdDecRel(j) Then
                            'Définir le numéro de séquence de l'élément de découpage
                            iSecDec = j
                            'Définir le polygone de découpage
                            gpPolygoneDecoupage = CType(pGeomRelColl.Geometry(j), IPolygon)
                            'Sortir de la boucle
                            Exit For
                        End If
                    Next

                    'Vérifier si le polygone de découpage est présent
                    If gpPolygoneDecoupage IsNot Nothing Then
                        'Interface pour enlever la partie du polygone de découpage
                        pRelOp = CType(pGeometry, IRelationalOperator)

                        'Vérifier si la géométrie de découpage est disjoint avec le polygone de découpage
                        If pRelOp.Disjoint(gpPolygoneDecoupage) Then
                            'Indiquer que les conditions ne sont pas respectés
                            bSucces = False
                            'Définir le message d'erreur de la géométrie
                            sMessageGeometrie = " #La géométrie de l'élément est complètement à l'extérieur du polygone de découpage"

                            'Si la géométrie de découpage n'est pas disjoint avec le polygone de découpage
                        Else
                            'Vérifier si la géométrie traité n'est pas un point
                            If pGeometry.GeometryType <> esriGeometryType.esriGeometryPoint Then
                                'Interface pour enlever la partie du polygone de découpage
                                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                                'Enlever la partie avec le polygone du découpage
                                pGeometry = pTopoOp.Difference(gpPolygoneDecoupage)

                                'Vérifier si la géométrie est vide
                                If pGeometry.IsEmpty Then
                                    'Définir la géométrie de l'élément
                                    pGeometry = pGeomSelColl.Geometry(i)

                                    'Si la géométrie n'est pas vide
                                Else
                                    'Indiquer que les conditions ne sont pas respectés
                                    bSucces = False
                                    'Définir le message d'erreur de la géométrie
                                    sMessageGeometrie = " #La géométrie de l'élément n'est pas à l'intérieur du polygone de découpage"
                                End If
                            End If
                        End If

                        'Si aucun polygone de découpage trouvé
                    Else
                        'Indiquer que les conditions ne sont pas respectés
                        bSucces = False
                        'Définir le message d'erreur de la géométrie
                        sMessageGeometrie = " #L'identifiant de l'élément ne correspond à aucun polygone de découpage"
                    End If

                    'Si un élément de découpage est trouvé
                Else
                    'Vérifier si le découpage est différent
                    If iSecDec <> iOidAdd(i) - 1 Then
                        'Définir le numéro de séquence de l'élément de découpage
                        iSecDec = iOidAdd(i) - 1

                        'Définir la géométrie de découpage
                        gpPolygoneDecoupage = CType(pGeomRelColl.Geometry(iSecDec), IPolygon)
                    End If

                    'Si l'dentifiant de découpage est invalide
                    If gsIdentifiantDecoupage <> sIdDecRel(iSecDec) Then
                        'Indiquer que les condition ne sont pas respectés
                        bSucces = False

                        'Définir le message de l'identifiant de découpage
                        sMessageAttribut = " #L'identifiant de l'élément est différent de celui de découpage : '" & gsIdentifiantDecoupage & "' <> '" & sIdDecRel(iSecDec) & "'"
                    End If
                End If

                'Vérifier si on doit sélectionner l'élément traité
                If (bEnleverSelection And bSucces = False) Or (bEnleverSelection = False And bSucces) Then
                    'Définir la liste des OIDs à sélectionner
                    iNbOid = iNbOid + 1

                    'Redimensionner la liste des OIDs à sélectionner
                    ReDim Preserve iOid(iNbOid - 1)

                    'Définir le OIDs dans la liste
                    iOid(iNbOid - 1) = iOidSel(i)

                    'Ajouter la géométrie en erreur dans le Bag
                    pGeomResColl.AddGeometry(pGeometry)

                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & iOidSel(i).ToString & sMessageAttribut & sMessageGeometrie & " /" & Parametres, pGeometry, iSecDec)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                'Récupération de la mémoire disponble
                pRelOp = Nothing
                pTopoOp = Nothing
                pGeometry = Nothing
                sMessageAttribut = Nothing
                sMessageGeometrie = Nothing
                GC.Collect()
            Next

            'Interface pour enlever ou ajouter les OIDs trouvés dans le SelectionSet
            pGDBridge = New GeoDatabaseHelperClass()

            'Ajouter les OIDs trouvés dans le SelectionSet
            pGDBridge.AddList(pSelectionSet, iOid)

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRelOp = Nothing
            pTopoOp = Nothing
            pGeometry = Nothing
            pGDBridge = Nothing
            sMessageAttribut = Nothing
            sMessageGeometrie = Nothing
            bSucces = Nothing
            iSecDec = Nothing
            iOid = Nothing
            iNbOid = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'identifiant et la géométrie respecte ou non 
    ''' l'identifiant de l'élément de découpage spécifié.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterDecoupageElementIdentifiant(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                        Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément traité.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour valider la géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour définir la géométrie en erreur.
        Dim sMessageAttribut As String = ""                 'Contient le message pour l'identifiant.
        Dim sMessageGeometrie As String = ""                'Contient le message pour la géométrie.
        Dim bSucces As Boolean = False                      'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        Try
            'Définir la géométrie par défaut
            TraiterDecoupageElementIdentifiant = New GeometryBag
            TraiterDecoupageElementIdentifiant.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDecoupageElementIdentifiant, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Projeter le polygone de découpage
                gpPolygoneDecoupage.Project(TraiterDecoupageElementIdentifiant.SpatialReference)

                'Interface pour valider la géométrie
                pTopoOp = CType(gpPolygoneDecoupage, ITopologicalOperator2)
                pRelOp = CType(pTopoOp.Buffer(gdPrecision), IRelationalOperator)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialisation du traitement
                    bSucces = True
                    sMessageAttribut = "#L'identifiant de découpage de l'élément est valide"
                    sMessageGeometrie = "#La géométrie de l'élément est à l'intérieur du polygone de découpage"

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie de l'élément
                    pGeometry.Project(TraiterDecoupageElementIdentifiant.SpatialReference)

                    'Vérifier si l'identifiant est valide
                    If pFeature.Value(giAttribut).ToString <> gsIdentifiantDecoupage Then
                        'Définir l'erreur
                        bSucces = False
                        sMessageAttribut = "#L'identifiant de découpage de l'élément est invalide," & gsNomAttribut & "=" & pFeature.Value(giAttribut).ToString
                    End If

                    'Vérifier si la géométrie est valide
                    If Not pRelOp.Contains(pGeometry) Then
                        'Interface pour définir la géométrie en erreur
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        'Définir la géométrie en erreur
                        pGeometry = pTopoOp.Difference(gpPolygoneDecoupage)
                        'Vérifier si la géométrie n'est pas vide
                        If Not pGeometry.IsEmpty Then
                            'Définir l'erreur
                            bSucces = False
                            sMessageGeometrie = "#La géométrie de l'élément n'est pas à l'intérieur du polygone de découpage"
                        Else
                            'Rédéfinir la géométrie de l'élément
                            pGeometry = CType(pTopoOp, IGeometry)
                        End If
                    End If

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pGeometry)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                                            & sMessageAttribut & sMessageGeometrie _
                                            & "/" & gsExpression _
                                            & "/Précision=" & gdPrecision.ToString, pGeometry)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pRelOp = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'identifiant et la géométrie respecte ou non 
    ''' l'identifiant de son élément de découpage en relation de position.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterDecoupageElement(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément traité.
        Dim pPoint As IPoint = Nothing                      'Interface contenant le point du centre de l'élément.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant la relation spatiale de base.
        Dim oFeatureColl As Collection = Nothing            'Objet contenant la collection des éléments en relation.
        Dim pFeatureDecoupage As IFeature = Nothing         'Interface contenant l'élément de découpage.
        Dim pGeometryDecoupage As IGeometry = Nothing       'Interface contenant la géométrie de l'élément de découpage.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour valider la géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour définir la géométrie en erreur.
        Dim sIdentifiant As String = ""                     'Identifiant de découpage à valider.
        Dim sIdentifiantPrec As String = ""                 'Identifiant de découpage précédant.
        Dim sMessageAttribut As String = ""                 'Contient le message pour l'identifiant.
        Dim sMessageGeometrie As String = ""                'Contient le message pour la géométrie.
        Dim bSucces As Boolean = False                      'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        Try
            'Définir la géométrie par défaut
            TraiterDecoupageElement = New GeometryBag
            TraiterDecoupageElement.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDecoupageElement, IGeometryCollection)

            'Créer la requête spatiale
            pSpatialFilter = New SpatialFilterClass
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            'Définir la référence spatiale de sortie dans la requête spatiale
            pSpatialFilter.OutputSpatialReference(gpFeatureLayerDecoupage.FeatureClass.ShapeFieldName) = TraiterDecoupageElement.SpatialReference

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Initialisation du traitement
                bSucces = True
                sMessageAttribut = "#L'identifiant de découpage de l'élément est valide"
                sMessageGeometrie = "#La géométrie de l'élément est à l'intérieur du polygone de découpage"

                'Interface pour projeter
                pGeometry = pFeature.ShapeCopy
                pGeometry.Project(TraiterDecoupageElement.SpatialReference)

                'Vérifier si l'identifiant précédant est différent celui de l'élément traité
                If pFeature.Value(giAttribut).ToString <> sIdentifiant Then
                    'Extraire le centre d'une géométrie
                    pPoint = CentreGeometrie(pGeometry)

                    'Définir la géométrie utilisée pour la relation spatiale
                    pSpatialFilter.Geometry = pPoint

                    'Extraire les éléments en relation
                    oFeatureColl = ExtraireElementsRelation(pSpatialFilter, gpFeatureLayerDecoupage)

                    'Vérifier l'absence d'un élément de découpage
                    If oFeatureColl.Count = 0 Then
                        'Définir l'erreur
                        bSucces = False
                        sMessageAttribut = "#Aucun élément de découpage"
                        sMessageGeometrie = ""

                        'Plusieurs éléments de découpage
                    ElseIf oFeatureColl.Count > 1 Then
                        'Définir l'erreur
                        bSucces = False
                        sMessageAttribut = "#Plusieurs éléments de découpage"
                        sMessageGeometrie = ""

                        'Un élément de découpage
                    Else
                        'Définir l'élément de découpage
                        pFeatureDecoupage = CType(oFeatureColl.Item(1), IFeature)

                        'Définir la géométrie de découpage
                        pGeometryDecoupage = pFeatureDecoupage.ShapeCopy
                        pGeometryDecoupage.Project(TraiterDecoupageElement.SpatialReference)

                        'Interface pour valider la géométrie
                        pTopoOp = CType(pGeometryDecoupage, ITopologicalOperator2)
                        pRelOp = CType(pTopoOp.Buffer(gdPrecision), IRelationalOperator)

                        'Définir l'identifiant de découpage
                        sIdentifiant = pFeatureDecoupage.Value(giPosAttributDecoupage).ToString

                        'Définir l'identifiant de découpage précédant
                        sIdentifiantPrec = sIdentifiant
                    End If
                End If

                'Vérifier si un découpage est présent
                If oFeatureColl.Count = 1 Then
                    'Vérifier si l'identifiant est valide
                    If pFeature.Value(giAttribut).ToString <> sIdentifiant Then
                        'Définir l'erreur
                        bSucces = False
                        sMessageAttribut = "#L'identifiant de découpage de l'élément est invalide," & gsNomAttribut & "=" & pFeature.Value(giAttribut).ToString
                    End If

                    'Vérifier si la géométrie est valide
                    If Not pRelOp.Contains(pGeometry) Then
                        'Interface pour définir la géométrie en erreur
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        'Définir la géométrie en erreur
                        pGeometry = pTopoOp.Difference(pGeometryDecoupage)
                        'Vérifier si la géométrie n'est pas vide
                        If Not pGeometry.IsEmpty Then
                            'Définir l'erreur
                            bSucces = False
                            sMessageGeometrie = "#La géométrie de l'élément n'est pas à l'intérieur du polygone de découpage"
                        Else
                            'Rédéfinir la géométrie de l'élément
                            pGeometry = CType(pTopoOp, IGeometry)
                        End If
                    End If
                End If

                'Vérifier si on doit sélectionner l'élément
                If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                    'Ajouter l'élément dans la sélection
                    pFeatureSel.Add(pFeature)
                    'Ajouter l'enveloppe de l'élément sélectionné
                    pGeomSelColl.AddGeometry(pGeometry)
                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                                        & sMessageAttribut & sMessageGeometrie _
                                        & "/" & gsExpression & "=" & sIdentifiant _
                                        & "/Précision=" & gdPrecision.ToString, pGeometry)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pPoint = Nothing
            pSpatialFilter = Nothing
            oFeatureColl = Nothing
            pFeatureDecoupage = Nothing
            pGeometryDecoupage = Nothing
            pRelOp = Nothing
            pTopoOp = Nothing
        End Try
    End Function
#End Region
End Class
