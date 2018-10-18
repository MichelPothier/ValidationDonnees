Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt

'**
'Nom de la composante : clsDistanceGeometrie.vb
'
'''<summary>
''' Classe qui permet de traiter les distances d'une géométrie pour les FeatureClass de type Polyline et Polygon seulement.
''' 
''' La classe permet de traiter les deux attributs de distance de géométrie LATERALE et LONGITUDINALE.
''' 
''' LATERALE : La distance latérale (Douglass-Peuker) entre la droite des sommets non-consécutifs 
'''            et la perpendiculaire des sommets entre ces derniers sommets d'une géométrie.
'''            ATTENTION : Ce calcul peut tenir compte de la topologie des éléments en relations, mais seulement si elle est active.
''' LONGITUDINALE : La distance entre les sommets consécutifs d'une géométrie.
''' 
''' Note : Pour sélectionner les sommets en trop (^1\.).
'''        Pour sélectionner les sommets manquants (^\d\.).
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 20 avril 2015
'''</remarks>
''' 
Public Class clsDistanceGeometrie
    Inherits clsValeurAttribut

    '''<summary>Indique si la topologie doit être créé pour le traitement de la distance latérale.</summary>
    Protected gbTopologie As Boolean = False

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "LATERALE"
            Expression = "1.5"

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
                   ByVal sNomAttribut As String, ByVal sExpression As String)
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
        gbTopologie = Nothing
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
            Nom = "DistanceGeometrie"
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
            'Vérifier la présence du paramètre TOPOLOGIE
            If gbTopologie = True Then
                'Retourner aussi le masque spatial
                Parametres = Parametres & " TOPOLOGIE"
            End If
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

            'Extraire les paramètres
            params = value.Split(CChar(" "))

            'Vérifier si les deux paramètres sont présents
            If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION")
            'Définir le nom de l'attribut
            gsNomAttribut = params(0)
            'Définir l'expression
            gsExpression = params(1)
            'Aucune relation n'est utilisé
            gbTopologie = False
            gpFeatureLayersRelation = Nothing

            'Vérifier si les deux paramètres sont présents
            If params.Length > 2 Then
                'vérifier si le paramètre topolgie est présent
                If params(2) = "TOPOLOGIE" Then
                    'Indique si la topologie doit être créée
                    gbTopologie = True
                    'Les relations doit être présente
                    gpFeatureLayersRelation = New Collection
                Else
                    Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION et un optionel TOPOLOGIE")
                End If
            End If
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"

    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Public Overloads Overrides Function ListeParametres() As Collection
        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir le paramètre pour trouver les distances latérales minimum
                ListeParametres.Add("LATERALE 1.5")
                'Définir le paramètre pour trouver les distances latérales minimum avec la topologie
                ListeParametres.Add("LATERALE 1.5 TOPOLOGIE")
                'Définir le paramètre pour trouver les distances latérales
                ListeParametres.Add("LATERALE ^(0\.|1\.[0-5])")
                'Définir le paramètre pour trouver les distances longitudinales minimum
                ListeParametres.Add("LONGITUDINALE 3.0")
                'Définir le paramètre pour trouver les distances longitudinales minimum avec la topologie
                ListeParametres.Add("LONGITUDINALE 3.0 TOPOLOGIE")
                'Définir le paramètre pour trouver les distances longitudinales
                ListeParametres.Add("LONGITUDINALE ^[3-9]\.|\d\d\.")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
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
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'La contrainte est valide
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide"
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polyline ou Polygon."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'attribut est valide.
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function AttributValide() As Boolean
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            AttributValide = False
            gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut

            'Vérifier si l'attribut est valide
            If gsNomAttribut = "LATERALE" Or gsNomAttribut = "LONGITUDINALE" Then
                'La contrainte est valide
                AttributValide = True
                gsMessage = "La contrainte est valide"
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la distance dans les géométries respecte ou non l'expression spécifiée.
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
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.

        Try
            'Sortir si la contrainte est invalide
            If Me.EstValide() = False Then Err.Raise(1, , Me.Message)

            'Définir la géométrie par défaut
            Selectionner = New GeometryBag
            Selectionner.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Conserver le type de sélection à effectuer
            gbEnleverSelection = bEnleverSelection

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Vérifier si des éléments sont sélectionnés
            If pSelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pSelectionSet = pFeatureSel.SelectionSet
            End If

            'Enlever la sélection
            pFeatureSel.Clear()

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Si le nom de l'attribut est LATERALE
            If gsNomAttribut = "LATERALE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                'Vérifier si l'expression régulière est numérique
                If TestDBL(gsExpression) Then
                    'Vérifier si on doit créer la topologie
                    If gbTopologie Then
                        'Afficher le message d'identification des points trouvés
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                        'Vérifier si la Layer de sélection est absent dans les Layers relations
                        If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                            'Ajouter le layer de sélection dans les layers en relation
                            gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                        End If
                        'Interface pour extraire la tolérance de précision de la référence spatiale
                        pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                        'Création de la topologie
                        pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)
                    End If

                    'Vérifier la topologie est valide
                    If pTopologyGraph IsNot Nothing Then
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon le MapTopology (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer et retourne les sommets trouvés en tenant compte de la topologie des éléments en relation
                        Selectionner = TraiterDistanceLateraleTopologie(ConvertDBL(gsExpression), pTopologyGraph,
                                                                        pFeatureCursor, pTrackCancel, bEnleverSelection)
                    Else
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon le FeatureLayer (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer et retourne les sommets trouvés
                        Selectionner = TraiterDistanceLateraleMinimum(ConvertDBL(gsExpression), pFeatureCursor, pTrackCancel, bEnleverSelection)
                    End If

                    'Si l'expression régulière n'est pas numérique
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterDistanceLaterale(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est LONGITUDINALE
            ElseIf gsNomAttribut = "LONGITUDINALE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Vérifier si l'expression régulière est numérique
                If TestDBL(gsExpression) Then
                    'Vérifier si on doit créer la topologie
                    If gbTopologie Then
                        'Afficher le message d'identification des points trouvés
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                        'Vérifier si la Layer de sélection est absent dans les Layers relations
                        If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                            'Ajouter le layer de sélection dans les layers en relation
                            gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                        End If
                        'Interface pour extraire la tolérance de précision de la référence spatiale
                        pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                        'Création de la topologie
                        pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)
                    End If

                    'Vérifier la topologie est valide
                    If pTopologyGraph IsNot Nothing Then
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon le MapTopology (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer et retourne les droites trouvés en tenant compte de la topologie des éléments en relation
                        Selectionner = TraiterDistanceLongitudinaleTopologie(ConvertDBL(gsExpression), pTopologyGraph,
                                                                             pFeatureCursor, pTrackCancel, bEnleverSelection)
                    Else
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon le FeatureLayer (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer
                        Selectionner = TraiterDistanceLongitudinaleMinimum(ConvertDBL(gsExpression), pFeatureCursor, pTrackCancel, bEnleverSelection)
                    End If

                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterDistanceLongitudinale(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If
            End If

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
            pTopologyGraph = Nothing
            pSRTolerance = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont ses sommets se retrouvent à l'intérieur 
    ''' de la distance latérale (Douglass-Peuker) en tenant compte des éléments en relations contenus dans la topologie et le découpage.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="dDistance"> Distance latérale à vérifier.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des erreurs.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLateraleTopologie(ByVal dDistance As Double, ByRef pTopologyGraph As ITopologyGraph4,
                                                      ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                      Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge. 
        Dim pPolyline As IPolyline = Nothing                'Interface contenant un edge. 
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim pMultiPointErr As IMultipoint = Nothing         'Interface contenant les sommets qui se trouvent à l'intérieur de la distance.
        Dim dDistLateral As Double = 0                      'Contient la valeur de la distance latérale.
        Dim dDistLateralMin As Double = -1                  'Contient la valeur de la distance latérale minimum.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLateraleTopologie = New GeometryBag
            TraiterDistanceLateraleTopologie.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLateraleTopologie, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Interface pour projeter
                        pGeometry = pTopoEdge.Geometry
                        'Projeter la géométrie
                        pGeometry.Project(pSpatialRef)

                        'Interface pour extraire les segments
                        pPointColl = CType(pGeometry, IPointCollection)

                        'Définir la distance minimum par défaut
                        pPolyline = CType(pGeometry, IPolyline)
                        dDistLateralMin = pPolyline.Length

                        'Calculer la distance latérale
                        pMultiPointErr = DistanceLaterale(dDistance, pPointColl, 0, pPointColl.PointCount - 1, dDistLateralMin)

                        'Vérifier si on doit sélectionner l'élément
                        If (pMultiPointErr.IsEmpty And Not bEnleverSelection) Or (Not pMultiPointErr.IsEmpty And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Vérifier l'absence des points trouvés
                            If pMultiPointErr.IsEmpty Then
                                'Ajouter le centre de la géométrie transformer en multipoint
                                pMultiPointErr = GeometrieToMultiPoint(CentreGeometrie(pTopoEdge.Geometry))
                            End If
                            'Ajouter les points trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(pMultiPointErr)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                pMultiPointErr, CSng(dDistLateralMin))
                        End If

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Définir la géométrie du Edge
                        pGeometry = pTopoEdge.Geometry
                        'Interface pour extraire la limite de la géométrie
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        'Projeter la limite de découpage
                        pLimiteDecoupage.Project(pGeometry.SpatialReference)
                        'Enlever la partie commune avec la limite du polygone de découpage
                        pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                        'Projeter la géométrie
                        pGeometry.Project(pSpatialRef)

                        'Interface pour extraire les segments
                        pPointColl = CType(pGeometry, IPointCollection)

                        'Définir la distance minimum par défaut
                        pPolyline = CType(pGeometry, IPolyline)
                        dDistLateralMin = pPolyline.Length

                        'Calculer la distance latérale
                        pMultiPointErr = DistanceLaterale(dDistance, pPointColl, 0, pPointColl.PointCount - 1, dDistLateralMin)

                        'Vérifier si on doit sélectionner l'élément
                        If (pMultiPointErr.IsEmpty And Not bEnleverSelection) Or (Not pMultiPointErr.IsEmpty And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Vérifier l'absence des points trouvés
                            If pMultiPointErr.IsEmpty Then
                                'Ajouter le centre de la géométrie transformer en multipoint
                                pMultiPointErr = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                            End If
                            'Ajouter les points trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(pMultiPointErr)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                pMultiPointErr, CSng(dDistLateralMin))
                        End If

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pPolyline = Nothing
            pTopoOp = Nothing
            pLimiteDecoupage = Nothing
            pPointColl = Nothing
            pMultiPointErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont ses sommets se retrouvent à l'intérieur 
    ''' de la distance latérale (Douglass-Peuker) en tenant compte du découpage seulement.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="dDistance"> Distance latérale à vérifier.</param>
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLateraleMinimum(ByVal dDistance As Double, ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPath As IPath = Nothing                        'Interface contenant un Path. 
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim pMultiPointErr As IMultipoint = Nothing         'Interface contenant les sommets qui se trouvent à l'intérieur de la distance.
        Dim dDistLateral As Double = 0                      'Contient la valeur de la distance latérale.
        Dim dDistLateralMin As Double = -1                  'Contient la valeur de la distance latérale minimum.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLateraleMinimum = New GeometryBag
            TraiterDistanceLateraleMinimum.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLateraleMinimum, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface contenant la composante à traiter
                        pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                        'Interface pour définir la distance 
                        pPath = CType(pGeomColl.Geometry(i), IPath)
                        'Définir la distance minimum par défaut
                        dDistLateralMin = pPath.Length

                        'Calculer la distance latérale
                        pMultiPointErr = DistanceLaterale(dDistance, pPointColl, 0, pPointColl.PointCount - 1, dDistLateralMin)

                        'Vérifier si on doit sélectionner l'élément
                        If (pMultiPointErr.IsEmpty And Not bEnleverSelection) Or (Not pMultiPointErr.IsEmpty And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Vérifier l'absence des points trouvés
                            If pMultiPointErr.IsEmpty Then
                                'Ajouter le centre de la géométrie transformer en multipoint
                                pMultiPointErr = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                            End If
                            'Ajouter les points trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(pMultiPointErr)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                pMultiPointErr, CSng(dDistLateralMin))
                        End If
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Projeter la limite de découpage
                    pLimiteDecoupage.Project(pGeometry.SpatialReference)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface contenant la composante à traiter
                        pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                        'Interface pour définir la distance
                        pPath = CType(pGeomColl.Geometry(i), IPath)
                        'Définir la distance minimum par défaut
                        dDistLateralMin = pPath.Length

                        'Calculer la distance latérale
                        pMultiPointErr = DistanceLaterale(dDistance, pPointColl, 0, pPointColl.PointCount - 1, dDistLateralMin)

                        'Vérifier si on doit sélectionner l'élément
                        If (pMultiPointErr.IsEmpty And Not bEnleverSelection) Or (Not pMultiPointErr.IsEmpty And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Vérifier l'absence des points trouvés
                            If pMultiPointErr.IsEmpty Then
                                'Ajouter le centre de la géométrie transformer en multipoint
                                pMultiPointErr = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                            End If
                            'Ajouter les points trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(pMultiPointErr)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                pMultiPointErr, CSng(dDistLateralMin))
                        End If
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
            pLimiteDecoupage = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pPointColl = Nothing
            pMultiPointErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la distance latérale minimum trouvée (Douglass-Peuker) de la géométrie
    ''' respecte ou non l'expression régulière spécifiée.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLaterale(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim pPath As IPath = Nothing                        'Interface pour extraire les composante de la géométrie.
        Dim pMultiPoint As IMultipoint = Nothing            'Interface du centre de la géométrie.
        Dim dDistLateral As Double = 0                      'Contient la valeur de la distance latérale.
        Dim dDistLateralMin As Double = -1                  'Contient la valeur de la distance latérale minimu trouvée.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLaterale = New GeometryBag
            TraiterDistanceLaterale.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLaterale, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire le Path
                        pPath = CType(pGeomColl.Geometry(i), IPath)

                        'Interface pour extraire les sommets du Path
                        pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                        'Définir la distance latérale minimum par défaut
                        If dDistLateralMin = -1 Then dDistLateralMin = pPath.Length

                        'Calculer la distance latérale minimum de la géométrie
                        dDistLateral = DistanceLateraleMinimum(pPath, 0, pPointColl.PointCount - 1)

                        'Définir la distance latérale minimum
                        If dDistLateral < dDistLateralMin Then dDistLateralMin = dDistLateral

                        'Valider la valeur d'attribut selon l'expression régulière
                        oMatch = oRegEx.Match(dDistLateral.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                        End If
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter le centre de la géométrie transformer en multipoint
                        pMultiPoint = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                        'Ajouter le centre de la géométrie
                        pGeomSelColl.AddGeometry(pGeometry)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F3") & " /ExpReg=" & gsExpression, _
                                            pMultiPoint, CSng(dDistLateralMin))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Projeter la limite de découpage
                    pLimiteDecoupage.Project(pGeometry.SpatialReference)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire le Path
                        pPath = CType(pGeomColl.Geometry(i), IPath)

                        'Interface pour extraire les sommets du Path
                        pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                        'Définir la distance latérale minimum par défaut
                        If dDistLateralMin = -1 Then dDistLateralMin = pPath.Length

                        'Calculer la distance latérale minimum de la géométrie
                        dDistLateral = DistanceLateraleMinimum(pPath, 0, pPointColl.PointCount - 1)

                        'Définir la distance latérale minimum
                        If dDistLateral < dDistLateralMin Then dDistLateralMin = dDistLateral

                        'Valider la valeur d'attribut selon l'expression régulière
                        oMatch = oRegEx.Match(dDistLateral.ToString("F1"))

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                        End If
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter le centre de la géométrie transformer en multipoint
                        pMultiPoint = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                        'Ajouter le centre de la géométrie
                        pGeomSelColl.AddGeometry(pGeometry)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Distance=" & dDistLateralMin.ToString("F1") & " /ExpReg=" & gsExpression, _
                                            pMultiPoint, CSng(dDistLateralMin))
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
            pLimiteDecoupage = Nothing
            pPointColl = Nothing
            pMultiPoint = Nothing
            pPath = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des droites se retrouvent à l'intérieur 
    ''' de la distance longitudinale en tenant compte des éléments en relations contenus dans la topologie et le découpage.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="dDistance"> Distance longitudinale à vérifier.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLongitudinaleTopologie(ByVal dDistance As Double, ByRef pTopologyGraph As ITopologyGraph4, ByRef pFeatureCursor As IFeatureCursor,
                                                           ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge. 
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire la longueur d'un segment de base.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant les lignes trouvées.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression demandée est un succès.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLongitudinaleTopologie = New GeometryBag
            TraiterDistanceLongitudinaleTopologie.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLongitudinaleTopologie, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Interface pour projeter
                        pGeometry = pTopoEdge.Geometry
                        'Projeter la géométrie
                        pGeometry.Project(pSpatialRef)

                        'Interface pour extraire les composantes
                        pGeomColl = CType(pGeometry, IGeometryCollection)

                        'Traiter toutes les composantes
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Interface pour extraire les segments
                            pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                            'Traiter tous les angles entre les segments consécutifs
                            For j = 0 To pSegColl.SegmentCount - 1
                                'Définir le segment de base
                                pLineBase = CType(pSegColl.Segment(j), ILine)

                                'Valider la distance
                                bSucces = pLineBase.Length >= dDistance

                                'Vérifier si on doit sélectionner l'élément
                                If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                    'Initialiser l'ajout
                                    bAjouter = True
                                    'Définir la ligne en erreur
                                    pPolylineErr = LineToPolyline(pLineBase)
                                    'Ajouter les sommets de l'élément sélectionné
                                    pGeomSelColl.AddGeometry(pPolylineErr)
                                    'Écrire une erreur
                                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                        pPolylineErr, CSng(pLineBase.Length))
                                End If
                            Next j
                        Next i

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Vérifier si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Interface pour projeter
                        pGeometry = pTopoEdge.Geometry
                        'Interface pour extraire la limite de la géométrie
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        'Projeter la limite de découpage
                        pLimiteDecoupage.Project(pGeometry.SpatialReference)
                        'Enlever la partie commune avec la limite du polygone de découpage
                        pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                        'Projeter la géométrie
                        pGeometry.Project(pSpatialRef)

                        'Interface pour extraire les composantes
                        pGeomColl = CType(pGeometry, IGeometryCollection)

                        'Traiter toutes les composantes
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Interface pour extraire les segments
                            pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                            'Traiter tous les angles entre les segments consécutifs
                            For j = 0 To pSegColl.SegmentCount - 1
                                'Définir le segment de base
                                pLineBase = CType(pSegColl.Segment(j), ILine)

                                'Valider la distance
                                bSucces = pLineBase.Length >= dDistance

                                'Vérifier si on doit sélectionner l'élément
                                If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                    'Initialiser l'ajout
                                    bAjouter = True
                                    'Définir la ligne en erreur
                                    pPolylineErr = LineToPolyline(pLineBase)
                                    'Ajouter les sommets de l'élément sélectionné
                                    pGeomSelColl.AddGeometry(pPolylineErr)
                                    'Écrire une erreur
                                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                        pPolylineErr, CSng(pLineBase.Length))
                                End If
                            Next j
                        Next i

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pLimiteDecoupage = Nothing
            pTopoOp = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pPolylineErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des droites se retrouvent à l'intérieur 
    ''' de la distance longitudinale en tenant compte du découpage seulement.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="dDistance"> Distance longitudinale à vérifier.</param>
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLongitudinaleMinimum(ByVal dDistance As Double, ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                         Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire la longueur d'un segment de base.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant les lignes trouvées.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression demandée est un succès.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLongitudinaleMinimum = New GeometryBag
            TraiterDistanceLongitudinaleMinimum.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLongitudinaleMinimum, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 1
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)

                            'Valider la distance
                            bSucces = pLineBase.Length >= dDistance

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Définir la ligne en erreur
                                pPolylineErr = LineToPolyline(pLineBase)
                                'Ajouter les sommets de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pPolylineErr)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                    pPolylineErr, CSng(pLineBase.Length))
                            End If
                        Next j
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Vérifier si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Projeter la limite de découpage
                    pLimiteDecoupage.Project(pGeometry.SpatialReference)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 1
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)

                            'Valider la distance
                            bSucces = pLineBase.Length >= dDistance

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Définir la ligne en erreur
                                pPolylineErr = LineToPolyline(pLineBase)
                                'Ajouter les sommets de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pPolylineErr)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /DistMin=" & dDistance.ToString, _
                                                    pPolylineErr, CSng(pLineBase.Length))
                            End If
                        Next j
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pLimiteDecoupage = Nothing
            pTopoOp = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pPolylineErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la distance longitudinale de la géométrie
    ''' respecte ou non l'expression régulière spécifiée.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    Private Function TraiterDistanceLongitudinale(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                  Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pSpatialRef As ISpatialReference = Nothing              'Interface contenant la référence spatiale de projection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire la longueur d'un segment de base.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant les lignes trouvées.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la référence spatiale de projection
            pSpatialRef = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir la géométrie par défaut
            TraiterDistanceLongitudinale = New GeometryBag
            TraiterDistanceLongitudinale.SpatialReference = pSpatialRef

            'Définir la limite du polygone de découpage
            pLimiteDecoupage = LimiteDecoupage
            'Vérifier si la limite est absente
            If pLimiteDecoupage Is Nothing Then
                'Créer une limite vide
                pLimiteDecoupage = New Polyline
                pLimiteDecoupage.SpatialReference = pSpatialRef
            End If

            'Vérifier si la référence est en géographique
            If TypeOf (pSpatialRef) Is GeographicCoordinateSystem Then
                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment
                'Définir la référence spatiale LCC NAD83 CSRS:3979
                pSpatialRef = pSpatialRefFact.CreateSpatialReference(3979)
                'Définir la tolérance XY
                pSpatialRefTol = CType(pSpatialRef, ISpatialReferenceTolerance)
                pSpatialRefRes = CType(pSpatialRef, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                pSpatialRefTol.XYTolerance = 0.001
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterDistanceLongitudinale, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 1
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)

                            'Valider la valeur d'attribut selon l'expression régulière
                            oMatch = oRegEx.Match(pLineBase.Length.ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Définir la ligne en erreur
                                pPolylineErr = LineToPolyline(pLineBase)
                                'Ajouter les sommets de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pPolylineErr)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                    pPolylineErr, CSng(pLineBase.Length))
                            End If
                        Next j
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Vérifier si la géométrie est de type Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Projeter la limite de découpage
                    pLimiteDecoupage.Project(pGeometry.SpatialReference)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(pSpatialRef)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 1
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)

                            'Valider la valeur d'attribut selon l'expression régulière
                            oMatch = oRegEx.Match(pLineBase.Length.ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Définir la ligne en erreur
                                pPolylineErr = LineToPolyline(pLineBase)
                                'Ajouter les sommets de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pPolylineErr)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pLineBase.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                    pPolylineErr, CSng(pLineBase.Length))
                            End If
                        Next j
                    Next i

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
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
            pSpatialRef = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pLimiteDecoupage = Nothing
            pTopoOp = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pPolylineErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de trouver la distance latérale minimum pour une composante d'élément.
    ''' 
    '''<param name="pPointColl"> Interface utilisé pour traiter les sommets de la composante.</param>
    '''<param name="iDebut"> Numéro du sommet de début à traiter.</param>
    '''<param name="iFin"> Numéro du sommet de fin à traiter.</param>
    ''' 
    '''<return> La distance minimum des distances maximum trouvées.</return>
    ''' 
    '''</summary>
    '''
    Private Function DistanceLateraleMinimum(ByRef pPath As IPath, ByVal iDebut As Integer, ByVal iFin As Integer) As Double
        'Déclarer les variables de travail
        Dim iNoeud As Integer = 0                       'Position du sommet le plus éloigné constituant un noeud.
        Dim dMax As Double = -1                         'Distance maximum trouvée.
        Dim dMin As Double = 0                          'Distance minimum des distances maximum trouvées.
        Dim dDist As Double = 0                         'Distance entre un sommet et la ligne de base.
        Dim pLine As ILine = New Line                   'Interface contenant la ligne de base.
        Dim pProxOp As IProximityOperator = Nothing     'Interface utilisé pour calculer la distance.
        Dim pPointColl As IPointCollection = Nothing    'Interface utilisé pour traiter les sommets de la composante.

        Try
            'Définir la référence spatiale
            pLine.SpatialReference = pPath.SpatialReference

            'Interface pour extraire les sommets de la composante
            pPointColl = CType(pPath, IPointCollection)

            'Construire la ligne de base
            pLine.PutCoords(pPointColl.Point(iDebut), pPointColl.Point(iFin))

            'Interface utilisé pour traiter les sommets de la composante.
            pPointColl = CType(pPath, IPointCollection)

            'Interface pour extraire la distance
            pProxOp = CType(pLine, IProximityOperator)

            'Traiter tous les sommets
            For i = iDebut + 1 To iFin - 1
                'Calculer la distance
                dDist = pProxOp.ReturnDistance(pPointColl.Point(i))

                'Vérifier si la distance trouvée est la plus grande trouvée
                If dDist > dMax Then
                    'Conserver le noeud
                    iNoeud = i
                    'Conserver la distance maximum
                    dMax = dDist
                End If
            Next i

            'Vérifier si un noeud a été trouvé 
            If iNoeud > 0 Then
                'Traiter la partie avant le noeud
                dMin = DistanceLateraleMinimum(pPath, iDebut, iNoeud)
                'Vérifier si une distance minimum est trouvée 
                If dMin >= 0 Then
                    If dMin < dMax Then
                        dMax = dMin
                    End If
                End If

                'Traiter la partie aprés le noeud
                dMin = DistanceLateraleMinimum(pPath, iNoeud, iFin)
                'Vérifier si une distance minimum est trouvée 
                If dMin >= 0 Then
                    If dMin < dMax Then
                        dMax = dMin
                    End If
                End If
            End If

            'Retourner la distance minimum des distances maximun
            DistanceLateraleMinimum = dMax

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pLine = Nothing
            pProxOp = Nothing
            pPointColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de trouver les sommets qui se trouvent à l'intérieur de la distance latérale pour une composante d'élément.
    ''' 
    '''<param name="dDistance"> Distance latérale à vérifier.</param>
    '''<param name="pPointColl"> Interface utilisé pour traiter les sommets de la composante.</param>
    '''<param name="iDebut"> Numéro du sommet de début à traiter.</param>
    '''<param name="iFin"> Numéro du sommet de fin à traiter.</param>
    ''' 
    '''<return> Interface contenant les sommets qui se trouvent à l'intérieur de la distance.</return>
    ''' 
    '''</summary>
    '''
    Private Function DistanceLaterale(ByVal dDistance As Double, ByRef pPointColl As IPointCollection, _
                                      ByVal iDebut As Integer, ByVal iFin As Integer, ByRef dDistanceMin As Double) As IMultipoint
        'Déclarer les variables de travail
        Dim iNoeud As Integer = 0                       'Position du sommet le plus éloigné constituant un noeud.
        Dim dMax As Double = -1                         'Distance maximum trouvée.
        Dim dDist As Double = 0                         'Distance entre un sommet et la ligne de base.
        Dim pLine As ILine = New Line                   'Interface contenant la ligne de base.
        Dim pProxOp As IProximityOperator = Nothing     'Interface utilisé pour calculer la distance.
        Dim pMultiPoint As IPointCollection = Nothing   'Interface utilisé pour ajouter les sommets trouvés.

        Try
            'Définir la valeur par défaut vide 
            DistanceLaterale = New Multipoint
            pMultiPoint = CType(DistanceLaterale, IPointCollection)

            'On sort si aucun sommet
            If pPointColl.PointCount = 0 Then Return DistanceLaterale

            'Définir la référence spatiale
            DistanceLaterale.SpatialReference = pPointColl.Point(iDebut).SpatialReference
            pLine.SpatialReference = DistanceLaterale.SpatialReference

            'Construire la ligne de base
            pLine.PutCoords(pPointColl.Point(iDebut), pPointColl.Point(iFin))

            'Interface pour extraire la distance
            pProxOp = CType(pLine, IProximityOperator)

            'Traiter tous les sommets
            For i = iDebut + 1 To iFin - 1
                'Calculer la distance
                dDist = pProxOp.ReturnDistance(pPointColl.Point(i))

                'Vérifier si la distance trouvée est la plus grande trouvée
                If dDist > dMax Then
                    'Conserver le noeud
                    iNoeud = i
                    'Conserver la distance maximum
                    dMax = dDist
                End If
            Next i
            'Debug.Print(iDebut.ToString & "," & iFin.ToString & "," & iNoeud.ToString & "," & dMax.ToString)

            'Vérifier si la distance maximale trouvée est plus grande que la distance latérale spécifiée 
            If dMax > dDistance Then
                'Traiter la partie avant le noeud
                pMultiPoint.AddPointCollection(CType(DistanceLaterale(dDistance, pPointColl, iDebut, iNoeud, dDistanceMin), IPointCollection))
                'Traiter la partie aprés le noeud
                pMultiPoint.AddPointCollection(CType(DistanceLaterale(dDistance, pPointColl, iNoeud, iFin, dDistanceMin), IPointCollection))

                'Si des sommets sont trouvés à l'intérieurs de la distance latérale
            ElseIf dMax >= 0 Then
                'Définir la distance minimum trouvée
                If dMax < dDistanceMin Then dDistanceMin = dMax
                'Traiter tous les sommets qui sont à l'intérieurs de la distance latérale
                For i = iDebut + 1 To iFin - 1
                    'Ajouter le sommet 
                    pMultiPoint.AddPoint(pPointColl.Point(i))
                Next i
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pLine = Nothing
            pProxOp = Nothing
            pMultiPoint = Nothing
        End Try
    End Function
#End Region
End Class
