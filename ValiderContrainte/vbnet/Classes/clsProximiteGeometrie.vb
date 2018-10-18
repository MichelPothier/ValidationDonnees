Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.EditorExt

'**
'Nom de la composante : clsProximiteGeometrie.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont la proximité des géométries
''' avec des éléments en relation respecte ou non la tolérance spécifiée.
''' 
''' Chaque géométrie des éléments traités est comparée avec ses éléments en relations.
''' 
''' La classe permet de traiter les quatres attributs de proximité POINT, LIGNE, TOPOLOGIE et LIGNE_INTERNE.
''' 
''' TOPOLOGIE : La proximité de type topologie entre deux géométries.
''' POINT : La proximité de type point entre deux géométries.
''' LIGNE : La proximité de type ligne entre deux géométries.
''' LIGNE_INTERNE : La proximité de type ligne interne des géométries traitées.
''' 
''' Note : Le traitement de proximité de point s'exécute seulement sur des éléments de type point, multipoint ou polyline.
'''        Le traitement de proximité de ligne s'exécute seulement sur des éléments de type polyline ou polygon.
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 2 juin 2015
'''</remarks>
''' 
Public Class clsProximiteGeometrie
    Inherits clsValeurAttribut

    '''<summary>Tolérance de proximité.</summary>
    Protected gdTolerance As Double = 2.0
    '''<summary>Longueur de proximité de ligne.</summary>
    Protected gdLongueur As Double = 10.0

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "POINT"
            Expression = "2.0"
            gpFeatureLayersRelation = New Collection

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
            gpFeatureLayersRelation = New Collection

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
            gpFeatureLayersRelation = New Collection

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
            gpMap = pMap
            gpFeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gdTolerance = Nothing
        gdLongueur = Nothing
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
            Nom = "ProximiteGeometrie"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la tolérance de proximité.
    '''</summary>
    ''' 
    Public Property Tolerance() As Double
        Get
            Tolerance = gdTolerance
        End Get
        Set(ByVal value As Double)
            gdTolerance = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la longueur de proximité de ligne.
    '''</summary>
    ''' 
    Public Property Longueur() As Double
        Get
            Longueur = gdLongueur
        End Get
        Set(ByVal value As Double)
            gdLongueur = value
        End Set
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
                'Si l'attribut est LIGNE_INTERNE
                If gsNomAttribut = "LIGNE_INTERNE" Then
                    'Aucun FeatureLayer en relation n'est requis
                    gpFeatureLayersRelation = Nothing
                Else
                    'Au moins un FeatureLayer en relation requis
                    gpFeatureLayersRelation = New Collection
                End If

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
        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir le paramètre pour trouver les proximités de topologie
                ListeParametres.Add("TOPOLOGIE 2.0")
                'Définir le paramètre pour trouver  les proximités de point
                ListeParametres.Add("POINT 2.0")
                'Définir le paramètre pour trouver les proximités de ligne avec les éléments en relation
                ListeParametres.Add("LIGNE 2.0,20.0")
                'Définir le paramètre pour trouver les proximités de ligne INTERNE
                'ListeParametres.Add("LIGNE_INTERNE 2.0,20.0")
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
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Vérifier si l'attribut est POINT
                    If gsNomAttribut = "POINT" _
                    And gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Point, Multipoint ou Polyline."

                        'Vérifier si l'attribut est LIGNE
                    ElseIf (gsNomAttribut = "LIGNE" Or gsNomAttribut = "LIGNE_INTERNE") _
                    And (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) Then
                        gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polyline ou Polygon."

                        'La Featureclass est valide
                    Else
                        'La contrainte est valide
                        FeatureClassValide = True
                        gsMessage = "La contrainte est valide."
                    End If
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Point, MultiPoint, Polyline ou Polygon."
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
            If gsNomAttribut = "TOPOLOGIE" Or gsNomAttribut = "POINT" Or gsNomAttribut = "LIGNE" Or gsNomAttribut = "LIGNE_INTERNE" Then
                'La contrainte est valide
                AttributValide = True
                gsMessage = "La contrainte est valide."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'expression contenant la tolérance de proximité est valide.
    ''' 
    '''<return>Boolean qui indique si l'expression est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function ExpressionValide() As Boolean
        'Déclarer les variables de travail
        Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression

        Try
            'Par défaut, l'expression est invalide
            ExpressionValide = False
            gsMessage = "ERREUR : L'expression est invalide :" & gsExpression

            'Vérifier si l'expression est numérique
            If IsNumeric(gsExpression) Then
                'Définir la tolérance
                gdTolerance = ConvertDBL(gsExpression)
                'vérifier si la tolérance est supérieur à 0
                If gdTolerance > 0 Then
                    'Retourner l'expression valide
                    ExpressionValide = True
                    gsMessage = "La contrainte est valide"
                Else
                    gsMessage = "ERREUR : La tolérance de proximité est inférieure à zéro :" & gsExpression
                End If

                'Si l'attribut est LIGNE
            ElseIf gsNomAttribut = "LIGNE" Or gsNomAttribut = "LIGNE_INTERNE" Then
                'Extraire les paramètres
                params = gsExpression.Split(CChar(","))

                'Vérifier si les deux paramètres sont présents, tolérance et longueur
                If params.Length = 2 Then
                    'Vérifier si le paramètre n'est pas numérique
                    If TestDBL(params(0)) Then
                        'Définir la tolérance
                        gdTolerance = ConvertDBL(params(0))
                        'Vérifier si la tolérance est supérieur à 0
                        If gdTolerance > 0 Then
                            'Vérifier si le paramètre n'est pas numérique
                            If TestDBL(params(1)) Then
                                'Définir la longueur
                                gdLongueur = ConvertDBL(params(1))
                                'Vérifier si la longueur est supérieur à la tolérance
                                If gdLongueur > gdTolerance Then
                                    'Retourner l'expression valide
                                    ExpressionValide = True
                                    gsMessage = "La contrainte est valide"
                                Else
                                    gsMessage = "ERREUR : La longueur de proximité est inférieure à la tolérance :" & params(1)
                                End If
                            Else
                                gsMessage = "ERREUR : La longueur de proximité n'est pas numérique :" & params(0)
                            End If
                        Else
                            gsMessage = "ERREUR : La tolérance de proximité est inférieure à zéro :" & gsExpression
                        End If
                    Else
                        gsMessage = "ERREUR : La tolérance de proximité n'est pas numérique :" & params(0)
                    End If

                Else
                    gsMessage = "ERREUR : La tolérance de proximité n'est pas numérique :" & gsExpression
                End If

                'Si l'expression n'est pas numérique
            Else
                gsMessage = "ERREUR : L'expression n'est pas numérique :" & gsExpression
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            params = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection entre deux géométries
    ''' respecte ou non l'expression régulière spécifiée.
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

            'Vérifier si l'attribut est TOPOLOGIE
            If gsNomAttribut = "TOPOLOGIE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Traiter le FeatureLayer
                Selectionner = TraiterProximiteTopologie(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est POINT
            ElseIf gsNomAttribut = "POINT" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                'Traiter le FeatureLayer
                Selectionner = TraiterProximitePoint(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est LIGNE
            ElseIf gsNomAttribut = "LIGNE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Traiter le FeatureLayer
                Selectionner = TraiterProximiteLigne(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est LIGNE
            ElseIf gsNomAttribut = "LIGNE_INTERNE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Traiter le FeatureLayer
                Selectionner = TraiterProximiteLigneInterne(pTrackCancel, bEnleverSelection)
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
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la proximité de point respecte ou non 
    ''' la tolérance spécifiée.
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
    '''<return>Les géométries de proximité entre les éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    ''' 
    Private Function TraiterProximiteTopologie(ByRef pTrackCancel As ITrackCancel,
                                               Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour extraire la différence.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeometryTopo As IGeometry = Nothing            'Interface contenant la géométrie de la topologie.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour traiter toutes les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire le nombre de sommets.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.

        Try
            'Définir la géométrie par défaut
            TraiterProximiteTopologie = New GeometryBag
            TraiterProximiteTopologie.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterProximiteTopologie, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Définir une nouvelle sélection Vide
            gpMap.ClearSelection()
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de création de la topolgie des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Création de la topologie des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Vérifier si la Layer de sélection est absent dans les Layers relations
            If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                'Ajouter le layer de sélection dans les layers en relation
                gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
            End If
            'Création de la topologie
            pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, TraiterProximiteTopologie.SpatialReference), gpFeatureLayersRelation, gdTolerance)

            'Vérifier si la topologie est présente
            If pTopologyGraph IsNot Nothing Then
                'Afficher le message de sélection des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments de proximité (" & gpFeatureLayerSelection.Name & ") ..."

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, True, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'définir la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(TraiterProximiteTopologie.SpatialReference)

                    'Extraire la géométrie de la topologie
                    pGeometryTopo = pTopologyGraph.GetParentGeometry(gpFeatureLayerSelection.FeatureClass, pFeature.OID)
                    'Si la géométrie est absente
                    If pGeometryTopo Is Nothing Then
                        'Définir la géométrie de l'élément par défaut
                        pGeometryTopo = pGeometry
                        'Si la géométrie est présente
                    Else
                        'Extraire la différence entre la géométrie de la topologie de l'élément
                        pTopoOp = CType(pGeometryTopo, ITopologicalOperator2)
                        pGeometryTopo = pTopoOp.Difference(pGeometry)
                    End If

                    'Vérifier si on doit sélectionner l'élément
                    If (pGeometryTopo.IsEmpty And Not bEnleverSelection) Or (Not pGeometryTopo.IsEmpty And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)

                        'Vérifier si la géométrie est vide (sans différence)
                        If pGeometryTopo.IsEmpty Then
                            'Interface pour extraire le nombre de sommet
                            pPointColl = CType(pGeometry, IPointCollection)
                            'Ajouter l'enveloppe de l'élément sélectionné
                            pGeomResColl.AddGeometry(pGeometry)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                            & " #Géométrie identique /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", NbSommets=" & pPointColl.PointCount.ToString, _
                            pGeometry, pPointColl.PointCount)

                            'Si la géométrie n'est pas vide (avec différence)
                        Else
                            'Interface pour traiter toutes les composantes
                            pGeomColl = CType(pGeometryTopo, IGeometryCollection)
                            'Vérifier si plusieurs composantes sont présentes
                            If pGeomColl.GeometryCount > 1 Then
                                'Traiter toutes les composantes
                                For i = 0 To pGeomColl.GeometryCount - 1
                                    'Si  la géométrie est un Path
                                    If pGeomColl.Geometry(i).GeometryType = esriGeometryType.esriGeometryPath Then
                                        'Créer une Polyline
                                        pGeometry = PathToPolyline(CType(pGeomColl.Geometry(i), IPath))
                                        'Si  la géométrie est un Ring
                                    ElseIf pGeomColl.Geometry(i).GeometryType = esriGeometryType.esriGeometryRing Then
                                        'Créer un Polyfon
                                        pGeometry = RingToPolygon(CType(pGeomColl.Geometry(i), IRing), New Polygon)
                                        'Si  la géométrie est un Point
                                    ElseIf pGeomColl.Geometry(i).GeometryType = esriGeometryType.esriGeometryPoint Then
                                        'Créer un multipoint
                                        pGeometry = GeometrieToMultiPoint(pGeomColl.Geometry(i))
                                        'Sinon
                                    Else
                                        'Définir la géométrie
                                        pGeometry = pGeometryTopo
                                    End If
                                    'Interface pour extraire le nombre de sommet
                                    pPointColl = CType(pGeometry, IPointCollection)
                                    'Ajouter l'élément sélectionné
                                    pGeomResColl.AddGeometry(pGeometry)
                                    'Écrire une erreur
                                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                                    & " #Géométrie différente /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", NbSommets=" & pPointColl.PointCount.ToString, _
                                    pGeometry, pPointColl.PointCount)
                                Next

                                'Si une seule composante est présente
                            Else
                                'Définir la géométrie
                                pGeometry = pGeometryTopo
                                'Interface pour extraire le nombre de sommet
                                pPointColl = CType(pGeometry, IPointCollection)
                                'Ajouter  l'élément sélectionné
                                pGeomResColl.AddGeometry(pGeometry)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                                & " #Géométrie différente /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", NbSommets=" & pPointColl.PointCount.ToString, _
                                pGeometry, pPointColl.PointCount)
                            End If
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Définir les éléments sélectionnés
                pFeatureSel.SelectionSet = pNewSelectionSet

                'Définir le résultat des géométries sélectionnées
                TraiterProximiteTopologie = CType(pGeomResColl, IGeometryBag)
            Else
                'Retourner l'erreur
                Err.Raise(-1, , "Erreur : La topologie est invalide !")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeomResColl = Nothing
            pFeatureLayerRel = Nothing
            pTopologyGraph = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pGeometryTopo = Nothing
            pNewSelectionSet = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la proximité de point respecte ou non la tolérance spécifiée.
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
    '''<return>Les géométries de proximité entre les éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    ''' 
    Private Function TraiterProximitePoint(ByRef pTrackCancel As ITrackCancel,
                                           Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométrie des éléments trouvés.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant la relation spatiale de base.

        Try
            'Définir la géométrie par défaut
            TraiterProximitePoint = New GeometryBag
            TraiterProximitePoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterProximitePoint, IGeometryCollection)

            'Interface pour conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver la sélection de départ
            pSelectionSet = pFeatureSel.SelectionSet

            'Vider la sélection
            gpMap.ClearSelection()
            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Proximité de point des géométries (" & gpFeatureLayerSelection.Name & ") ..."

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Créer la requête spatiale
            pSpatialFilter = New SpatialFilterClass
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Traiter la proximité de point de l'élément
                ElementProximitePoint(pFeature, pSpatialFilter, pNewSelectionSet, pGeomResColl, bEnleverSelection)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeature = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter la proximité de point d'un élément.
    '''</summary>
    '''
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pSpatialFilter">Interface contenant la requête spatial pour trouver les éléments en relation.</param>
    '''<param name="pNewSelectionSet">'Interface pour sélectionner les éléments.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométrie des éléments trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    Private Sub ElementProximitePoint(ByRef pFeature As IFeature, ByRef pSpatialFilter As ISpatialFilter, _
                                      ByRef pNewSelectionSet As ISelectionSet, ByRef pGeomResColl As IGeometryCollection, ByVal bEnleverSelection As Boolean)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets.
        Dim pMultiPoint As IPointCollection = Nothing       'Interface contenant les points trouvés.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour créer un Buffer.
        Dim pProxOp As IProximityOperator = Nothing         'Interface pour extraire la distance de proximité.
        Dim dDistance As Double = 0                         'Contient la distance de proximité.
        Dim sMessage As String = ""                         'Contient le message.

        Try
            'Initialiser le message de retour
            sMessage = ""

            'Définir le multipoint vide par défaut
            pGeometry = New Multipoint
            pGeometry.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour ajouter les points
            pMultiPoint = CType(pGeometry, IPointCollection)

            'Définir la géométrie à traiter
            pGeometry = pFeature.ShapeCopy
            'Projeter la géométrie à traiter
            pGeometry.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)

            'Vérifier si la géométrie est une Polyline
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Interface pour extraire la limite
                pTopoOp = CType(pGeometry, ITopologicalOperator)
                'Interface contenant la limite de la géométrie
                pGeometry = pTopoOp.Boundary
            End If

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Distance minimum entre un sommet d'un élément et ses éléments en relation
                dDistance = DistanceMinimumPointRelation(CType(pGeometry, IPoint), pFeature, pSpatialFilter)

                'Vérifier si la distance est plus grande que la précision
                If dDistance > gdPrecision Then
                    'Ajouter le point dans le multipoint
                    pMultiPoint.AddPoint(CType(pGeometry, IPoint))

                    'Définir le message
                    sMessage = "(" & dDistance.ToString("F3") & ")"
                End If

                'Traiter un multipoint
            Else
                'Interface pour extraire les sommets
                pPointColl = CType(pGeometry, IPointCollection)

                'Traiter tous les sommets avec ses éléments en relation
                For j = 0 To pPointColl.PointCount - 1
                    'Distance minimum entre un sommet d'un élément et ses éléments en relation
                    dDistance = DistanceMinimumPointRelation(pPointColl.Point(j), pFeature, pSpatialFilter)

                    'Vérifier si le point ne connecte pas
                    If dDistance > gdPrecision Then
                        'Indiquer que le point ne connecte pas
                        pPointColl.Point(j).M = -1

                        'Vérifier si le point peut connecter
                        If dDistance < gdTolerance Then
                            'Ajouter le point dans le multipoint
                            pMultiPoint.AddPoint(pPointColl.Point(j))
                            'Définir le message
                            sMessage = sMessage & "(" & j.ToString & "=" & dDistance.ToString("F3") & ")"
                        End If
                    End If
                Next

                'Vérifier la proximité des points entre eux (interne)
                For j = 0 To pPointColl.PointCount - 2
                    'Si le point n'est pas connecté
                    If pPointColl.Point(j).M = -1 Then
                        'Interface pour extraire la distance
                        pProxOp = CType(pPointColl.Point(j), IProximityOperator)

                        'Traiter les autres sommets
                        For k = j + 1 To pPointColl.PointCount - 1
                            'Extraire la distance
                            dDistance = pProxOp.ReturnDistance(pPointColl.Point(k))

                            'Vérifier si la distance est plus grande que la précision
                            If dDistance > gdPrecision And dDistance < gdTolerance Then
                                'Ajouter le point dans le multipoint
                                pMultiPoint.AddPoint(pPointColl.Point(j))

                                'Définir le message
                                sMessage = sMessage & "(" & j.ToString & ":" & k.ToString & "=" & dDistance.ToString("F3") & ")"
                            End If
                        Next
                    End If
                Next
            End If

            'Vérifier si on doit sélectionner l'élément
            If (pMultiPoint.PointCount = 0 And Not bEnleverSelection) Or (pMultiPoint.PointCount > 0 And bEnleverSelection) Then
                'Ajouter l'élément dans la sélection
                pNewSelectionSet.Add(pFeature.OID)

                'Vérifier si le multipoint n'est pas vide
                If pMultiPoint.PointCount > 0 Then
                    'Définir la géométrie
                    pGeometry = CType(pMultiPoint, IGeometry)
                Else
                    'Définir le message
                    sMessage = "Aucune proximité de point"
                End If

                'Définir la géométrie de l'élément
                pGeomResColl.AddGeometry(pGeometry)

                'Écrire une erreur
                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage & " /" & gdTolerance.ToString("F1") & ", " & gdPrecision.ToString, pGeometry)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pPointColl = Nothing
            pMultiPoint = Nothing
            pTopoOp = Nothing
            pProxOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la distance minimum entre un sommet d'un élément et ses éléments en relation.
    '''</summary>
    '''
    '''<param name="pPoint">Interface contenant le point de vérification de la proximité.</param>
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pSpatialFilter">Interface contenant la requête spatial pour trouver les éléments en relation.</param>
    ''' 
    '''<returns>Distance minimum entre un sommet d'un élément et ses éléments en relation.</returns>
    '''
    Private Function DistanceMinimumPointRelation(ByRef pPoint As IPoint, ByRef pFeature As IFeature, ByRef pSpatialFilter As ISpatialFilter) As Double
        'Déclarer les variables de travail
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim oFeatureColl As Collection = Nothing            'Objet contenant la collection des éléments en relation.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant un élément en relation.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie en relation.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour créer un Buffer.
        Dim pProxOp As IProximityOperator = Nothing         'Interface pour extraire la distance.
        Dim dDistanceMin As Double = gdTolerance            'Contient la tolérance minimum trouvée

        'Définir la valeur de retour par défaut
        DistanceMinimumPointRelation = gdTolerance

        Try
            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Définir la référence spatiale de sortie dans la requête spatiale
                pSpatialFilter.OutputSpatialReference(pFeatureLayerRel.FeatureClass.ShapeFieldName) = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

                'Interface pour créer un buffer
                pTopoOp = CType(pPoint, ITopologicalOperator)
                'Définir la géométrie utilisée pour la relation spatiale
                pSpatialFilter.Geometry = pTopoOp.Buffer(gdTolerance)

                'Extraire les éléments en relation
                oFeatureColl = ExtraireElementsRelation(pSpatialFilter, pFeatureLayerRel)

                'Vérifier la présence d'éléments en relation
                If oFeatureColl.Count > 0 Then
                    'Interface pour extaire la distance
                    pProxOp = CType(pPoint, IProximityOperator)

                    'Traiter tous les éléments en relation
                    For Each pFeatureRel In oFeatureColl
                        'Vérifier si l'élément est différent
                        If Not (pFeature.OID = pFeatureRel.OID And gpFeatureLayerSelection.Name = pFeatureLayerRel.Name) Then
                            'Définir la géométrie à traiter
                            pGeometryRel = pFeatureRel.ShapeCopy
                            'Projeter la géométrie à traiter
                            pGeometryRel.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)
                            'Vérifier si la géométrie est une surface
                            If pGeometryRel.GeometryType = esriGeometryType.esriGeometryPolygon Then
                                'Interface pour extraire la limite
                                pTopoOp = CType(pGeometryRel, ITopologicalOperator)
                                'Extraire la limite
                                pGeometryRel = pTopoOp.Boundary
                            End If

                            'Extraire la distance
                            DistanceMinimumPointRelation = pProxOp.ReturnDistance(pGeometryRel)

                            'Vérifier si la distance est inférieure à la précision
                            If DistanceMinimumPointRelation < gdPrecision Then
                                'Sortir de la fonction
                                Exit Function

                                'Si la distance est inférieure à la distance minimum trouvée
                            ElseIf DistanceMinimumPointRelation < dDistanceMin Then
                                'Définir la distance minimum trouvée
                                dDistanceMin = DistanceMinimumPointRelation
                            End If
                        End If
                    Next
                End If
            Next

            'Retourner la distance minimum trouvé au besoin
            If dDistanceMin < gdTolerance Then DistanceMinimumPointRelation = dDistanceMin

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayerRel = Nothing
            oFeatureColl = Nothing
            pFeatureRel = Nothing
            pGeometryRel = Nothing
            pTopoOp = Nothing
            pProxOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la proximité de ligne interne respecte ou non 
    ''' la tolérance et la longueur spécifiée.
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
    '''<return>Les géométries de proximité entre les éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterProximiteLigneInterne(ByRef pTrackCancel As ITrackCancel,
                                                  Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométrie des éléments trouvés.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.

        Try
            'Définir la géométrie par défaut
            TraiterProximiteLigneInterne = New GeometryBag
            TraiterProximiteLigneInterne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterProximiteLigneInterne, IGeometryCollection)

            'Interface pour conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver la sélection de départ
            pSelectionSet = pFeatureSel.SelectionSet

            'Vider la sélection
            gpMap.ClearSelection()
            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Proximité de ligne interne (" & gpFeatureLayerSelection.Name & ") ..."

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Traiter la proximité de point de l'élément
                ElementProximiteLigneInterne(pFeature, pNewSelectionSet, pGeomResColl, bEnleverSelection)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeature = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter la proximité de ligne Interne d'un élément.
    '''</summary>
    '''
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pNewSelectionSet">'Interface pour sélectionner les éléments.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométrie des éléments trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    Private Sub ElementProximiteLigneInterne(ByRef pFeature As IFeature, ByRef pNewSelectionSet As ISelectionSet, _
                                             ByRef pGeomResColl As IGeometryCollection, ByVal bEnleverSelection As Boolean)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les points trouvés.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour créer un Buffer.
        Dim sMessage As String = ""                         'Contient le message.

        Try
            'Initialiser le message de retour
            sMessage = ""

            'Définir la polyline vide par défaut
            pGeometry = New Polyline
            pGeometry.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour ajouter les lignes
            pPolylineColl = CType(pGeometry, IGeometryCollection)

            'Définir la géométrie à traiter
            pGeometry = pFeature.ShapeCopy
            'Projeter la géométrie à traiter
            pGeometry.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)

            'Vérifier si la géométrie est un Polygon
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire la limite
                pTopoOp = CType(pGeometry, ITopologicalOperator)
                'Interface contenant la limite de la géométrie
                pGeometry = pTopoOp.Boundary
            End If

            'Traiter la proximité de ligne interne d'une composante
            pPolylineColl.AddGeometryCollection(CType(ProximiteLigneInterneGeometrie(pGeometry, gdTolerance, gdLongueur), IGeometryCollection))

            'Vérifier si on doit sélectionner l'élément
            If (pPolylineColl.GeometryCount = 0 And Not bEnleverSelection) Or (pPolylineColl.GeometryCount > 0 And bEnleverSelection) Then
                'Ajouter l'élément dans la sélection
                pNewSelectionSet.Add(pFeature.OID)

                'Vérifier si le multipoint n'est pas vide
                If pPolylineColl.GeometryCount > 0 Then
                    'Définir la géométrie
                    pGeometry = CType(pPolylineColl, IGeometry)
                    'Définir le message
                    sMessage = "Proximité de ligne interne"
                Else
                    'Définir le message
                    sMessage = "Aucune proximité de ligne interne"
                End If

                'Définir la géométrie de l'élément
                pGeomResColl.AddGeometry(pGeometry)

                'Écrire une erreur
                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage & " /" & gdTolerance.ToString("F1") & ", " & gdLongueur.ToString("F1"), pGeometry)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pPolylineColl = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de valider et retourner les incohérences de proximité intérieure 
    ''' contenues dans une géométrie en fonction des tolérances de largeur et de longueur.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie de l'élément à valider.</param>
    '''<param name="dLargeur">Contient la tolérance de proximité en largeur.</param>
    '''<param name="dLongueur">Contient la tolérance de proximité en longueur.</param>
    ''' 
    '''<returns>"IPolyline" contenant les incohérences de proximité intérieure, "Nothing" sinon.</returns>
    '''
    Private Function ProximiteLigneInterneGeometrie(ByRef pGeometry As IGeometry, ByVal dLargeur As Double, ByVal dLongueur As Double) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                    'Interface utiliser pour construire une ligne vide.
        Dim pIncoherenceColl As IGeometryCollection = Nothing   'Interface qui permet d'ajouter des composantes dans l'incohérence.
        Dim pGeometryColl As IGeometryCollection = Nothing      'Interface qui permet d'accéder aux composantes de la géométrie traitée.
        Dim pPath As IPath = Nothing                            'Interface contenant une composante à traiter.
        Dim pBaseSegColl As ISegmentCollection = Nothing        'Interface contenant une composante de base à traiter.
        Dim pCompareSegColl As ISegmentCollection = Nothing     'Interface contenant une composante à comparer à la composante de base.
        Dim pRelOp As IRelationalOperator = Nothing             'Interface utilisé pour créer une zone tampon.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface utilisé pour vérifier si les composantes se touchent entre elles.

        'Définir la valeur de retour par défaut
        ProximiteLigneInterneGeometrie = New Polyline
        ProximiteLigneInterneGeometrie.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les composantes l'incohérence
            pIncoherenceColl = CType(ProximiteLigneInterneGeometrie, IGeometryCollection)

            'Interface pour accéder aux composantes des géométries
            pGeometryColl = CType(pGeometry, IGeometryCollection)

            'Valider toutes les composantes de la géométrie
            For i = 0 To pGeometryColl.GeometryCount - 1
                'Définir la composante à traiter
                pPath = CType(pGeometryColl.Geometry(i), IPath)

                'Extraire les lignes en proximité à l'intérieur de la composante
                pPolyline = ProximiteLigneInterneComposante(pPath, dLargeur, dLongueur)

                'Ajouter les incohérences de la composante de la géométrie
                pIncoherenceColl.AddGeometryCollection(CType(pPolyline, IGeometryCollection))

                'Trouver les incohérences de proximité entre les composantes de la géométrie
                If i < pGeometryColl.GeometryCount - 1 Then
                    'Créer la ligne de base vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = pGeometry.SpatialReference
                    pBaseSegColl = CType(pPolyline, ISegmentCollection)

                    'Définir la composante de base
                    pBaseSegColl.AddSegmentCollection(CType(pGeometryColl.Geometry(i), ISegmentCollection))

                    'Comparer avec toutes les autres composantes de la géométrie
                    For j = i + 1 To pGeometryColl.GeometryCount - 1
                        'Créer la ligne à comparer vide
                        pPolyline = New Polyline
                        pPolyline.SpatialReference = pGeometry.SpatialReference
                        pCompareSegColl = CType(pPolyline, ISegmentCollection)

                        'Définir la composante à comparer
                        pCompareSegColl.AddSegmentCollection(CType(pGeometryColl.Geometry(j), ISegmentCollection))

                        'Interface pour créer une zone tampon
                        pTopoOp = CType(pCompareSegColl, ITopologicalOperator2)

                        'Créer une zone tampon pour la géométrie à comparer
                        pRelOp = CType(pTopoOp.Buffer(dLargeur), IRelationalOperator)

                        'Vérifier si les composantes se touchent entre elles
                        If Not pRelOp.Disjoint(CType(pBaseSegColl, IGeometry)) Then
                            'Interface pour trouver l'intersection
                            pTopoOp = CType(pRelOp, ITopologicalOperator2)

                            'Trouver les intersections entre les composantes
                            pPolyline = CType(pTopoOp.Intersect(CType(pBaseSegColl, IGeometry), esriGeometryDimension.esriGeometry1Dimension), IPolyline)

                            'Ajouter les composantes d'incohérences
                            pIncoherenceColl.AddGeometryCollection(CType(pPolyline, IGeometryCollection))
                        End If
                    Next j
                End If
            Next i

            'Détruire les lignes qui ne respectent pas la longueur requise
            DetruireSegmentLongueur(CType(pIncoherenceColl, IPolyline), dLongueur)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pIncoherenceColl = Nothing
            pGeometryColl = Nothing
            pBaseSegColl = Nothing
            pCompareSegColl = Nothing
            pRelOp = Nothing
            pTopoOp = Nothing
            pPath = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de valider et retourner les incohérences de proximité de ligne interne 
    ''' contenues dans une composante de géométrie en fonction des tolérances de largeur et de longueur.
    '''</summary>
    '''
    '''<param name="pPath">Interface ESRI contenant une composante de géométrie à valider.</param>
    '''<param name="dLargeur">Contient la tolérance de proximité en largeur.</param>
    '''<param name="dLongueur">Contient la tolérance de proximité en longueur.</param>
    ''' 
    '''<returns>"IPolyline" contenant les incohérences de proximité intérieure, "Nothing" sinon.</returns>
    '''
    Private Function ProximiteLigneInterneComposante(ByRef pPath As IPath, ByVal dLargeur As Double, ByVal dLongueur As Double) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                    'Interface utiliser pour construire une ligne vide.
        Dim pIncoherenceColl As IGeometryCollection = Nothing   'Interface qui permet d'ajouter des composantes dans l'incohérence.
        Dim pSegmentColl As ISegmentCollection = Nothing        'Interface qui permet d'accéder aux segments de la composante à traiter
        Dim pBaseSegColl As ISegmentCollection = Nothing        'Interface contenant une partie de composante de base à traiter.
        Dim pCompareSegColl As ISegmentCollection = Nothing     'Interface contenant une partie de composante à comparer à celle de base.
        Dim pRelOp As IRelationalOperator = Nothing             'Interface utilisé pour créer une zone tampon.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface utilisé pour vérifier si les composantes se touchent entre elles.

        'Définir la valeur de retour par défaut
        ProximiteLigneInterneComposante = New Polyline
        ProximiteLigneInterneComposante.SpatialReference = pPath.SpatialReference

        Try
            'Interface pour ajouter les composantes l'incohérence
            pIncoherenceColl = CType(ProximiteLigneInterneComposante, IGeometryCollection)

            'Interface pour traiter les segments de la composante
            pSegmentColl = CType(pPath, ISegmentCollection)

            'Traiter tous les segments de la composante ligne
            For i = 0 To pSegmentColl.SegmentCount - 2
                'Créer la ligne de base vide
                pPolyline = New Polyline
                pPolyline.SpatialReference = pPath.SpatialReference
                pBaseSegColl = CType(pPolyline, ISegmentCollection)

                'Initialiser la polyline
                pBaseSegColl.AddSegment(pSegmentColl.Segment(i))

                'Créer la ligne à comparer vide
                pPolyline = New Polyline
                pPolyline.SpatialReference = pPath.SpatialReference
                pCompareSegColl = CType(pPolyline, ISegmentCollection)

                'Vérifier si les segments se croises
                For j = i + 1 To pSegmentColl.SegmentCount - 1
                    'Initialiser la polyline à comparer
                    pCompareSegColl.AddSegment(pSegmentColl.Segment(j))
                Next j

                'Interface pour créer un buffer
                pTopoOp = CType(pCompareSegColl, ITopologicalOperator2)

                'Créer le buffer de la partie à comparer
                pRelOp = CType(pTopoOp.Buffer(dLargeur), IRelationalOperator)

                'Vérifier si deux segments se touchent
                If Not pRelOp.Disjoint(CType(pBaseSegColl, IGeometry)) Then
                    'Interface pour extraire l'intersection
                    pTopoOp = CType(pRelOp, ITopologicalOperator2)

                    'Extraire la partie commune entre la ligne de base et le buffer de celui à comparer
                    pPolyline = CType(pTopoOp.Intersect(CType(pBaseSegColl, IGeometry), esriGeometryDimension.esriGeometry1Dimension), IPolyline)

                    'Ajouter les nouvelles erreurs
                    pIncoherenceColl.AddGeometryCollection(CType(pPolyline, IGeometryCollection))
                End If
            Next i

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pIncoherenceColl = Nothing
            pSegmentColl = Nothing
            pBaseSegColl = Nothing
            pCompareSegColl = Nothing
            pRelOp = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la proximité de ligne respecte ou non la tolérance et la longueur spécifiée.
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
    '''<return>Les géométries de proximité entre les éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterProximiteLigne(ByRef pTrackCancel As ITrackCancel,
                                           Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométrie des éléments trouvés.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant la relation spatiale de base.

        Try
            'Définir la géométrie par défaut
            TraiterProximiteLigne = New GeometryBag
            TraiterProximiteLigne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterProximiteLigne, IGeometryCollection)

            'Interface pour conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver la sélection de départ
            pSelectionSet = pFeatureSel.SelectionSet

            'Vider la sélection
            gpMap.ClearSelection()
            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Proximité de ligne des géométries (" & gpFeatureLayerSelection.Name & ") ..."

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Créer la requête spatiale
            pSpatialFilter = New SpatialFilterClass
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Traiter la proximité de point de l'élément
                ElementProximiteLigne(pFeature, pSpatialFilter, pNewSelectionSet, pGeomResColl, bEnleverSelection)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeature = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter la proximité de ligne d'un élément.
    '''</summary>
    '''
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pSpatialFilter">Interface contenant la requête spatial pour trouver les éléments en relation.</param>
    '''<param name="pNewSelectionSet">'Interface pour sélectionner les éléments.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométrie des éléments trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    Private Sub ElementProximiteLigne(ByRef pFeature As IFeature, ByRef pSpatialFilter As ISpatialFilter, _
                                      ByRef pNewSelectionSet As ISelectionSet, ByRef pGeomResColl As IGeometryCollection, ByVal bEnleverSelection As Boolean)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvées.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour créer un Buffer.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire le nombre de sommets.
        Dim sMessage As String = ""                         'Contient le message.

        Try
            'Initialiser le message de retour
            sMessage = ""

            'Définir la géométrie à traiter
            pGeometry = pFeature.ShapeCopy
            'Projeter la géométrie à traiter
            pGeometry.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)

            'Vérifier si la géométrie est un Polygon
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire la limite
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                'Interface contenant la limite de la géométrie
                pGeometry = pTopoOp.Boundary
            End If

            'Interface pour extraire le nombre de composantes
            pPolylineColl = CType(ProximiteLigneGeometrie(pGeometry, pFeature, pSpatialFilter), IGeometryCollection)

            'Vérifier si on doit sélectionner l'élément
            If (pPolylineColl.GeometryCount = 0 And Not bEnleverSelection) Or (pPolylineColl.GeometryCount > 0 And bEnleverSelection) Then
                'Ajouter l'élément dans la sélection
                pNewSelectionSet.Add(pFeature.OID)

                'Vérifier si la ligne résultante contient seulement 1 composante
                If pPolylineColl.GeometryCount = 1 Then
                    'Définir la géométrie
                    pGeometry = CType(pPolylineColl, IGeometry)
                    'Définir le message
                    sMessage = "Proximité de ligne"

                    'Interface pour extraire le nombre de sommets
                    pPointColl = CType(pGeometry, IPointCollection)

                    'Définir la géométrie de l'élément
                    pGeomResColl.AddGeometry(pGeometry)

                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage _
                                        & " /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", Longueur=" & gdLongueur.ToString("F1") _
                                        & ", NbSommets=" & pPointColl.PointCount.ToString, pGeometry, pPointColl.PointCount)

                    'Si la ligne résultante contient plusieurs composantes
                ElseIf pPolylineColl.GeometryCount > 1 Then
                    'Définir le message
                    sMessage = "Proximité de ligne"

                    'Traiter toutes les composantes
                    For i = 0 To pPolylineColl.GeometryCount - 1
                        'Définir la géométrie
                        pGeometry = PathToPolyline(CType(pPolylineColl.Geometry(i), IPath))

                        'Interface pour extraire le nombre de sommets
                        pPointColl = CType(pGeometry, IPointCollection)

                        'Définir la géométrie de l'élément
                        pGeomResColl.AddGeometry(pGeometry)

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage _
                                            & " /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", Longueur=" & gdLongueur.ToString("F1") _
                                            & ", NbSommets=" & pPointColl.PointCount.ToString, pGeometry, pPointColl.PointCount)
                    Next

                    'Si la ligne résultante est vide
                Else
                    'Définir le message
                    sMessage = "Aucune proximité de ligne"

                    'Interface pour extraire le nombre de sommets
                    pPointColl = CType(pGeometry, IPointCollection)

                    'Définir la géométrie de l'élément
                    pGeomResColl.AddGeometry(pGeometry)

                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage _
                                        & " /Précision=" & gdPrecision.ToString & ", Tolérance=" & gdTolerance.ToString & ", Longueur=" & gdLongueur.ToString("F1") _
                                        & ", NbSommets=" & pPointColl.PointCount.ToString, pGeometry, pPointColl.PointCount)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pPolylineColl = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire les proximités de ligne avec les géométries en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la ligne à traitée.</param>
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pSpatialFilter">Interface contenant la requête spatial pour trouver les éléments en relation.</param>
    ''' 
    '''<returns>Distance minimum entre un sommet d'un élément et ses éléments en relation.</returns>
    '''
    Private Function ProximiteLigneGeometrie(ByRef pGeometry As IGeometry, ByRef pFeature As IFeature, ByRef pSpatialFilter As ISpatialFilter) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvées.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim oFeatureColl As Collection = Nothing            'Objet contenant la collection des éléments en relation.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant un élément en relation.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie en relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour créer un Buffer.
        Dim pSpatialIndex As ISpatialIndex2 = Nothing        'Interface utilisée pour définir un index spatial dans un GeometryBag.

        'Définir la valeur de retour par défaut
        ProximiteLigneGeometrie = New Polyline
        ProximiteLigneGeometrie.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les lignes trouvées
            pPolylineColl = CType(ProximiteLigneGeometrie, IGeometryCollection)

            'Interface pour créer un buffer
            pTopoOp = CType(pGeometry, ITopologicalOperator2)
            'Définir la géométrie utilisée pour la relation spatiale
            pSpatialFilter.Geometry = pTopoOp.Buffer(gdTolerance)

            'Indexer la géométrie
            pSpatialIndex = CType(pSpatialFilter.Geometry, ISpatialIndex2)
            pSpatialIndex.AllowIndexing = True
            pSpatialIndex.Invalidate()

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Définir la référence spatiale de sortie dans la requête spatiale
                pSpatialFilter.OutputSpatialReference(pFeatureLayerRel.FeatureClass.ShapeFieldName) = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

                'Extraire les éléments en relation
                oFeatureColl = ExtraireElementsRelation(pSpatialFilter, pFeatureLayerRel)

                'Vérifier la présence d'éléments en relation
                If oFeatureColl.Count > 0 Then
                    'Traiter tous les éléments en relation
                    For Each pFeatureRel In oFeatureColl
                        'Vérifier si l'élément est différent
                        If Not (pFeature.OID = pFeatureRel.OID And gpFeatureLayerSelection.Name = pFeatureLayerRel.Name) Then
                            'Définir la géométrie à traiter
                            pGeometryRel = pFeatureRel.ShapeCopy
                            'Projeter la géométrie à traiter
                            pGeometryRel.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)
                            'Vérifier si la géométrie est une surface
                            If pGeometryRel.GeometryType = esriGeometryType.esriGeometryPolygon Then
                                'Interface pour extraire la limite du polygone
                                pTopoOp = CType(pGeometryRel, ITopologicalOperator2)
                                'Extraire la limite
                                pGeometryRel = pTopoOp.Boundary
                            End If

                            'Ajouter les proximités de ligne trouvées
                            pPolylineColl.AddGeometryCollection(CType(pGeometryRel, IGeometryCollection))
                        End If
                    Next
                End If
            Next

            'Définir la Géométrie des éléments en relation
            pGeometryRel = CType(pPolylineColl, IGeometry)
            'Interface pour simplifier la géométrie des éléments en relation
            pTopoOp = CType(pGeometryRel, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            ''Indexer la géométrie
            'pSpatialIndex = CType(pGeometryRel, ISpatialIndex2)
            'pSpatialIndex.AllowIndexing = True
            'pSpatialIndex.Invalidate()

            'Interface pour extraire l'intersection avec les géométries en relation
            pTopoOp = CType(pSpatialFilter.Geometry, ITopologicalOperator2)
            'Extraire l'intersection entre la buffer de la géométrie et la géométrie en relation
            pTopoOp = CType(pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry1Dimension), ITopologicalOperator2)
            'Extraire la partie commune entre la géométrie traitée et celle en relation
            pPolyline = CType(pTopoOp.Difference(pGeometry), IPolyline)
            'Définir le résultat obtenu
            ProximiteLigneGeometrie = pPolyline

            'Détruire les lignes qui ne respectent pas la longueur requise
            DetruireSegmentLongueur(ProximiteLigneGeometrie, gdLongueur)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pPolylineColl = Nothing
            pFeatureLayerRel = Nothing
            oFeatureColl = Nothing
            pFeatureRel = Nothing
            pGeometryRel = Nothing
            pTopoOp = Nothing
            pSpatialIndex = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'extraire les proximités de ligne avec les géométries en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la ligne à traitée.</param>
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pSpatialFilter">Interface contenant la requête spatial pour trouver les éléments en relation.</param>
    ''' 
    '''<returns>Distance minimum entre un sommet d'un élément et ses éléments en relation.</returns>
    '''
    Private Function ProximiteLigneGeometrieOld(ByRef pGeometry As IGeometry, ByRef pFeature As IFeature, ByRef pSpatialFilter As ISpatialFilter) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvées.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim oFeatureColl As Collection = Nothing            'Objet contenant la collection des éléments en relation.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant un élément en relation.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie en relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour créer un Buffer.
        Dim pSpatialIndex As ISpatialIndex2 = Nothing        'Interface utilisée pour définir un index spatial dans un GeometryBag.

        'Définir la valeur de retour par défaut
        ProximiteLigneGeometrieOld = New Polyline
        ProximiteLigneGeometrieOld.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les lignes trouvées
            pPolylineColl = CType(ProximiteLigneGeometrieOld, IGeometryCollection)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Définir la référence spatiale de sortie dans la requête spatiale
                pSpatialFilter.OutputSpatialReference(pFeatureLayerRel.FeatureClass.ShapeFieldName) = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

                'Interface pour créer un buffer
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                'Définir la géométrie utilisée pour la relation spatiale
                pSpatialFilter.Geometry = pTopoOp.Buffer(gdTolerance)

                'Indexer le Bag des éléments en relation
                pSpatialIndex = CType(pSpatialFilter.Geometry, ISpatialIndex2)
                pSpatialIndex.AllowIndexing = True
                pSpatialIndex.Invalidate()

                'Extraire les éléments en relation
                oFeatureColl = ExtraireElementsRelation(pSpatialFilter, pFeatureLayerRel)

                'Vérifier la présence d'éléments en relation
                If oFeatureColl.Count > 0 Then
                    'Traiter tous les éléments en relation
                    For Each pFeatureRel In oFeatureColl
                        'Vérifier si l'élément est différent
                        If Not (pFeature.OID = pFeatureRel.OID And gpFeatureLayerSelection.Name = pFeatureLayerRel.Name) Then
                            'Définir la géométrie à traiter
                            pGeometryRel = pFeatureRel.ShapeCopy
                            'Projeter la géométrie à traiter
                            pGeometryRel.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)
                            'Vérifier si la géométrie est une surface
                            If pGeometryRel.GeometryType = esriGeometryType.esriGeometryPolygon Then
                                'Interface pour extraire la limite du polygone
                                pTopoOp = CType(pGeometryRel, ITopologicalOperator2)
                                'Extraire la limite
                                pGeometryRel = pTopoOp.Boundary
                            End If
                            'Interface pour extraire l'intersection avec les géométries en relation
                            pTopoOp = CType(pSpatialFilter.Geometry, ITopologicalOperator2)
                            'Extraire l'intersection entre la buffer de la géométrie et la géométrie en relation
                            pTopoOp = CType(pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry1Dimension), ITopologicalOperator2)
                            'Extraire la partie commune entre la géométrie traitée et celle en relation
                            pPolyline = CType(pTopoOp.Difference(pGeometry), IPolyline)
                            'Ajouter les proximités de ligne trouvées
                            pPolylineColl.AddGeometryCollection(CType(pPolyline, IGeometryCollection))
                        End If
                    Next
                End If
            Next

            'Détruire les lignes qui ne respectent pas la longueur requise
            DetruireSegmentLongueur(ProximiteLigneGeometrieOld, gdLongueur)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pPolylineColl = Nothing
            pFeatureLayerRel = Nothing
            oFeatureColl = Nothing
            pFeatureRel = Nothing
            pGeometryRel = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de de détruire les segments dont la longueur est inférieure à une longueur.
    '''</summary>
    '''
    '''<param name="pPolyline">Interface ESRI contenant une ligne de proximité.</param>
    '''<param name="dLongueur">Contient la tolérance de proximité en longueur.</param>
    '''
    Private Sub DetruireSegmentLongueur(ByRef pPolyline As IPolyline, ByVal dLongueur As Double)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface qui permet d'ajouter des composantes dans l'incohérence.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante.
        Dim pSegmentColl As ISegmentCollection = Nothing    'Interface qui permet d'accéder aux segments de la composante à traiter.
        Dim pSegment As ISegment = Nothing                  'Interface contenant un segment de la composante à traiter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour vérifier si les composantes se touchent entre elles.

        Try
            'Sortir si vide
            If pPolyline.IsEmpty Then Return

            'Simplifier la géométrie en erreur
            pTopoOp = CType(pPolyline, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Interface pour extraire les composantes de la Polyline
            pGeomColl = CType(pPolyline, IGeometryCollection)

            'Traiter tous les composantes
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Interface pour extraire la longueur de la composante
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Interface pour extraire les segments
                pSegmentColl = CType(pPath, ISegmentCollection)

                'Si deux segments et plus
                If pSegmentColl.SegmentCount > 1 Then
                    'Définir le segment
                    pSegment = pSegmentColl.Segment(pSegmentColl.SegmentCount - 1)
                    'Vérifier si le segment est inférieure à une longueur
                    If pSegment.Length < dLongueur Then
                        'Détruire le segment
                        pSegmentColl.RemoveSegments(pSegmentColl.SegmentCount - 1, 1, False)
                    End If
                    'Définir le segment
                    pSegment = pSegmentColl.Segment(0)
                    'Vérifier si le segment est inférieure à une longueur
                    If pSegment.Length < dLongueur Then
                        'Détruire le segment
                        pSegmentColl.RemoveSegments(0, 1, False)
                    End If
                End If

                'Vérifier si la composante est inférieure à une longueur
                If pPath.Length < dLongueur Then
                    'Détruire la composante
                    pGeomColl.RemoveGeometries(i, 1)
                End If
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPath = Nothing
            pSegmentColl = Nothing
            pSegment = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region
End Class
