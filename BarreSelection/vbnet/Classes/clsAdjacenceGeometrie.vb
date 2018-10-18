Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.ArcMapUI

'**
'Nom de la composante : clsAdjacenceGeometrie.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont l'adjacence entre les géométries respecte ou non les paramètres spécifiés.
''' 
''' Ce traitement est utilisé seulement sur des FeatureClass de type Polyline ou Polygone.
''' 
''' La classe permet de traiter les deux attributs d'adjacence SEGMENTATION et FUSION.
''' 
''' SEGMENTATION : 
'''    La segmentation est valide lorsqu'il y a les conditions suivantes :
'''    -Lorsque deux éléments d'une même classe de type ligne sont adjacents en un point
'''     ou lorsque deux éléments d'une même classe de type surface sont adjacents en une ligne.
'''    -lorsqu'au moins une valeur d'attribut est différente entre les deux éléments
'''     ou lorsqu'il y a plus de deux éléments adjacents.
''' 
''' FUSION : 
'''    La fusion est valide lorsqu'il y a les conditions suivantes :
'''      -Lorsqu'il n'y a aucun élément en relation.
'''      -Pour les lignes, si le nombre de Edges est inférieure ou égale à deux.
'''      -Pour les surfaces, il n'y a pas d'intersection entre l'intérieure de la surface et l'intérieur de la limite des surfaces en relation.
'''    Si ces conditions ne sont pas respectées, la segmentation est nécessaire afin de les rendre adjacents.
''' 
''' Note : La topologie est utilisée pour traiter l'adjacence entre les géométries sauf pour valider la fusion des surfaces.
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 11 juin 2015
'''</remarks>
''' 
Public Class clsAdjacenceGeometrie
    Inherits clsValeurAttribut

    'Déclarer les variables globales
    '''<summary>Liste des attributs à exclure du traitement d'adjacence.</summary>
    Protected gsExclureAttribut As String = ""

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "SEGMENTATION"
            Expression = gdPrecision.ToString("F3")
            ExclureAttribut = "OBJECTID"
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
    '''<param name="sExclureAttribut"> Attributs à exclure de la comparaison.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String, Optional ByVal sExclureAttribut As String = "")
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            ExclureAttribut = sExclureAttribut
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
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
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
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gsExclureAttribut = Nothing
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
            Nom = "AdjacenceGeometrie"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la liste des attributs à exclure du traitement de comparaison.
    '''</summary>
    ''' 
    Public Property ExclureAttribut() As String
        Get
            ExclureAttribut = gsExclureAttribut
        End Get
        Set(ByVal value As String)
            gsExclureAttribut = value
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
            'Vérifier la présence du masque
            If gsExclureAttribut <> "" Then
                'Retourner aussi le masque spatial
                Parametres = Parametres & " " & gsExclureAttribut
            End If
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

            'Mettre en majuscule les paramètres
            value = value
            'Extraire les paramètres
            params = value.Split(CChar(" "))
            'Vérifier si les deux paramètres sont présents
            If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION [EXCLURE_ATTRIBUT]")

            'Définir les valeurs par défaut
            gsNomAttribut = params(0).ToUpper
            gsExpression = params(1)

            'Vérifier si le troisième paramètre est présent
            If params.Length > 2 Then
                'Définir les attributs à exclure de la comparaison
                gsExclureAttribut = params(2)
            Else
                gsExclureAttribut = ""
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
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing            'Interface contenant les attributs de la FeatureClass.
        Dim sAttributs As String = ""               'Liste des attributs de la Featureclass de sélection.
        Dim sAttributsNonEditable As String = ""    'Liste des attributs de la Featureclass de sélection non éditables.

        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir l'interface contenant les attributs de la Featureclass
                pFields = gpFeatureLayerSelection.FeatureClass.Fields

                'Traiter tous les attributs
                For i = 0 To pFields.FieldCount - 1
                    'Vérifier si l'attribut est un OID ou un shape
                    If Not (pFields.Field(i).Name = gpFeatureLayerSelection.FeatureClass.OIDFieldName _
                    Or pFields.Field(i).Name = gpFeatureLayerSelection.FeatureClass.ShapeFieldName) Then
                        'Vérifier si c'est le premier attribut
                        If sAttributs = "" Then
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = pFields.Field(i).Name
                        Else
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = sAttributs & "," & pFields.Field(i).Name
                        End If
                    End If
                    'vérifier si l'attribut est non éditable
                    If pFields.Field(i).Editable = False Then
                        'Vérifier si c'est le premier attribut
                        If sAttributsNonEditable = "" Then
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributsNonEditable = pFields.Field(i).Name
                        Else
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributsNonEditable = sAttributsNonEditable & "," & pFields.Field(i).Name
                        End If
                    End If
                Next

                'Définir le paramètre pour segmenter selon la géométrie
                ListeParametres.Add("FUSION " & gdPrecision.ToString("F3"))

                'Définir le paramètre pour fusionner selon tous les attributs sauf les non-éditables
                ListeParametres.Add("SEGMENTATION " & gdPrecision.ToString("F3") & " BDG_ID,MEP_ID,ZT_ID")

                'Définir le paramètre pour fusionner selon tous les attributs sauf les non-éditables
                ListeParametres.Add("SEGMENTATION " & gdPrecision.ToString("F3") & " " & sAttributsNonEditable)

                'Définir le paramètre pour fusionner selon seulement la géométrie
                ListeParametres.Add("SEGMENTATION " & gdPrecision.ToString("F3") & " " & gpFeatureLayerSelection.FeatureClass.OIDFieldName & "," & sAttributs)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFields = Nothing
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
            'Définir la valeur par défaut, la contrainte est invalide
            FeatureClassValide = False
            gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polyline ou Polygon."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'Vérifier si la FeatureClass est de type Polyline ou Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'La contrainte est valide.
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide"
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
        'La contrainte est invalide par défaut.
        AttributValide = False
        gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut

        Try
            'Vérifier si l'attribut est valide
            If gsNomAttribut = "SEGMENTATION" Or gsNomAttribut = "FUSION" Then
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
    ''' Routine qui permet d'indiquer si l'expression régulière est valide.
    ''' 
    '''<return>Boolean qui indique si l'expression régulière est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function ExpressionValide() As Boolean
        Try
            'Retourner l'expression invalide par défaut
            ExpressionValide = False
            gsMessage = "ERREUR : L'expression est invalide."

            'Vérifier si l'expression est numérique
            If TestDBL(gsExpression) Then
                'Retourner l'expression valide
                ExpressionValide = True
                gsMessage = "La contrainte est valide"
                'Définir la précision
                gdPrecision = ConvertDBL(gsExpression)
                'Si l'expression n'est pas numérique
            Else
                gsMessage = "ERREUR : L'expression n'est pas numérique."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les paramètres d'adjacence avec ses éléments en relation.
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

            'Vérifier si la classe à traiter est de type ligne
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                'Afficher le message de création de la topologie
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If

                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Vérifier si la précision spécifié est supérieure à celle par défaut
                If gdPrecision < pSRTolerance.XYTolerance Then gdPrecision = pSRTolerance.XYTolerance
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, gdPrecision)

                'Vérifier la topologie est valide
                If pTopologyGraph IsNot Nothing Then
                    'Si le nom de l'attribut est SEGMENTATION
                    If gsNomAttribut = "SEGMENTATION" Then
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon la segmentation des lignes (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer
                        Selectionner = TraiterAdjacenceSegmentationLigne(pTopologyGraph, pTrackCancel, bEnleverSelection)

                        'Si le nom de l'attribut est FUSION
                    ElseIf gsNomAttribut = "FUSION" Then
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon la fusion des lignes (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer
                        Selectionner = TraiterAdjacenceFusionLigne(pTopologyGraph, pTrackCancel, bEnleverSelection)
                    End If

                    'Si la topologie est invalide
                Else
                    'Retourner une erruer de création de la topologie
                    Err.Raise(-1, , "Incapable de créer la topologie!")
                End If

                'Vérifier si la classe à traiter est de type surface
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Si le nom de l'attribut est SEGMENTATION
                If gsNomAttribut = "SEGMENTATION" Then
                    'Afficher le message de création de la topologie
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                    'Vérifier si la Layer de sélection est absent dans les Layers relations
                    If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                        'Ajouter le layer de sélection dans les layers en relation
                        gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                    End If

                    'Interface pour extraire la tolérance de précision de la référence spatiale
                    pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                    'Vérifier si la précision spécifié est supérieure à celle par défaut
                    If gdPrecision < pSRTolerance.XYTolerance Then gdPrecision = pSRTolerance.XYTolerance
                    'Création de la topologie
                    pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, gdPrecision)

                    'Vérifier la topologie est valide
                    If pTopologyGraph IsNot Nothing Then
                        'Afficher le message de sélection en cours
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon la segmentation des surfaces (" & gpFeatureLayerSelection.Name & ") ..."
                        'Traiter le FeatureLayer
                        Selectionner = TraiterAdjacenceSegmentationSurface(pTopologyGraph, pTrackCancel, bEnleverSelection)

                        'Si la topologie est invalide
                    Else
                        'Retourner une erruer de création de la topologie
                        Err.Raise(-1, , "Incapable de créer la topologie!")
                    End If

                    'Si le nom de l'attribut est FUSION
                ElseIf gsNomAttribut = "FUSION" Then
                    'Afficher le message de sélection en cours
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon la fusion des surfaces (" & gpFeatureLayerSelection.Name & ") ..."
                    'Traiter le FeatureLayer
                    Selectionner = TraiterAdjacenceFusionSurface(pTrackCancel, bEnleverSelection)
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
            pTopologyGraph = Nothing
            pSRTolerance = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la fusion entre les surfaces d'une même classe respecte les paramètres spécifiés. 
    ''' La fusion est valide lorsqu'il y a les conditions suivantes :
    '''    -Lorsqu'il n'y a aucun élément en relation.
    '''    -Pour les surfaces, il n'y a pas d'intersection entre l'intérieure de la surface et l'intérieur de la limite des surfaces en relation. 
    ''' Si ces conditions ne sont pas respectées, la segmentation est nécessaire afin de les rendre adjacents.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    '''
    Private Function TraiterAdjacenceFusionSurface(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
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

        Try
            'Définir la géométrie par défaut
            TraiterAdjacenceFusionSurface = New GeometryBag
            TraiterAdjacenceFusionSurface.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterAdjacenceFusionSurface, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Lire les éléments en relation
                LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, True)

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale " + gsNomAttribut + " (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.RelationEx(CType(pGeomRelColl, IGeometryBag), esriSpatialRelationEnum.esriSpatialRelationInteriorIntersection)

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Traiter les Oids d'éléments trouvés
                TraiterListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, iOidAdd, bEnleverSelection, pTrackCancel, pGeomResColl)
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

            'Conserver la sélection trouvé
            pFeatureSel.SelectionSet = pSelectionSet

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
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter la liste des Oids d'éléments du FeatureLayer trouvés qui doivent être segmentés. 
    ''' L'écriture des erreurs de segmentation peut être effectué au besoin.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés par élément.</param>
    '''<param name="bEnleverSelection">Indique si on veut enlever les éléments trouvés de la sélection.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl">GeometryBag contenant les géométries en erreur.</param>
    ''' 
    Private Sub TraiterListeOid(ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, ByVal pGeomRelColl As IGeometryCollection, _
                                 ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByRef iOidAdd() As Integer, ByVal bEnleverSelection As Boolean, _
                                 ByRef pTrackCancel As ITrackCancel, ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la ligne en erreur utilisée pour la segmentation.
        Dim pPolylineErr As IPolyline = Nothing             'Contient la ligne en erreur utilisée pour la segmentation.
        Dim pLimiteSurface As IPolyline = Nothing           'Contient la limite de la surface traitée.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface pour extraire le nombre de ligne.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Indiquer que le OID respecte la relation
                iOidAdd(iSel) = iOidAdd(iSel) + 1

                'Vérifier s'il faut écrire les erreurs de segementation
                If bEnleverSelection Then
                    'Interface pour extraire la ligne utilisée pour la segmentation.
                    pTopoOp = CType(pGeomSelColl.Geometry(iSel), ITopologicalOperator2)

                    'Définir la limite de la surface traitée
                    pLimiteSurface = CType(pTopoOp.Boundary, IPolyline)

                    'Extraire la ligne en erreur utilisée pour la segmentation.
                    pTopoOp = CType(pTopoOp.Intersect(pGeomRelColl.Geometry(iRel), esriGeometryDimension.esriGeometry1Dimension), ITopologicalOperator2)

                    'Extraire la ligne en erreur utilisée pour la segmentation.
                    pPolylineErr = CType(pTopoOp.Difference(pLimiteSurface), IPolyline)

                    'Interface pour extraire le nombre de lignes
                    pGeometryColl = CType(pPolylineErr, IGeometryCollection)

                    'Ajouter la géométrie en erreur dans le Bag
                    pGeomResColl.AddGeometry(pPolylineErr)

                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & iOidSel(iSel).ToString & "," & iOidRel(iRel).ToString & " #Fusion invalide /NbLignes=" & pGeometryColl.GeometryCount.ToString & " /" & Parametres, _
                                        pPolylineErr, pGeometryColl.GeometryCount)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPolylineErr = Nothing
            pLimiteSurface = Nothing
            pGeometryColl = Nothing
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non la segementation 
    ''' et écriture des erreurs qui respectent la segmentation au besoin.
    '''</summary>
    '''
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on veut enlever les éléments trouvés de la sélection.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pSelectionSet">Interface contenant les éléments sélectionnés.</param>
    '''<param name="pGeomResColl">GéométryBag contenant les géométries en erreur.</param>
    ''' 
    Private Sub SelectionnerElementErreur(ByVal pGeomSelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidAdd() As Integer, _
                                          ByVal bEnleverSelection As Boolean, ByRef pTrackCancel As ITrackCancel, ByRef pSelectionSet As ISelectionSet, _
                                          ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la ligne en erreur utilisée pour la segmentation.
        Dim pPolylineErr As IPolyline = Nothing             'Contient la ligne en erreur.
        Dim pGDBridge As IGeoDatabaseBridge2 = Nothing      'Interface pour ajouter une liste de OID dans le SelectionSet.
        Dim iOid(0) As Integer      'Liste des OIDs trouvés.
        Dim iNbOid As Integer = 0   'Compteur de OIDs trouvés.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si on veut conserver ceux qui ne respectent pas la relation
            If bEnleverSelection Then
                'Traiter tous les éléments
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Vérifier si on doit ajouter le OID dans la sélection
                    If iOidAdd(i) > 0 Then
                        'Définir la liste des OIDs à sélectionner
                        iNbOid = iNbOid + 1

                        'Redimensionner la liste des OIDs à sélectionner
                        ReDim Preserve iOid(iNbOid - 1)

                        'Définir le OIDs dans la liste
                        iOid(iNbOid - 1) = iOidSel(i)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Si on veut conserver ceux qui respectent la relation
            Else
                'Traiter tous les éléments
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Vérifier si on doit ajouter le OID dans la sélection
                    If iOidAdd(i) = 0 Then
                        'Définir la liste des OIDs à sélectionner
                        iNbOid = iNbOid + 1

                        'Redimensionner la liste des OIDs à sélectionner
                        ReDim Preserve iOid(iNbOid - 1)

                        'Définir le OIDs dans la liste
                        iOid(iNbOid - 1) = iOidSel(i)

                        'Interface pour extraire la limite de la surface traitée.
                        pTopoOp = CType(pGeomSelColl.Geometry(i), ITopologicalOperator2)

                        'Extraire la limite de la surface traitée.
                        pPolylineErr = CType(pTopoOp.Boundary, IPolyline)

                        'Ajouter la géométrie en erreur dans le Bag
                        pGeomResColl.AddGeometry(pPolylineErr)

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #NbOidRel=" & iOidAdd(i).ToString & " /Fusion valide /" & Parametres, pPolylineErr, iOidAdd(i))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

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
            pTopoOp = Nothing
            pPolylineErr = Nothing
            pGDBridge = Nothing
            iOid = Nothing
            iNbOid = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la segmentation présente entre deux surfaces adjacentes en une ligne respecte les paramètres spécifiés. 
    ''' La segmentation est valide lorsqu'il y a les conditions suivantes :
    '''    -Lorsque deux éléments d'une même classe de type surface sont adjacents en une ligne.
    '''    -lorsqu'au moins une valeur d'attribut est différente entre les deux éléments
    '''     ou lorsqu'il y a plus de deux éléments adjacents.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterAdjacenceSegmentationSurface(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                         Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge. 
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pZAware As IZAware = Nothing                    'Interface utilisé pour enlever le Z. 
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant les EDGES en erreur. 
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe d'un élément
        Dim pFeatureAdj As IFeature = Nothing               'Interface contenant l'élément adjacent.
        Dim sMessage As String = ""                         'Contient le message d'erreur.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bFusion As Boolean = False                      'Indique que l'élément est à fusionner avec un autre.

        Try
            'Définir la géométrie par défaut
            TraiterAdjacenceSegmentationSurface = New GeometryBag
            TraiterAdjacenceSegmentationSurface.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet
            pFeatureSel.Clear()

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterAdjacenceSegmentationSurface, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
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
                        'Interface pour extraire le nombre d'intersections
                        pEnumTopoParent = pTopoEdge.Parents()

                        'Si le nombre d'éléments dans le Edge est 2, on peut peut-être fusionner
                        If pEnumTopoParent.Count = 2 Then
                            'Initialiser l'extraction
                            pEnumTopoParent.Reset()
                            'Extraire le premier élément
                            pEsriTopoParent = pEnumTopoParent.Next()
                            'Extraire la FeatureClass
                            pFeatureClass = pEsriTopoParent.m_pFC
                            'Extraire l'élément
                            pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                            'Vérifier si l'élément adjacent est différent l'élément traité
                            If pFeature.OID = pFeatureAdj.OID And pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                'Extraire le premier élément
                                pEsriTopoParent = pEnumTopoParent.Next()
                                'Extraire la FeatureClass
                                pFeatureClass = pEsriTopoParent.m_pFC
                                'Extraire l'élément
                                pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                            End If

                            'Vérifier si la classe est la même entre les deux éléments
                            If pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                'Comparer les valeurs d'attribut entre les 2 éléments
                                sMessage = ComparerAttributElement(pFeature, pFeatureAdj, gsExclureAttribut)
                                'Vérifier si les valeurs d'attributs sont les mêmes
                                If sMessage = "" Then
                                    'La fusion doit se faire avec l'élément adjacent
                                    bFusion = True
                                    sMessage = "," & pFeatureAdj.OID.ToString & " #Segmentation invalide, une fusion est nécessaire."

                                    'S'il y a une différence entre les valeurs d'attributs
                                Else
                                    'Aucune fusion n'est possible
                                    bFusion = False
                                    sMessage = "," & pFeatureAdj.OID.ToString & " #Segmentation valide, aucune fusion à effectuer /Différence=" & sMessage
                                End If

                                'Si la classe entre les 2 éléments est différent
                            Else
                                'Aucune fusion n'est possible
                                bFusion = False
                                sMessage = "," & pFeatureAdj.OID.ToString & " #Segmentation valide, aucune fusion à effectuer /Différence=" & pFeature.Class.AliasName & "<>" & pFeatureAdj.Class.AliasName
                            End If

                            'Si le nombre d'éléments dans le Edge est <> 2
                        Else
                            'Aucune fusion n'est possible
                            bFusion = False
                            sMessage = " #Segmentation valide, aucune fusion à effectuer /Différence=NbElement<>2"
                        End If

                        'Vérifier si on doit sélectionner l'élément
                        If (bFusion And bEnleverSelection) Or (Not bFusion And Not bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Conserver le EDGE en erreur
                            pPolylineErr = CType(pTopoEdge.Geometry, IPolyline)
                            'Enlever le Z du EDGE
                            pZAware = CType(pPolylineErr, IZAware)
                            pZAware.DropZs()
                            pZAware.ZAware = False
                            'Ajouter les EDGES trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(CType(pPolylineErr, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & sMessage & " /Précision=" & gsExpression & " /Attribut_exclut=" & gsExclureAttribut, _
                                                pPolylineErr, pEnumTopoParent.Count)
                        End If

                        'Vider la mémoire
                        pTopoEdge = Nothing
                        pEnumTopoParent = Nothing
                        pZAware = Nothing
                        pPolylineErr = Nothing
                        pEsriTopoParent = Nothing
                        pFeatureClass = Nothing
                        pFeatureAdj = Nothing
                        sMessage = Nothing
                        'Récupération de la mémoire disponible
                        GC.Collect()

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
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pZAware = Nothing
            pPolylineErr = Nothing
            pEsriTopoParent = Nothing
            pFeatureClass = Nothing
            pFeatureAdj = Nothing
            sMessage = Nothing
            bAjouter = Nothing
            bFusion = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la segmentation présente entre deux lignes adjacentes en une point respecte les paramètres spécifiés. 
    ''' La segmentation est valide lorsqu'il y a les conditions suivantes :
    '''    -Lorsque deux éléments d'une même classe de type ligne sont adjacents en un point
    '''    -lorsqu'au moins une valeur d'attribut est différente entre les deux éléments
    '''     ou lorsqu'il y a plus de deux éléments adjacents.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterAdjacenceSegmentationLigne(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel, _
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant la géométrie d'un l'élément en relation.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim sExclureAttribut As String = ""                 'Contient les noms d'attributs à exclure.

        Try
            'Définir la géométrie par défaut
            TraiterAdjacenceSegmentationLigne = New GeometryBag
            TraiterAdjacenceSegmentationLigne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterAdjacenceSegmentationLigne, IGeometryCollection)

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
            pFeatureSel.Clear()
            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Ajouter une virgule à la fin dans la liste des attributs à exlure pour corriger le problème de nom d'attributs partiels
            If gsExclureAttribut <> "" Then sExclureAttribut = gsExclureAttribut & ","

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Traiter l'ajacence de fusion pour un élément
                ElementSegmentationLigne(pFeature, sExclureAttribut, pTopologyGraph, pNewSelectionSet, pGeomResColl, bEnleverSelection)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
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
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter la segmentation d'un élément de type ligne et ses éléments en relation.
    ''' 
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="sExclureAttribut">Contient les noms d'attributs à exclure.</param>
    '''<param name="pTopologyGraph">Interface contenant la topologie.</param>
    '''<param name="pNewSelectionSet">'Interface pour sélectionner les éléments.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométrie des éléments trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''</summary>
    ''' 
    Private Sub ElementSegmentationLigne(ByRef pFeature As IFeature, ByRef sExclureAttribut As String, ByRef pTopologyGraph As ITopologyGraph4, _
                                                  ByRef pNewSelectionSet As ISelectionSet, ByRef pGeomResColl As IGeometryCollection, ByVal bEnleverSelection As Boolean)
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing            'Interface des points trouvés.
        Dim pMultiPointColl As IPointCollection = Nothing   'Interface des points trouvés.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne de l'élément.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface utilisé pour extraire les Nodes d'un élément.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node.
        Dim pPoint As IPoint = Nothing                      'Interface pour comparer les sommets.
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface utilisé pour extraire les Edges du node.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface utilisé pour extraire les éléments du Edge.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe d'un élément
        Dim pFeatureAdj As IFeature = Nothing               'Interface contenant l'élément adjacent.
        Dim sMessage As String = ""                         'Contient le message.

        Try
            'Initialiser l'ajout
            pMultiPoint = New Multipoint
            pMultiPoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            pMultiPointColl = CType(pMultiPoint, IPointCollection)

            'Définir la géométrie à traiter
            pPolyline = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Table, IFeatureClass), pFeature.OID), IPolyline)

            'Vérifier si la ligne de la topologie est présente
            If pPolyline IsNot Nothing Then
                'Vérifier si la ligne est non fermée
                If Not pPolyline.IsClosed Then
                    'Interface pour extraire les composantes
                    pEnumTopoNode = pTopologyGraph.GetParentNodes(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                    'Initialisation de l'extraction
                    pEnumTopoNode.Reset()
                    'Extraire le premier Node
                    pTopoNode = pEnumTopoNode.Next

                    'Traiter tous les Nodes
                    Do Until pTopoNode Is Nothing
                        'Vérifier si le nombre de Edges est 2
                        If pTopoNode.Degree = 2 Then
                            'Interface contenant le Node traité
                            pPoint = CType(pTopoNode.Geometry, IPoint)

                            'Vérifier si le Node correspond au premier ou au dernier sommet
                            If pPoint.Compare(pPolyline.FromPoint) = 0 Or pPoint.Compare(pPolyline.ToPoint) = 0 Then
                                'Initialiser l'ajout
                                sMessage = ""
                                'Interface pour extraire les Edges du Node
                                pEnumNodeEdge = pTopoNode.Edges(True)
                                'Initialisation de l'extraction des Edges du Node
                                pEnumNodeEdge.Reset()
                                'Extraire le premier Edge
                                pEnumNodeEdge.Next(pTopoEdge, True)

                                'Traiter tous les Edge
                                Do Until pTopoEdge Is Nothing
                                    'Contient les parents sélectionnés
                                    pEnumTopoParent = pTopoEdge.Parents
                                    'Initialiser l'extraction
                                    pEnumTopoParent.Reset()
                                    'Extraire le premier élément
                                    pEsriTopoParent = pEnumTopoParent.Next()

                                    'Traiter tous les élément
                                    Do Until pEsriTopoParent.m_pFC Is Nothing
                                        'Extraire la FeatureClass
                                        pFeatureClass = pEsriTopoParent.m_pFC
                                        'Extraire l'élément
                                        pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                                        'Comparer les attributs de l'élément à traiter avec celui à comparer
                                        sMessage = sMessage & ComparerAttributElement(pFeature, pFeatureAdj, sExclureAttribut)
                                        'Extraire le prochain élément
                                        pEsriTopoParent = pEnumTopoParent.Next()
                                    Loop

                                    'Extraire le prochain Edge
                                    pEnumNodeEdge.Next(pTopoEdge, True)
                                Loop

                                'Vérifier si les valeurs d'attributs sont les mêmes
                                If sMessage = "" Then
                                    'Ajouter le point d'adjacence
                                    pMultiPointColl.AddPoint(pPoint)
                                End If
                            End If
                        End If

                        'Extraire le prochain Node
                        pTopoNode = pEnumTopoNode.Next
                    Loop
                End If
            End If

            'Vérifier si on doit sélectionner l'élément
            If (pMultiPoint.IsEmpty And Not bEnleverSelection) Or (Not pMultiPoint.IsEmpty And bEnleverSelection) Then
                'Ajouter l'élément dans la sélection
                pNewSelectionSet.Add(pFeature.OID)

                'Vérifier l'absence des points trouvés
                If pMultiPoint.IsEmpty Then
                    'Ajouter le point de début
                    pMultiPointColl.AddPoint(pPolyline.FromPoint)
                    'Ajouter le point de fin
                    pMultiPointColl.AddPoint(pPolyline.ToPoint)
                    'Définir le message
                    sMessage = "Aucune fusion, plusieurs adjacents et/ou plusieurs différences entre les valeurs d'attributs."
                Else
                    'Définir le message
                    sMessage = "Fusion à effectuer, un seul adjacent et aucune différence entre les valeurs d'attributs."
                End If

                'Ajouter les points trouvés de l'élément sélectionné
                pGeomResColl.AddGeometry(pMultiPoint)

                'Écrire une erreur
                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage & " /Précision=" & gsExpression & " /Attribut_exclut=" & gsExclureAttribut, pMultiPoint)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pMultiPoint = Nothing
            pMultiPointColl = Nothing
            pPolyline = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pPoint = Nothing
            pEnumNodeEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pFeatureClass = Nothing
            pFeatureAdj = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de comparer tous les attributs qui ne sont pas exclut entre deux éléments.
    ''' 
    ''' Les attributs non étitables et les géométries sont toujours excluts de la comparaison.
    '''</summary>
    ''' 
    '''<param name="pFeature"> Interface contenant l'élément à traiter.</param>
    '''<param name="pFeatureRel"> Interface contenant l'élément à comparer.</param>
    '''<param name="sExclureAttribut"> Liste des nom d'attributs à exclure de la comparaison.</param>
    ''' 
    ''' <return> String contenant les différences entre les 2 éléments.</return>
    '''
    Private Function ComparerAttributElement(ByVal pFeature As IFeature, ByVal pFeatureRel As IFeature, ByVal sExclureAttribut As String) As String
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing        'Interface contenant les attribut des éléments à traiter.
        Dim pFieldsRel As IFields = Nothing     'Interface contenant les attribut des éléments à comparer.
        Dim iPosAtt As Integer = -1             'Contient la position de l'attribut en relation.

        'Définir la valeur par défaut
        ComparerAttributElement = ""

        Try
            'Interface pour traiter les attributs
            pFields = pFeature.Fields
            pFieldsRel = pFeatureRel.Fields

            'Traiter tous les attributs
            For i = 0 To pFields.FieldCount - 1
                'Vérifier si l'attribut est à exclure
                If Not sExclureAttribut.Contains(pFields.Field(i).Name & ",") Then
                    'Vérifier si l'attribut est éditable et n'est pas de type "Geometry"
                    If pFields.Field(i).Editable And Not pFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Then
                        'Définir la position de l'attribut de l'élément en relation
                        iPosAtt = pFeatureRel.Fields.FindField(pFields.Field(i).Name)

                        'Vérifier si l'attribut est présent de la FeatureClass en relation
                        If iPosAtt >= 0 Then
                            'Vérifier si la valeur de l'attribut est différent
                            If pFeature.Value(i).ToString() <> pFeatureRel.Value(iPosAtt).ToString() Then
                                'Ajouter la différence
                                ComparerAttributElement = ComparerAttributElement & "#" & pFields.Field(i).Name & "=" & pFeature.Value(i).ToString _
                                                                                  & "/" & pFeatureRel.Value(iPosAtt).ToString
                                'Sortir de la fonction
                                Exit Function
                            End If

                            'si l'attribut est absent de la FeatureClass en relation
                        Else
                            'Ajouter la différence
                            ComparerAttributElement = ComparerAttributElement & "#" & pFields.Field(i).Name & "/Aucun attribut correspond"
                            'Sortir de la fonction
                            Exit Function
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFields = Nothing
            pFieldsRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la fusion respecte les paramètres spécifiés.
    '''    La fusion est valide lorsqu'il y a les conditions suivantes :
    '''    -Lorsqu'il n'y a aucun élément en relation et si le nombre de Edges est inférieure ou égale à deux. 
    '''    Si ces conditions ne sont pas respectées, la segmentation est nécessaire afin de les rendre adjacents.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterAdjacenceFusionLigne(ByVal pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel, _
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant la géométrie d'un l'élément en relation.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.

        Try
            'Définir la géométrie par défaut
            TraiterAdjacenceFusionLigne = New GeometryBag
            TraiterAdjacenceFusionLigne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterAdjacenceFusionLigne, IGeometryCollection)

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
            pFeatureSel.Clear()
            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Traiter l'ajacence de segmentation pour un élément
                ElementFusionLigne(pFeature, pTopologyGraph, pNewSelectionSet, pGeomResColl, bEnleverSelection)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
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
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
        End Try
    End Function


    '''<summary>
    ''' Routine qui permet de traiter la fusion d'un éléments avec ses éléments en relation.
    ''' 
    '''<param name="pFeature">Interface contenant contenant l'élément à traiter.</param>
    '''<param name="pTopologyGraph">Interface contenant la topologie.</param>
    '''<param name="pNewSelectionSet">'Interface pour sélectionner les éléments.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométrie des éléments trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''</summary>
    ''' 
    Private Sub ElementFusionLigne(ByRef pFeature As IFeature, ByRef pTopologyGraph As ITopologyGraph4, _
                                   ByRef pNewSelectionSet As ISelectionSet, ByRef pGeomResColl As IGeometryCollection, ByVal bEnleverSelection As Boolean)
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing            'Interface des points trouvés.
        Dim pMultiPointColl As IPointCollection = Nothing   'Interface pour ajouter les points trouvés.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne de l'élément.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface utilisé pour extraire les Nodes d'un élément.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node.
        Dim pPoint As IPoint = Nothing                      'Interface pour comparer les sommets.
        Dim sMessage As String = ""                         'Contient le message.

        Try
            'Initialiser l'ajout
            sMessage = ""
            pMultiPoint = New Multipoint
            pMultiPoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            pMultiPointColl = CType(pMultiPoint, IPointCollection)

            'Définir la géométrie à traiter
            pPolyline = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Table, IFeatureClass), pFeature.OID), IPolyline)

            'Vérifier si la ligne de la topologie est présente
            If pPolyline IsNot Nothing Then
                'Interface pour extraire les composantes
                pEnumTopoNode = pTopologyGraph.GetParentNodes(CType(pFeature.Table, IFeatureClass), pFeature.OID)

                'Vérifier la présence de noeud interne à la
                If (pPolyline.IsClosed And pEnumTopoNode.Count > 1) Or (Not pPolyline.IsClosed And pEnumTopoNode.Count > 2) Then
                    'Initialisation de l'extraction
                    pEnumTopoNode.Reset()

                    'Extraire le premier node
                    pTopoNode = pEnumTopoNode.Next

                    'Traiter tous les Nodes
                    Do Until pTopoNode Is Nothing
                        'Vérifier si le nombre de Edges est supérieur à 2
                        If pTopoNode.Degree > 2 Then
                            'Interface contenant le Node traité
                            pPoint = CType(pTopoNode.Geometry, IPoint)

                            'Vérifier si le Node est différent du premier et dernier sommet
                            If Not (pPoint.Compare(pPolyline.FromPoint) = 0 Or pPoint.Compare(pPolyline.ToPoint) = 0) Then
                                'Ajouter le point d'adjacence
                                pMultiPointColl.AddPoint(CType(pTopoNode.Geometry, IPoint))
                            End If
                        End If

                        'Extraire le prochain node
                        pTopoNode = pEnumTopoNode.Next
                    Loop
                End If
            End If

            'Vérifier si on doit sélectionner l'élément
            If (pMultiPoint.IsEmpty And Not bEnleverSelection) Or (Not pMultiPoint.IsEmpty And bEnleverSelection) Then
                'Ajouter l'élément dans la sélection
                pNewSelectionSet.Add(pFeature.OID)

                'Vérifier l'absence des points trouvés
                If pMultiPoint.IsEmpty Then
                    'Ajouter le point de début
                    pMultiPointColl.AddPoint(pPolyline.FromPoint)
                    'Ajouter le point de fin
                    pMultiPointColl.AddPoint(pPolyline.ToPoint)
                    'Définir le message
                    sMessage = "Fusion valide, aucune segmentation à effectuer."
                Else
                    'Définir le message
                    sMessage = "Fusion invalide, segmentation à effectuer, plusieurs adjacents présents."
                End If

                'Ajouter les points trouvés de l'élément sélectionné
                pGeomResColl.AddGeometry(pMultiPoint)

                'Écrire une erreur
                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage & " /Précision=" & gsExpression, pMultiPoint, pMultiPointColl.PointCount)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pMultiPoint = Nothing
            pMultiPointColl = Nothing
            pPolyline = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pPoint = Nothing
        End Try
    End Sub
#End Region
End Class
