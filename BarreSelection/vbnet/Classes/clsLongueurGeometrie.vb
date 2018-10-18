Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsLongueurGeometrie.vb
'
'''<summary>
''' Classe qui permet de de sélectionner les éléments du FeatureLayer qui respecte ou non la longueur totale d’une géométrie, 
''' d’une composante de géométrie ou d’un segment d’une composante.
''' 
''' La classe permet de traiter les trois attributs de longueur de géométrie TOTAL, COMPOSANTE et SEGMENT.
''' 
''' TOTAL : La longueur totale de la géométrie.
''' COMPOSANTE : La longueur de chaque composante de la géométrie.
''' SEGMENT : La longueur de chaque segment d'une ligne ou d'un anneau.
''' 
''' Note : Une Polyline est composée de 1 à N ligne(s).
'''        Un Polygone est composé de 1 à N anneau(x) extérieur(s) et 0 à N anneau(x) intérieur(s).
'''        Les anneaux intérieurs sont liés obligatoirement à un anneau extérieur.       
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 20 avril 2015
'''</remarks>
''' 
Public Class clsLongueurGeometrie
    Inherits clsValeurAttribut

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            gsNomAttribut = "TOTAL"
            gsExpression = "100"

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
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

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

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de retourner le nom de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides ReadOnly Property Nom() As String
        Get
            Nom = "LongueurGeometrie"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides Property Parametres() As String
        Get
            Parametres = gsNomAttribut + " " + gsExpression
        End Get
        Set(ByVal value As String)
            gsNomAttribut = value.Split(CChar(" "))(0)
            gsExpression = value.Split(CChar(" "))(1)

            'Si l'action est EDGE
            If gsNomAttribut = "EDGE" Then
                'Les relations doit être présente
                gpFeatureLayersRelation = New Collection
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
                'Définir le paramètre pour trouver les longueurs totale des géométries
                ListeParametres.Add("EDGE 100")
                'Définir le paramètre pour trouver les longueurs totale des géométries
                ListeParametres.Add("TOTAL 100")
                'Définir le paramètre pour trouver les longueurs des composantes de géométries
                ListeParametres.Add("COMPOSANTE 10")
                'Définir le paramètre pour trouver les longueurs des segments de composantes
                ListeParametres.Add("SEGMENT 3")
                'Définir le paramètre pour trouver les longueurs totale des géométries
                ListeParametres.Add("TOTAL \d\d\d\.")
                'Définir le paramètre pour trouver les longueurs des composantes de géométries
                ListeParametres.Add("COMPOSANTE \d\d\.")
                'Définir le paramètre pour trouver les longueurs des segments de composantes
                ListeParametres.Add("SEGMENT ^[3-9]\.|\d\d\.")
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
                    gsMessage = "La contrainte est valide."
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
            If gsNomAttribut = "EDGE" Or gsNomAttribut = "TOTAL" Or gsNomAttribut = "COMPOSANTE" Or gsNomAttribut = "SEGMENT" Then
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des géométries respecte ou non l'expression régulière spécifiée.
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
        Dim pFeatureSel As IFeatureSelection = Nothing              'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing                'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                            'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing              'Interface utilisé pour extraire les éléments à traiter.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing             'Interface contenant la topologie.
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

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Si le nom de l'attribut est TOTAL
            If gsNomAttribut = "EDGE" Then
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

                'Afficher le message d'identification des points trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection selon la longueur des Edges, LongMin=" & gsExpression & " ..."
                'Traiter le FeatureLayer
                Selectionner = TraiterLongueurEdgeMinimum(pFeatureCursor, pTopologyGraph, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est COMPOSANTE
            ElseIf gsNomAttribut = "TOTAL" Then
                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurTotaleMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurTotale(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est COMPOSANTE
            ElseIf gsNomAttribut = "COMPOSANTE" Then
                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurComposanteMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurComposante(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est SEGMENT
            ElseIf gsNomAttribut = "SEGMENT" Then
                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurSegmentMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterLongueurSegment(pFeatureCursor, pTrackCancel, bEnleverSelection)
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des Edges respecte ou non la longueur minimale spécifiée.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterLongueurEdgeMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolyline As IPolyline = Nothing                'Interface pour extraire la longueur.
        Dim pLigneErreur As IPolyline = Nothing             'Interface contenant une ligne en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne de la polyligne.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier si une ligne touche le découpage.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dLongMin As Double = 0                          'Contient la longueur minimum.
        Dim bErreur As Boolean = False

        Try
            'Définir la géométrie par défaut
            TraiterLongueurEdgeMinimum = New GeometryBag
            TraiterLongueurEdgeMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurEdgeMinimum, IGeometryCollection)

            'Définir la longueur minimum
            dLongMin = ConvertDBL(gsExpression)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire la longueur
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    pPolyline.Project(TraiterLongueurEdgeMinimum.SpatialReference)

                    'Vérifier si les lignes de la géométrie de l'élément ont été corrigées
                    If CorrigerLongueurLignesElement(pTopologyGraph, pFeature, dLongMin, pLigneErreur) Then
                        'Une erreur a été trouvée
                        bSucces = False
                        'Interface pour extraire les lignes de la Polyligne
                        pGeomColl = CType(pLigneErreur, IGeometryCollection)
                    Else
                        'Aucune erreur trouvée
                        bSucces = True
                        'Interface pour extraire les lignes de la Polyligne
                        pGeomColl = CType(pLigneErreur, IGeometryCollection)
                    End If

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Aucune erreur par défaut
                        bErreur = False

                        'Traiter toutes les lignes de la Polyligne
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Définir la ligne
                            pPath = CType(pGeomColl.Geometry(i), IPath)
                            'Définir la polyligne
                            pPolyline = PathToPolyline(pPath)
                            pPolyline.SpatialReference = TraiterLongueurEdgeMinimum.SpatialReference

                            'Vérifier si la ligne de découpage est absente 
                            If gpLimiteDecoupage Is Nothing Then
                                'Indiquer la présence d'une erreur
                                bErreur = True
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolyline.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                    pPolyline, CSng(pPolyline.Length))

                                'Si la ligne de découpage est présente
                            Else
                                'Projeter la ligne en géographique
                                pPolyline.Project(gpLimiteDecoupage.SpatialReference)
                                'Interface pour vérifier si la ligne touche le découpage
                                pRelOp = CType(gpLimiteDecoupage, IRelationalOperator)
                                'Vérifier si la ligne touche le découpage
                                If pRelOp.Disjoint(pPolyline) Then
                                    'Indiquer la présence d'une erreur
                                    bErreur = True
                                    'Projeter la ligne
                                    pPolyline.Project(TraiterLongueurEdgeMinimum.SpatialReference)
                                    'Écrire une erreur
                                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolyline.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                        pPolyline, CSng(pPolyline.Length))
                                End If
                            End If

                            'Vérifier la présence d'une erreur
                            If bErreur Then
                                'Ajouter l'élément dans la sélection
                                pFeatureSel.Add(pFeature)
                                'Ajouter la géométrie de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pPolyline)
                            End If
                        Next
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pPolyline = Nothing
                    GC.Collect()

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
            pPolyline = Nothing
            pLigneErreur = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des lignes d'un élément selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerLongueurLignesElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dLongMin As Double, _
                                                   ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour enlever la partie de la géométrie en erreur.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerLongueurLignesElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Créer le Bag vide des lignes en erreur
                    pLigneErreur = New Polyline
                    pLigneErreur.SpatialReference = pFeature.Shape.SpatialReference

                    'Interface pour extraire les Edges de l'élément dans la topologie
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                    pEnumTopoEdge.Reset()

                    'Corriger la longueur des lignes de la géométrie de l'élément
                    CorrigerLongueurLignesElement = CorrigerLongueurLignesGeometrie(pTopologyGraph, pEnumTopoEdge, dLongMin, pLigneErreur)
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pEnumTopoEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des lignes d'une géométrie selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pEnumTopoEdge"> Interface contenant les Edges de la topologie de l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerLongueurLignesGeometrie(ByVal pTopologyGraph As ITopologyGraph, ByVal pEnumTopoEdge As IEnumTopologyEdge, ByVal dLongMin As Double, _
                                                     ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pTopoNodeInt As ITopologyNode = Nothing         'Interface contenant un noeud d'intersection. 
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans un Bag.
        Dim pLigneEdges As IPolyline = Nothing              'Interface contenant la ligne des Edges continus.
        Dim pCollEdges As Collection = Nothing              'Objet contenant une collection de Edges.

        'Par défaut, une modification a été effectuée
        CorrigerLongueurLignesGeometrie = False

        Try
            'Si les Edges de l'élément sont valident
            If pEnumTopoEdge IsNot Nothing Then
                'Si au moins un Edge est présent
                If pEnumTopoEdge.Count > 0 Then
                    'Interface pour ajouter les lignes en erreur
                    pGeomColl = CType(pLigneErreur, IGeometryCollection)

                    'Extraire la premier Edge
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter tous les Edges
                    Do Until pTopoEdge Is Nothing
                        'Vérifier si le Edge n'est pas traité
                        If Not pTopoEdge.Visited Then
                            'Vérifier si le Edge contient une fin de ligne 
                            If pTopoEdge.FromNode.Edges(True).Count = 1 Or pTopoEdge.ToNode.Edges(True).Count = 1 Then
                                'Si le noeud de début est une extrémité
                                If pTopoEdge.FromNode.Edges(True).Count = 1 Then
                                    'Définir la ligne des Edges continus
                                    Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt)

                                    'Si le noeud de fin est une extrémité
                                Else
                                    'Définir la ligne des Edges continus
                                    Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt)
                                End If

                                'Vérifier si la ligne est fermée ou si aucun noeud d'intersection n'est présent (ligne isolée)
                                If pLigneEdges.IsClosed Or pTopoNodeInt Is Nothing Then
                                    'Vérifier si la ligne Edge est plus petite que la longueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If

                                    'Si la ligne n'est pas fermée
                                Else
                                    'Traiter le noeud d'intersection
                                    Call TraiterNoeudIntersection(dLongMin, pTopologyGraph, pTopoNodeInt, pLigneEdges, pCollEdges)
                                End If

                                'Si le Edge ne contient pas une fin de ligne 
                            Else
                                'Définir la ligne des Edges continus
                                Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt, False)

                                'Vérifier si la ligne est fermée ou  si aucun noeud d'intersection n'est présent (ligne isolée)
                                If pLigneEdges.IsClosed Then
                                    'Vérifier si la ligne Edge est plus petite que la longueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If
                                End If
                            End If
                        End If

                        'Vérifier si on doit détruire le Edge
                        If pTopoEdge.IsSelected Then
                            'Détruire le Edge dans la topologie
                            pTopologyGraph.DeleteEdge(pTopoEdge)

                            'Ajouter la ligne du Edge en erreur
                            pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

                            'Indiquer qu'une correction a été effectuée
                            CorrigerLongueurLignesGeometrie = True
                        End If

                        'Extraire les Edges suivants
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoNodeInt = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pLigneEdges = Nothing
            pCollEdges = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter un noeud d'intersection afin de conserver l'extrémité la plus longue. 
    '''</summary>
    ''' 
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTopoNode"> Interface contenant le noeud d'intersection.</param>
    '''<param name="pLigneEdges"> Interface contenant la ligne des Edges continus.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    ''' 
    Private Sub TraiterNoeudIntersection(ByVal dLongMin As Double, ByVal pTopologyGraph As ITopologyGraph, ByVal pTopoNode As ITopologyNode, _
                                         ByVal pLigneEdges As IPolyline, ByVal pCollEdges As Collection)
        'Déclarer les variables de travail
        Dim pTopoNodeEdge As ITopologyEdge = Nothing        'Interface contenant un Edge d'un noeud de la topologie. 
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant la ligne des Edges continus du noeud d'intersection.
        Dim pLigneNodeEdges As IPolyline = Nothing          'Interface contenant les Edges d'un noeud.
        Dim pCollNodeEdges As Collection = Nothing          'Collection des Edges de la ligne continue du noeud d'intersection.
        Dim pTopoNodeLigne As ITopologyNode = Nothing       'Interface contenant un noeud de la topologie. 
        Dim iNbExt As Integer = 0                           'Contient le nombre d'extémités.
        Dim iNbEdge As Integer = 0                          'Contient le nombre d'éléments.

        Try
            'Vérifier si le noeud possède plus de 2 Edges
            If Not pTopoNode.Visited Then
                'Indiquer que le noeud a été visité
                pTopoNode.Visited = True

                'Vérifier si le noeud possède plus de 2 Edges
                If pTopoNode.Edges(True).Count > 2 Then
                    'Interface pour extraire les Edges du noeud
                    pEnumNodeEdges = pTopoNode.Edges(True)

                    'Extraire la premier Edge
                    pEnumNodeEdges.Next(pTopoNodeEdge, True)

                    'Traiter tous les Edges du noeud
                    Do Until pTopoNodeEdge Is Nothing
                        'Vérifier si le Edge est déjà traité
                        If Not pTopoNodeEdge.Visited Then
                            'Compter le nombre de Edges
                            iNbEdge = iNbEdge + 1

                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoNode, pTopoNodeEdge, pLigneNodeEdges, pCollNodeEdges, pTopoNodeLigne, False)

                            'Vérifier si aucun noeud d'intersection autre que celui de début n'a été trouvé
                            'Si la ligne est une extrémité
                            If pTopoNodeLigne Is Nothing Then
                                'Compter le nombre d'extrémités trouvées
                                iNbExt = iNbExt + 1

                                'Initialiser les noeuds visités
                                Call InitEdgesVisiter(pCollEdges, True)

                                'Vérifier si la ligne continue de départ est la plus longue
                                If pLigneEdges.Length > pLigneNodeEdges.Length Then
                                    'Vérifier si la ligne est inférieure à la logueur minimale
                                    If pLigneNodeEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollNodeEdges)
                                    End If

                                    'Si la ligne continue de départ n'est plus la plus longue
                                Else
                                    'Vérifier si la ligne est inférieure à la logueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If

                                    'Redéfinir la ligne continue la plus longue 
                                    pLigneEdges = pLigneNodeEdges
                                    'Redéfinir la collection des edges de la ligne continue la plus longue 
                                    pCollEdges = pCollNodeEdges
                                End If

                                'Si la ligne n'est pas une extrémité
                            Else
                                'Mettre l'attribut Visited=False pour tous les Edges
                                Call InitEdgesVisiter(pCollNodeEdges)
                            End If
                        End If

                        'Extraire le prochain Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)
                    Loop

                    'Vérifier si aucune autre extrémité n'a été trouvée
                    If iNbEdge > 0 And iNbExt = 0 And pLigneEdges.Length <= dLongMin Then
                        'Sélectionner les Edges de la topologie à détruire
                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdges = Nothing
            pTopoNodeEdge = Nothing
            pLigneNodeEdges = Nothing
            pCollNodeEdges = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'initialiser l'attribut "Visited=False" pour tous les Edges. 
    '''</summary>
    ''' 
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    '''<param name="bValeur"> Contient la valeur d'initialisation, Faux par défaut.</param>
    ''' 
    Private Sub InitEdgesVisiter(ByVal pCollEdges As Collection, Optional ByVal bValeur As Boolean = False)
        'Déclarer les variables de travail
        Dim pEdge As ITopologyEdge = Nothing        'Interface contenant un Edge de la topologie

        Try
            'Traiter tous les Edges de la collection
            For Each pEdge In pCollEdges
                'Mettre l'attribut Visited=False pour le Edge
                pEdge.Visited = bValeur
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEdge = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier tous les Edges d'une lignes continue. 
    '''</summary>
    ''' 
    '''<param name="pTopoNodeDebut"> Interface contenant le noeud d'intersection de début.</param>
    '''<param name="pTopoEdge"> Interface contenant le Edge à traiter.</param>
    '''<param name="pLigneEdges"> Interface contenant la ligne des Edges continus.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    '''<param name="pTopoNodeFin"> Interface contenant le noeud d'intersection de fin.</param>
    '''<param name="bVisiter"> Indiquer si on doit indiquer si le Edge a été visité.</param>
    '''<param name="bInitialiser"> Indiquer si on doit indiquer si le Edge a été visité.</param>
    ''' 
    Private Sub IdentifierLigneEdge(ByVal pTopoNodeDebut As ITopologyNode, ByVal pTopoEdge As ITopologyEdge, _
                                    ByRef pLigneEdges As IPolyline, ByRef pCollEdges As Collection, ByRef pTopoNodeFin As ITopologyNode,
                                    Optional ByVal bVisiter As Boolean = True, Optional ByVal bInitialiser As Boolean = True)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des lignes.
        Dim pTopoNodeEdge As ITopologyEdge = Nothing        'Interface contenant un Edge d'un noeud de la topologie. 
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant les Edges d'un noeud.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier la ligne continue.

        Try
            'Indiquer que le Edge a été traité
            If bVisiter Then pTopoEdge.Visited = True

            'Vérifier si on doit initialiser
            If bInitialiser Then
                pLigneEdges = Nothing
                pCollEdges = Nothing
                pTopoNodeFin = Nothing
            End If

            'Si la polyligne des Edges est invalide
            If pLigneEdges Is Nothing Then
                'Créer une polyligne vide
                pLigneEdges = New Polyline
                pLigneEdges.SpatialReference = pTopoEdge.Geometry.SpatialReference
            End If

            'Si la collection des Edges est invalide
            If pCollEdges Is Nothing Then
                'Créer une collection vide
                pCollEdges = New Collection
            End If

            'Ajouter le Edge dans la collection
            pCollEdges.Add(pTopoEdge)

            'Interface pour ajouter la ligne du Edge dans la polyligne
            pGeomColl = CType(pLigneEdges, IGeometryCollection)
            'Ajouter la ligne du Edge dans la polyligne
            pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

            'Vérifier si plusieurs composantes sont présentes dans la ligne
            If pGeomColl.GeometryCount > 1 Then
                'Simplifier la géométrie
                pTopoOp = CType(pLigneEdges, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

            'Vérifier si le Edge n'est pas fermé
            If Not pLigneEdges.IsClosed Then
                'Si le noeud de début est différent de celui de départ
                If Not pTopoNodeDebut.Equals(pTopoEdge.FromNode) Then
                    'Si le début du Edge se continu
                    If pTopoEdge.FromNode.Edges(True).Count = 2 Then
                        'Interface pour extraire les Edges du noeud
                        pEnumNodeEdges = pTopoEdge.FromNode.Edges(True)
                        pEnumNodeEdges.Reset()

                        'Extraire la premier Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)

                        'Vérifier si le Edge est le même que celui de départ
                        If pTopoEdge.Equals(pTopoNodeEdge) Then
                            'Extraire la prochain Edge
                            pEnumNodeEdges.Next(pTopoNodeEdge, True)
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)

                            'Si le Edge n'est pas le même que celui de départ
                        Else
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)
                        End If

                        'Si le noeud contient une intersection
                    ElseIf pTopoEdge.FromNode.Edges(True).Count > 2 Then
                        'Définir le noeud d'intersection
                        pTopoNodeFin = pTopoEdge.FromNode
                    End If
                End If

                'Si le noeud de fin est différent de celui de départ
                If Not pTopoNodeDebut.Equals(pTopoEdge.ToNode) Then
                    'Si la fin du Edge se continu
                    If pTopoEdge.ToNode.Edges(True).Count = 2 Then
                        'Interface pour extraire les Edges du noeud
                        pEnumNodeEdges = pTopoEdge.ToNode.Edges(True)
                        pEnumNodeEdges.Reset()

                        'Extraire la premier Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)

                        'Vérifier si le Edge est le même que celui de départ
                        If pTopoEdge.Equals(pTopoNodeEdge) Then
                            'Extraire la premier Edge
                            pEnumNodeEdges.Next(pTopoNodeEdge, True)
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)

                            'Si le Edge n'est pas le même que celui de départ
                        Else
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)
                        End If

                        'Si le noeud contient une intersection
                    ElseIf pTopoEdge.ToNode.Edges(True).Count > 2 Then
                        'Définir le noeud d'intersection
                        pTopoNodeFin = pTopoEdge.ToNode
                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pEnumNodeEdges = Nothing
            pTopoNodeEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner tous les Edges d'une lignes continue. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    ''' 
    Private Sub SelectionnerEdges(ByVal pTopologyGraph As ITopologyGraph, ByVal pCollEdges As Collection)
        'Déclarer les variables de travail
        Dim pEdge As ITopologyEdge = Nothing        'Interface contenant un Edge de la topologie

        Try
            'Traiter tous les Edges de la collection
            For Each pEdge In pCollEdges
                'Sélectionner le Edge
                pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pEdge)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEdge = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur totale respecte ou non la longueur minimale spécifiée.
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
    '''
    Private Function TraiterLongueurTotaleMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                  Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolyline As IPolyline = Nothing                'Interface pour extraire la longueur.
        Dim pPolygon As IPolygon = Nothing                  'Interface pour extraire la longueur.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dLongMin As Double = 0           'Contient la longueur minimum.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurTotaleMinimum = New GeometryBag
            TraiterLongueurTotaleMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurTotaleMinimum, IGeometryCollection)

            'Définir la longueur minimum
            dLongMin = ConvertDBL(gsExpression)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire la longueur
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    pPolyline.Project(TraiterLongueurTotaleMinimum.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression minimum
                    bSucces = pPolyline.Length >= dLongMin

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolyline)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolyline.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                            pPolyline, CSng(pPolyline.Length))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pPolyline = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est un Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire la longueur
                    pPolygon = CType(pFeature.ShapeCopy, IPolygon)
                    pPolygon.Project(TraiterLongueurTotaleMinimum.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression minimum
                    bSucces = pPolygon.Length >= dLongMin

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolygon.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                            PolygonToPolyline(pPolygon), CSng(pPolygon.Length))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pPolygon = Nothing
                    GC.Collect()

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
            pPolygon = Nothing
            pPolyline = Nothing
            bSucces = Nothing
            dLongMin = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des composantes respecte ou non la longueur minimale spécifiée.
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
    '''
    Private Function TraiterLongueurComposanteMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                      Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes.
        Dim pPath As IPath = Nothing                        'Interface pour extraire la longueur.
        Dim pRing As IRing = Nothing                        'Interface pour extraire la longueur.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dLongMin As Double = 0           'Contient la longueur minimum.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurComposanteMinimum = New GeometryBag
            TraiterLongueurComposanteMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurComposanteMinimum, IGeometryCollection)

            'Définir la longueur minimum
            dLongMin = ConvertDBL(gsExpression)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurComposanteMinimum.SpatialReference)

                    'Interface pour extraire la longueur
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante
                        pPath = CType(pGeomColl.Geometry(i), IPath)

                        'Valider la valeur d'attribut selon l'expression minimum
                        bSucces = pPath.Length >= dLongMin

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurComposanteMinimum.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne trouvée
                            pPolylineColl.AddGeometry(pPath)
                            'Ajouter la lignes trouvée
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPath.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                pPolyline, CSng(pPath.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pGeomColl = Nothing
                    pPath = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est un Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurComposanteMinimum.SpatialReference)

                    'Interface pour extraire la longueur
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante
                        pRing = CType(pGeomColl.Geometry(i), IRing)

                        'Valider la valeur d'attribut selon l'expression minimum
                        bSucces = pRing.Length >= dLongMin

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurComposanteMinimum.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne trouvée
                            pPolylineColl.AddGeometryCollection(CType(PathToPolyline(pRing), IGeometryCollection))
                            'Ajouter la ligne trouvée
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pRing.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                pPolyline, CSng(pRing.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pRing = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pGeomColl = Nothing
                    GC.Collect()

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
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pRing = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
            bSucces = Nothing
            dLongMin = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des segments respecte ou non la longueur minimale spécifiée.
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
    '''
    Private Function TraiterLongueurSegmentMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments.
        Dim pSegment As ISegment = Nothing                  'Interface pour extraire la longueur.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les segments trouvés.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dLongMin As Double = 0           'Contient la longueur minimum.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurSegmentMinimum = New GeometryBag
            TraiterLongueurSegmentMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurSegmentMinimum, IGeometryCollection)

            'Définir la longueur minimum
            dLongMin = ConvertDBL(gsExpression)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then _
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurSegmentMinimum.SpatialReference)

                    'Interface pour extraire les segments d'une géométrie
                    pSegColl = CType(pGeometry, ISegmentCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pSegColl.SegmentCount - 1
                        'Définir la composante
                        pSegment = CType(pSegColl.Segment(i), ISegment)

                        'Valider la valeur d'attribut selon l'expression minimum
                        bSucces = pSegment.Length >= dLongMin

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurSegmentMinimum.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne d'erreur
                            pPolylineColl.AddGeometryCollection(CType(LineToPolyline(CType(pSegment, ILine)), IGeometryCollection))
                            'Ajouter la ligne d'erreur
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pSegment.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                pPolyline, CSng(pSegment.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pSegColl = Nothing
                    pSegment = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est un polygone
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Définir la limite du polygone de découpage
                pLimiteDecoupage = LimiteDecoupage

                'Créer limite vide si elle est absente
                If pLimiteDecoupage Is Nothing Then pLimiteDecoupage = New Polyline

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface contenant la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurSegmentMinimum.SpatialReference)

                    'Interface pour extraire les segments d'une géométrie
                    pSegColl = CType(pGeometry, ISegmentCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pSegColl.SegmentCount - 1
                        'Définir la composante
                        pSegment = CType(pSegColl.Segment(i), ISegment)

                        'Valider la valeur d'attribut selon l'expression minimum
                        bSucces = pSegment.Length >= dLongMin

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurSegmentMinimum.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne d'erreur
                            pPolylineColl.AddGeometryCollection(CType(LineToPolyline(CType(pSegment, ILine)), IGeometryCollection))
                            'Ajouter la ligne d'erreur
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pSegment.Length.ToString("F3") & " /LongMin=" & dLongMin.ToString, _
                                                pPolyline, CSng(pSegment.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pSegColl = Nothing
                    pSegment = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pTopoOp = Nothing
                    GC.Collect()

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
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pSegColl = Nothing
            pSegment = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
            pTopoOp = Nothing
            pLimiteDecoupage = Nothing
            dLongMin = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur totale respecte ou non l'expression régulière spécifiée.
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
    '''
    Private Function TraiterLongueurTotale(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                           Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolyline As IPolyline = Nothing                'Interface pour extraire la longueur.
        Dim pPolygon As IPolygon = Nothing                  'Interface pour extraire la longueur.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurTotale = New GeometryBag
            TraiterLongueurTotale.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurTotale, IGeometryCollection)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire la longueur
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    pPolyline.Project(TraiterLongueurTotale.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pPolyline.Length.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolyline)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolyline.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                            pPolyline, CSng(pPolyline.Length))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pPolyline = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est un Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire la longueur
                    pPolygon = CType(pFeature.ShapeCopy, IPolygon)
                    pPolygon.Project(TraiterLongueurTotale.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pPolygon.Length.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPolygon.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                            PolygonToPolyline(pPolygon), CSng(pPolygon.Length))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pPolygon = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pPolygon = Nothing
            pPolyline = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des composantes respecte ou non l'expression régulière spécifiée.
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
    '''
    Private Function TraiterLongueurComposante(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                               Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes.
        Dim pPath As IPath = Nothing                        'Interface pour extraire la longueur.
        Dim pRing As IRing = Nothing                        'Interface pour extraire la longueur.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurComposante = New GeometryBag
            TraiterLongueurComposante.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurComposante, IGeometryCollection)

            'Vérifier si la géométrie est une Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurComposante.SpatialReference)

                    'Interface pour extraire la longueur
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante
                        pPath = CType(pGeomColl.Geometry(i), IPath)

                        'Valider la valeur d'attribut selon l'expression régulière
                        oMatch = oRegEx.Match(pPath.Length.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurComposante.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne d'erreur
                            pPolylineColl.AddGeometry(pPath)
                            'Ajouter la ligne d'erreur
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pPath.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                pPolyline, CSng(pPath.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pGeomColl = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pPath = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Si la géométrie est un Polygon
            ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurComposante.SpatialReference)

                    'Interface pour extraire la longueur
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante
                        pRing = CType(pGeomColl.Geometry(i), IRing)

                        'Valider la valeur d'attribut selon l'expression régulière
                        oMatch = oRegEx.Match(pRing.Length.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurComposante.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne d'erreur
                            pPolylineColl.AddGeometryCollection(CType(PathToPolyline(pRing), IGeometryCollection))
                            'Ajouter la ligne d'erreur
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pRing.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                pPolyline, CSng(pRing.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pGeomColl = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pRing = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pRing = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la longueur des segments respecte ou non l'expression régulière spécifiée.
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
    '''
    Private Function TraiterLongueurSegment(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                            Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments.
        Dim pSegment As ISegment = Nothing                  'Interface pour extraire la longueur.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lines trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterLongueurSegment = New GeometryBag
            TraiterLongueurSegment.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterLongueurSegment, IGeometryCollection)

            'Vérifier si la géométrie est une Polyline ou un polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Extraire la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterLongueurSegment.SpatialReference)

                    'Interface pour extraire la longueur
                    pSegColl = CType(pGeometry, ISegmentCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pSegColl.SegmentCount - 1
                        'Définir la composante
                        pSegment = CType(pSegColl.Segment(i), ISegment)

                        'Valider la valeur d'attribut selon l'expression régulière
                        oMatch = oRegEx.Match(pSegment.Length.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Initialiser la ligne d'erreur
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterLongueurSegment.SpatialReference
                            pPolylineColl = CType(pPolyline, IGeometryCollection)
                            'Construire la ligne d'erreur
                            pPolylineColl.AddGeometryCollection(CType(LineToPolyline(CType(pSegment, ILine)), IGeometryCollection))
                            'Ajouter la ligne d'erreur
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Longueur=" & pSegment.Length.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                pPolyline, CSng(pSegment.Length))
                        End If
                    Next

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pSegColl = Nothing
                    pPolyline = Nothing
                    pPolylineColl = Nothing
                    pSegment = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pSegColl = Nothing
            pSegment = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
        End Try
    End Function
#End Region
End Class
