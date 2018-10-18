Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.IO

'**
'Nom de la composante : clsAngleGeometrie.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non l'angle en degré 
''' entre les segments consécutifs de la géométrie de ce dernier.
''' 
''' La classe permet de traiter les trois attributs d'angle de géométrie ANICROCHE, SAILLANT et RENTRANT.
''' 
''' ANICROCHE : L'angle d'anicroche minimum (0-180 degré) entre les segments consécutifs d'une géométrie.
''' 
''' SAILLANT : L'angle saillant (0-180 degré) entre les segments consécutifs d'une géométrie.
'''            
''' RENTRANT : L'angle rentrant (180-360 degré) entre les segments consécutifs d'une géométrie.
''' 
''' Note : Pour sélectionner les angles d'anicroches (+/- 0 degré).
'''        Pour sélectionner les angles droits (+/- 90 degré).
'''        Pour sélectionner les angles PLATS (+/- 180 degré).
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 21 avril 2015
'''</remarks>
''' 
Public Class clsAngleGeometrie
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
            NomAttribut = "ANICROCHE"
            Expression = "10"

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
            Nom = "AngleGeometrie"
        End Get
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
                'Définir le paramètre pour trouver les angles d'anicroche minimum
                ListeParametres.Add("ANICROCHE 10")
                'Définir le paramètre pour trouver les angles d'anicroche via les expressions régulières
                ListeParametres.Add("SAILLANT \d\d\.")
                'Définir le paramètre pour trouver les angles droits saillants
                ListeParametres.Add("SAILLANT ^(89|90)\.9")
                'Définir le paramètre pour trouver les angles droits rentrants
                ListeParametres.Add("RENTRANT ^(269|270)\.9")
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
            If gsNomAttribut = "ANICROCHE" Or gsNomAttribut = "SAILLANT" Or gsNomAttribut = "RENTRANT" Then
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

            'Vérifier si le nom de l'attribut est ANICROCHE
            If gsNomAttribut = "ANICROCHE" Then
                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Retourner l'expression valide
                    ExpressionValide = True
                    gsMessage = "La contrainte est valide"
                    'Si l'expression n'est pas numérique
                Else
                    gsMessage = "ERREUR : L'expression n'est pas numérique."
                End If
            Else
                'Retourner l'expression valide
                ExpressionValide = MyBase.ExpressionValide()
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'angle des géométries respecte ou non l'expression régulière spécifiée.
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

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPoint)

            'Si le nom de l'attribut est ANICROCHE
            If gsNomAttribut = "ANICROCHE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterAngleAnicroche(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est SAILLANT
            ElseIf gsNomAttribut = "SAILLANT" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterAngleSaillant(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est RENTRANT
            ElseIf gsNomAttribut = "RENTRANT" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterAngleRentrant(pFeatureCursor, pTrackCancel, bEnleverSelection)
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
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'angle saillant (0-180 degré) entre les segments consécutifs de la géométrie
    ''' respecte ou non l'angle d'anicroche minimum spécifiée.
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
    Private Function TraiterAngleAnicroche(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                           Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire l'angle d'un segment de base.
        Dim pLineCompare As ILine = Nothing                 'Interface pour extraire l'angle d'un segment à comparer.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire la limite de l'élément.
        Dim pLimiteDecoupage As IPolyline = Nothing         'Interface contenant la limite du polygone de découpage.
        Dim dAngleSaillant As Double = 0                    'Contient la valeur de l'angle saillant.
        Dim bSucces As Boolean = False                      'Indique si l'expression demandée est un succès.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim dAngleMin As Double = 0             'Contient l'angle minimum.

        Try
            'Définir la géométrie par défaut
            TraiterAngleAnicroche = New GeometryBag
            TraiterAngleAnicroche.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterAngleAnicroche, IGeometryCollection)

            'Définir l'angle minimum
            dAngleMin = ConvertDBL(gsExpression)

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
                    pGeometry.Project(TraiterAngleAnicroche.SpatialReference)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 2
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)
                            'Définir le segment consécutif à comparer
                            pLineCompare = CType(pSegColl.Segment(j + 1), ILine)
                            'Calculer l'angle saillant entre les deux segments consécutifs
                            dAngleSaillant = AngleSaillant(pLineBase, pLineCompare)
                            'Indiquer si l'angle d'anicroche minimum est un succès
                            bSucces = dAngleSaillant >= dAngleMin

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter le point de l'angle sélectionné
                                pGeomSelColl.AddGeometry(pLineBase.ToPoint)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Angle=" & dAngleSaillant.ToString("F3") & " /AngleMinimum=" & dAngleMin.ToString,
                                                    pLineBase.ToPoint, CSng(dAngleSaillant))
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

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    'Interface pour extraire la limite de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire la limite de la géométrie
                    pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                    'Enlever la partie commune avec la limite du polygone de découpage
                    pGeometry = pTopoOp.Difference(pLimiteDecoupage)
                    'Projeter la géométrie
                    pGeometry.Project(TraiterAngleAnicroche.SpatialReference)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 2
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)
                            'Définir le segment consécutif à comparer
                            pLineCompare = CType(pSegColl.Segment(j + 1), ILine)
                            'Calculer l'angle saillant entre les deux segments consécutifs
                            dAngleSaillant = AngleSaillant(pLineBase, pLineCompare)
                            'Indiquer si l'angle d'anicroche minimum est un succès
                            bSucces = dAngleSaillant >= dAngleMin

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter le point de l'angle sélectionné
                                pGeomSelColl.AddGeometry(pLineBase.ToPoint)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & dAngleSaillant.ToString("F3") & " /AngleMinimum=" & dAngleMin.ToString,
                                                    pLineBase.ToPoint, CSng(dAngleSaillant))
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
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pLineCompare = Nothing
            pTopoOp = Nothing
            pLimiteDecoupage = Nothing
            dAngleMin = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'angle saillant (0-180 degré) entre les segments consécutifs de la géométrie
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
    '''
    Private Function TraiterAngleSaillant(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                          Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire l'angle d'un segment de base.
        Dim pLineCompare As ILine = Nothing                 'Interface pour extraire l'angle d'un segment à comparer.
        Dim dAngleSaillant As Double = 0                    'Contient la valeur de l'angle saillant.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterAngleSaillant = New GeometryBag
            TraiterAngleSaillant.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterAngleSaillant, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(TraiterAngleSaillant.SpatialReference)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 2
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)
                            'Définir le segment consécutif à comparer
                            pLineCompare = CType(pSegColl.Segment(j + 1), ILine)
                            'Calculer l'angle saillant entre les deux segments consécutifs
                            dAngleSaillant = AngleSaillant(pLineBase, pLineCompare)

                            'Valider la valeur d'attribut selon l'expression régulière
                            oMatch = oRegEx.Match(dAngleSaillant.ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter le point de l'angle sélectionné
                                pGeomSelColl.AddGeometry(pLineBase.ToPoint)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Angle=" & dAngleSaillant.ToString("F3") & " /ExpReg=" & gsExpression,
                                                    pLineBase.ToPoint, CSng(dAngleSaillant))
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
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pLineCompare = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'angle rentrant (180-360 degré) entre les segments consécutifs de la géométrie
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
    '''
    Private Function TraiterAngleRentrant(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                          Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments d'une composante.
        Dim pLineBase As ILine = Nothing                    'Interface pour extraire l'angle d'un segment de base.
        Dim pLineCompare As ILine = Nothing                 'Interface pour extraire l'angle d'un segment à comparer.
        Dim dAngleRentrant As Double = 0                    'Contient la valeur de l'angle rentrant (180-360).
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterAngleRentrant = New GeometryBag
            TraiterAngleRentrant.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterAngleRentrant, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour projeter
                    pGeometry = pFeature.Shape
                    pGeometry.Project(TraiterAngleRentrant.SpatialReference)

                    'Interface pour extraire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Interface pour extraire les segments
                        pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)

                        'Traiter tous les angles entre les segments consécutifs
                        For j = 0 To pSegColl.SegmentCount - 2
                            'Définir le segment de base
                            pLineBase = CType(pSegColl.Segment(j), ILine)
                            'Définir le segment consécutif à comparer
                            pLineCompare = CType(pSegColl.Segment(j + 1), ILine)
                            'Calculer l'angle rentrant entre les deux segments consécutifs
                            dAngleRentrant = 360 - AngleSaillant(pLineBase, pLineCompare)

                            'Valider la valeur d'attribut selon l'expression régulière
                            oMatch = oRegEx.Match(dAngleRentrant.ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter le point de l'angle sélectionné
                                pGeomSelColl.AddGeometry(pLineBase.ToPoint)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Angle=" & dAngleRentrant.ToString("F3") & " /ExpReg=" & gsExpression,
                                                    pLineBase.ToPoint, CSng(dAngleRentrant))
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
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLineBase = Nothing
            pLineCompare = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de retourner l'angle saillant (0-180 degré) entre deux segments.
    '''</summary>
    '''
    '''<param name="pLineBase">Interface ESRI contenant le segment de base.</param>
    '''<param name="pLineCompare">Interface ESRI contenant le segment à comparer.</param> 
    '''
    '''<returns>Retourne l'angle saillant entre deux segments.</returns>
    '''
    Private Function AngleSaillant(ByVal pLineBase As ILine, ByVal pLineCompare As ILine) As Double
        'Déclarer les variables de travail
        Dim dAngleBase As Double    'Angle du segment de base.
        Dim dAngleCompare As Double 'Angle du segment à comparer.

        Try
            'Définir l'angle de la ligne de base
            dAngleBase = (pLineBase.Angle * 360 / (2 * Math.PI))

            'Définir l'angle de la ligne à comparer
            dAngleCompare = (pLineCompare.Angle * 360 / (2 * Math.PI))

            'Calculer l'angle saillant entre les deux segments
            AngleSaillant = dAngleBase - dAngleCompare + 180

            'Si l'angle dépasse la limite supérieure
            If AngleSaillant >= 360 Then
                'Soustraire la limite supérieure
                AngleSaillant = AngleSaillant - 360

                'Si l'angle dépasse la limite inférieure
            ElseIf AngleSaillant < 0 Then
                'Additionner la limite supérieure
                AngleSaillant = AngleSaillant + 360
            End If

            'Si l'angle est rentrant
            If AngleSaillant > 180 Then
                'Définir l'angle saillant
                AngleSaillant = 360 - AngleSaillant
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function
#End Region
End Class
