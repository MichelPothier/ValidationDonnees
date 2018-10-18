Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsSuperficieGeometrie.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non la superficie totale d’une géométrie, 
''' d’un anneau, d’un anneau extérieur ou d’un anneau intérieur..
''' 
''' La classe permet de traiter les trois attributs de superficie de géométrie TOTAL, ANNEAU, EXTERIEUR et INTERIEUR.
''' 
''' TOTAL : La superficie totale de la géométrie.
''' ANNEAU : La superficie de chaque anneau (extérieur ou intérieur) de la géométrie. 
''' EXTERIEUR : La superficie de chaque anneau extérieur.
''' INTERIEUR : La superficie de chaque anneau intérieur.
''' 
''' Note : Un Polygone est composé de 1 à N anneau(x) extérieur(s) et 0 à N anneau(x) intérieur(s).
'''        Les anneaux intérieurs sont liés obligatoirement à un anneau extérieur.     
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 20 avril 2015
'''</remarks>
''' 
Public Class clsSuperficieGeometrie
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
            NomAttribut = "TOTAL"
            Expression = "100"

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
            Nom = "SuperficieGeometrie"
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
                'Définir le paramètre pour trouver les superficies totale des géométries
                ListeParametres.Add("TOTAL 100")
                'Définir le paramètre pour trouver les superficies des anneaux (extérieurs ou intérieurs) de géométries
                ListeParametres.Add("ANNEAU 100")
                'Définir le paramètre pour trouver les superficies des anneaux extérieurs de géométries
                ListeParametres.Add("EXTERIEUR 100")
                'Définir le paramètre pour trouver les superficies des anneaux intérieurs de géométries
                ListeParametres.Add("INTERIEUR 100")
                'Définir le paramètre pour trouver les superficies totale des géométries
                ListeParametres.Add("TOTAL \d\d\d\.")
                'Définir le paramètre pour trouver les superficies des anneaux (extérieurs ou intérieurs) de géométries
                ListeParametres.Add("ANNEAU \d\d\d\.")
                'Définir le paramètre pour trouver les superficies des anneaux extérieurs de géométries
                ListeParametres.Add("EXTERIEUR \d\d\d\.")
                'Définir le paramètre pour trouver les superficies des anneaux intérieurs de géométries
                ListeParametres.Add("INTERIEUR \d\d\d\.")
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
                'Vérifier si la FeatureClass est de type Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'La contrainte est valide
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide"
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polygon."
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
            If gsNomAttribut = "TOTAL" Or gsNomAttribut = "ANNEAU" Or gsNomAttribut = "EXTERIEUR" Or gsNomAttribut = "INTERIEUR" Then
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des géométries respecte ou non l'expression régulière spécifiée.
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

            'Si le nom de l'attribut est TOTAL
            If gsNomAttribut = "TOTAL" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieTotaleMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieTotale(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est ANNEAU
            ElseIf gsNomAttribut = "ANNEAU" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieAnneauMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieAnneau(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est EXTERIEUR
            ElseIf gsNomAttribut = "EXTERIEUR" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieExterieurMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieExterieur(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est INTERIEUR
            ElseIf gsNomAttribut = "INTERIEUR" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si l'expression est numérique
                If TestDBL(gsExpression) Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieInterieurMinimum(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSuperficieInterieur(pFeatureCursor, pTrackCancel, bEnleverSelection)
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
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie totale respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieTotaleMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                    Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon = Nothing                  'Interface pour projeter.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dSupMin As Double = 0       'Contient la superficie minimum.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieTotaleMinimum = New GeometryBag
            TraiterSuperficieTotaleMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieTotaleMinimum, IGeometryCollection)

            'Définir la superficie minimum
            dSupMin = ConvertDBL(gsExpression)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pPolygon = CType(pFeature.Shape, IPolygon)
                    pPolygon.Project(TraiterSuperficieTotaleMinimum.SpatialReference)

                    'Interface pour extraire la superficie
                    pArea = CType(pPolygon, IArea)
                    'Valider la valeur d'attribut selon l'expression minimum
                    bSucces = pArea.Area >= dSupMin

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & pArea.Area.ToString("F3") & " /SupMin=" & dSupMin.ToString, _
                                            pPolygon, CSng(pArea.Area))
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
            pPolygon = Nothing
            pArea = Nothing
            dSupMin = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux (extérieurs ou intérieurs) respecte ou non l'expression régulière spécifiée.
    '''</summary>
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
    Private Function TraiterSuperficieAnneauMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                    Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes.
        Dim pRing As IRing = Nothing                        'Interface contenant un anneau.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier si l'anneau est disjoit de la limite de découpage.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim bDisjoint As Boolean = False                    'Indique si l'anneau est disjoint de la limite de découpage.
        Dim dSupMin As Double = 0       'Contient la superficie minimum.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieAnneauMinimum = New GeometryBag
            TraiterSuperficieAnneauMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieAnneauMinimum, IGeometryCollection)

            'Définir la superficie minimum
            dSupMin = ConvertDBL(gsExpression)

            'Si la géométrie est un Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = TraiterSuperficieAnneauMinimum.SpatialReference
                    pPolylineColl = CType(pPolyline, IGeometryCollection)

                    'Interface contenant la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Interface pour extraire les anneaux
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante par défaut
                        pRing = CType(pGeomColl.Geometry(i), IRing)

                        'L'anneau est disjoint du découpage par défaut
                        bDisjoint = True
                        'Vérifier si la limite de découpage est spécifiée
                        If gpLimiteDecoupage IsNot Nothing Then
                            'Interface pour vérifier si l'anneau touche la limite
                            pRelOp = CType(gpLimiteDecoupage, IRelationalOperator)
                            'Définir la ligne de l'anneau
                            pPolyline = PathToPolyline(pRing)
                            'Vérifier si l'anneau ne touche pas à la limite du découpage
                            bDisjoint = pRelOp.Disjoint(pPolyline)
                        End If

                        'Projeter l'anneau
                        pRing.Project(TraiterSuperficieAnneauMinimum.SpatialReference)
                        'Interface pour extraire la superficie
                        pArea = CType(pRing, IArea)
                        'Vérifier si l'anneau respecte la superficie minimale
                        bSucces = Not (Math.Abs(pArea.Area) < dSupMin And bDisjoint)

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Ajouter la ligne trouvée
                            pPolyline = PathToPolyline(pRing)
                            'Ajouter les lignes trouvées
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & Math.Abs(pArea.Area).ToString & " /SupMin=" & dSupMin.ToString, _
                                                pPolyline, CSng(Math.Abs(pArea.Area)))
                        End If
                    Next

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

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
            pArea = Nothing
            pRing = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
            pRelOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux extérieurs respecte ou non l'expression régulière spécifiée.
    '''</summary>
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
    Private Function TraiterSuperficieExterieurMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dSupMin As Double = 0       'Contient la superficie minimum.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieExterieurMinimum = New GeometryBag
            TraiterSuperficieExterieurMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieExterieurMinimum, IGeometryCollection)

            'Définir la superficie minimum
            dSupMin = ConvertDBL(gsExpression)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les anneaux extérieurs
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Projeter le polygon
                    pPolygon.Project(TraiterSuperficieExterieurMinimum.SpatialReference)

                    'Interface pour extraire les anneaux extérieurs
                    pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Définir la composante
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                        'Interface pour extraire la superficie
                        pArea = CType(pRingExt, IArea)
                        'Valider la valeur d'attribut selon l'expression minimum
                        bSucces = pArea.Area >= dSupMin

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Ajouter la ligne trouvée
                            pPolyline = PathToPolyline(pRingExt)
                            'Ajouter les lignes trouvées
                            pGeomSelColl.AddGeometry(pPolyline)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & pArea.Area.ToString("F3") & " /SupMin=" & dSupMin.ToString, _
                                                pPolyline, CSng(pArea.Area))
                        End If
                    Next

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

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
            pGeomCollExt = Nothing
            pArea = Nothing
            pRingExt = Nothing
            pPolyline = Nothing
            dSupMin = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux intérieurs respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieInterieurMinimum(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pRingInt As IRing = Nothing                     'Interface contenant l'anneau intérieur.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.
        Dim bSucces As Boolean = False                      'Indique si l'expression est un succès.
        Dim dSupMin As Double = 0       'Contient la superficie minimum.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieInterieurMinimum = New GeometryBag
            TraiterSuperficieInterieurMinimum.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieInterieurMinimum, IGeometryCollection)

            'Définir la superficie minimum
            dSupMin = ConvertDBL(gsExpression)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les anneaux extérieurs et intérieurs
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Projeter le polygon
                    pPolygon.Project(TraiterSuperficieInterieurMinimum.SpatialReference)

                    'Interface pour extraire les anneaux extérieurs
                    pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                    'Traiter tous les anneaux extérieurs
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Définir l'anneau extérieur
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                        'Extraire tous les anneaux intérieurs
                        pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                        'Traiter tous les anneaux intérieurs
                        For j = 0 To pGeomCollInt.GeometryCount - 1
                            'Définir l'anneau intérieur
                            pRingInt = CType(pGeomCollInt.Geometry(j), IRing)

                            'Interface pour extraire la superficie
                            pArea = CType(pRingInt, IArea)
                            'Valider la valeur d'attribut selon l'expression minimum
                            bSucces = Math.Abs(pArea.Area) >= dSupMin

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter la ligne trouvée
                                pPolyline = PathToPolyline(pRingInt)
                                'Ajouter les lignes trouvées
                                pGeomSelColl.AddGeometry(pPolyline)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & Math.Abs(pArea.Area).ToString & " /SupMin=" & dSupMin.ToString, _
                                                    pPolyline, CSng(Math.Abs(pArea.Area)))
                            End If
                        Next j
                    Next i

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

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
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pArea = Nothing
            pRingExt = Nothing
            pRingInt = Nothing
            pPolyline = Nothing
            dSupMin = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie totale respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieTotale(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon = Nothing                  'Interface pour projeter.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieTotale = New GeometryBag
            TraiterSuperficieTotale.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieTotale, IGeometryCollection)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pPolygon = CType(pFeature.Shape, IPolygon)
                    pPolygon.Project(TraiterSuperficieTotale.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression régulière
                    pArea = CType(pPolygon, IArea)
                    oMatch = oRegEx.Match(pArea.Area.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter la géométrie de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & pArea.Area.ToString("F3") & " /ExpReg=" & gsExpression, _
                                            pPolygon, CSng(pArea.Area))
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
            pPolygon = Nothing
            pArea = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux (extérieurs ou intérieurs) respecte ou non 
    ''' l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieAnneau(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes.
        Dim pRing As IRing = Nothing                        'Interface contenant un anneau.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieAnneau = New GeometryBag
            TraiterSuperficieAnneau.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieAnneau, IGeometryCollection)

            'Si la géométrie est un Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = TraiterSuperficieAnneau.SpatialReference
                    pPolylineColl = CType(pPolyline, IGeometryCollection)

                    'Interface contenant la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie
                    pGeometry.Project(TraiterSuperficieAnneau.SpatialReference)

                    'Interface pour extraire les anneaux
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la composante
                        pRing = CType(pGeomColl.Geometry(i), IRing)

                        'Valider la valeur d'attribut selon l'expression régulière
                        pArea = CType(pRing, IArea)
                        oMatch = oRegEx.Match(Math.Abs(pArea.Area).ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Ajouter la ligne trouvée
                            pPolylineColl.AddGeometryCollection(CType(PathToPolyline(pRing), IGeometryCollection))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & Math.Abs(pArea.Area).ToString & " /ExpReg=" & gsExpression, _
                                                PathToPolyline(pRing), CSng(Math.Abs(pArea.Area)))
                        End If
                    Next

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter les lignes trouvées
                        pGeomSelColl.AddGeometry(pPolyline)
                    End If

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
            pGeomColl = Nothing
            pArea = Nothing
            pRing = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux extérieurs respecte ou non 
    ''' l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieExterieur(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieExterieur = New GeometryBag
            TraiterSuperficieExterieur.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieExterieur, IGeometryCollection)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = TraiterSuperficieExterieur.SpatialReference
                    pPolylineColl = CType(pPolyline, IGeometryCollection)

                    'Interface pour extraire les anneaux extérieurs
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Projeter le polygon
                    pPolygon.Project(TraiterSuperficieExterieur.SpatialReference)

                    'Interface pour extraire les anneaux extérieurs
                    pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                    'Traiter toutes les composantes
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Définir la composante
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                        'Valider la valeur d'attribut selon l'expression régulière
                        pArea = CType(pRingExt, IArea)
                        oMatch = oRegEx.Match(pArea.Area.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Ajouter la ligne trouvée
                            pPolylineColl.AddGeometryCollection(CType(PathToPolyline(pRingExt), IGeometryCollection))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & pArea.Area.ToString("F3") & " /ExpReg=" & gsExpression, _
                                                PathToPolyline(pRingExt), CSng(pArea.Area))
                        End If
                    Next

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter les lignes trouvées
                        pGeomSelColl.AddGeometry(pPolyline)
                    End If

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
            pGeomCollExt = Nothing
            pArea = Nothing
            pRingExt = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la superficie des anneaux intérieurs respecte ou non 
    ''' l'expression régulière spécifiée.
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
    Private Function TraiterSuperficieInterieur(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pRingInt As IRing = Nothing                     'Interface contenant l'anneau intérieur.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes trouvés.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface contenant les lignes trouvés.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterSuperficieInterieur = New GeometryBag
            TraiterSuperficieInterieur.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSuperficieInterieur, IGeometryCollection)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = TraiterSuperficieInterieur.SpatialReference
                    pPolylineColl = CType(pPolyline, IGeometryCollection)

                    'Interface pour extraire les anneaux extérieurs et intérieurs
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Projeter le polygon
                    pPolygon.Project(TraiterSuperficieInterieur.SpatialReference)

                    'Interface pour extraire les anneaux extérieurs
                    pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                    'Traiter tous les anneaux extérieurs
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Définir l'anneau extérieur
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                        'Extraire tous les anneaux intérieurs
                        pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                        'Traiter tous les anneaux intérieurs
                        For j = 0 To pGeomCollInt.GeometryCount - 1
                            'Définir l'anneau intérieur
                            pRingInt = CType(pGeomCollInt.Geometry(j), IRing)

                            'Valider la valeur d'attribut selon l'expression régulière
                            pArea = CType(pRingInt, IArea)
                            oMatch = oRegEx.Match(Math.Abs(pArea.Area).ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter la ligne trouvée
                                pPolylineColl.AddGeometryCollection(CType(PathToPolyline(pRingInt), IGeometryCollection))
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Superficie=" & Math.Abs(pArea.Area).ToString & " /ExpReg=" & gsExpression, _
                                                    PathToPolyline(pRingInt), CSng(Math.Abs(pArea.Area)))
                            End If
                        Next j
                    Next i

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter les lignes trouvées
                        pGeomSelColl.AddGeometry(pPolyline)
                    End If

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
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pArea = Nothing
            pRingExt = Nothing
            pRingInt = Nothing
            pPolyline = Nothing
            pPolylineColl = Nothing
        End Try
    End Function
#End Region
End Class
