Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsCreerGeometrie.vb
'
'''<summary>
''' Classe qui permet de créer des nouvelles géométries complexes ou multiples selon diverses méthodes.
''' 
''' La classe permet de traiter les trois attributs de géométrie UNIQUE, INTERSECTE et ATTRIBUT.
''' 
''' UNIQUE : Créer une nouvelle géométrie unique (un seul élément contenant toutes les géométries) complexe ou multiple.
''' LIMITE : Créer une nouvelle géométrie pour chaque partie de limite ou pour l'ensemble des parties de limite d'une géométrie.
''' TROU : Créer une nouvelle géométrie pour chaque trou ou pour l'ensemble des trous d'un Polygone.
''' INTERSECTE : Créer des nouvelles géométries complexes ou multiples qui s'intersectent entre elles.
''' ATTRIBUT : Créer des nouvelles géométries complexes ou multiples qui possèdent les mêmes valeurs d'attributs spécifiées entre elles.
''' 
''' Note : Une géométrie complexe est une géométrie pour laquelle une simplification a été effectuée contrairement à une géométrie multiple. 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 07 Mars 2016
'''</remarks>
''' 
Public Class clsCreerGeometrie
    Inherits clsValeurAttribut
    '''<summary>Liste des attributs utilisée pour fusionner les éléments.</summary>
    Protected gsListeAttribut As String = ""

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "INTERSECTE"
            Expression = "VRAI"
            ListeAttribut = ""

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    '''</summary>
    '''
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sNomAttribut"> Nom de l'attribut à traiter.</param>
    '''<param name="sExpression"> Expression régulière à traiter.</param>
    '''<param name="sListeAttribut"> Liste des attributs utilisée pour fusionner les géométries.</param>
    ''' 
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String, Optional ByVal slisteAttribut As String = "")
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            ListeAttribut = slisteAttribut

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    '''</summary>
    '''
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
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
    '''</summary>
    '''
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
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

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gsListeAttribut = Nothing
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
            Nom = "CreerGeometrie"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la liste des attributs utilisée pour fusionner les géométries.
    '''</summary>
    ''' 
    Public Property ListeAttribut() As String
        Get
            ListeAttribut = gsListeAttribut
        End Get
        Set(ByVal value As String)
            gsListeAttribut = value
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
            If gsListeAttribut <> "" Then
                'Retourner aussi le masque spatial
                Parametres = Parametres & " " & gsListeAttribut
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
            If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION [LISTE_ATTRIBUT]")

            'Définir les valeurs par défaut
            gsNomAttribut = params(0).ToUpper
            gsExpression = params(1)

            'Vérifier si le troisième paramètre est présent
            If params.Length > 2 Then
                'Définir la liste des attributs utilisée pour fusionner les géométries
                gsListeAttribut = params(2)
            Else
                gsListeAttribut = ""
            End If
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    '''<returns> Collection contenant la liste des paramètres.</returns>
    ''' 
    Public Overloads Overrides Function ListeParametres() As Collection
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing            'Interface contenant les attributs de la FeatureClass.
        Dim sAttributs As String = ""               'Liste des attributs de la Featureclass de sélection.

        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir l'interface contenant les attributs de la Featureclass
                pFields = gpFeatureLayerSelection.FeatureClass.Fields

                'Traiter tous les attributs
                For i = 0 To pFields.FieldCount - 1
                    'vérifier si l'attribut est non éditable
                    If pFields.Field(i).Editable And pFields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry Then
                        'Vérifier si c'est le premier attribut
                        If sAttributs = "" Then
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = pFields.Field(i).Name
                        Else
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = sAttributs & "," & pFields.Field(i).Name
                        End If

                        'Définir le paramètre pour créer des géométries complexes qui possèdent les mêmes valeurs d'attributs entres elles.
                        ListeParametres.Add("ATTRIBUT VRAI " & pFields.Field(i).Name)
                    End If
                Next

                'Définir le paramètre pour créer des géométries complexes qui possèdent les mêmes valeurs d'attributs entres elles.
                ListeParametres.Add("ATTRIBUT VRAI " & sAttributs)
                'Définir le paramètre pour créer des géométries multiples qui possèdent les mêmes valeurs d'attributs entres elles.
                ListeParametres.Add("ATTRIBUT FAUX " & sAttributs)

                'Définir le paramètre pour créer des géométries pour chaque partie de limite d'une géométrie.
                ListeParametres.Add("LIMITE VRAI")
                'Définir le paramètre pour créer des géométries pour l'ensemble des parties de limite d'une géométrie.
                ListeParametres.Add("LIMITE FAUX")

                'Définir le paramètre pour créer des géométries pour chaque trou d'un Polygone.
                ListeParametres.Add("TROU VRAI")
                'Définir le paramètre pour créer des géométries pour l'ensemble des trous d'un Polygone.
                ListeParametres.Add("TROU FAUX")

                'Définir le paramètre pour créer des géométries complexes qui s'intersectent entre elles.
                ListeParametres.Add("INTERSECTE VRAI")
                'Définir le paramètre pour créer des géométries multiples qui s'intersectent entre elles.
                ListeParametres.Add("INTERSECTE FAUX")

                'Définir le paramètre pour créer une géométrie complexe unique.
                ListeParametres.Add("UNIQUE VRAI")
                'Définir le paramètre pour créer une géométrie multiple unique.
                ListeParametres.Add("UNIQUE FAUX")
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
    ''' Routine qui permet d'indiquer si l'attribut est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    Public Overloads Overrides Function AttributValide() As Boolean
        Try
            'La contrainte est valide par défaut
            AttributValide = True
            gsMessage = "La contrainte est valide."

            'Vérifier si l'attribut est valide
            If gsNomAttribut = "UNIQUE" Or gsNomAttribut = "TROU" Or gsNomAttribut = "INTERSECTE" Or gsNomAttribut = "ATTRIBUT" Then
                'Si l'attribut est intersecte et que la géométrie est un point
                If gsNomAttribut = "INTERSECTE" And gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    'Définir la valeur par défaut, la contrainte est invalide.
                    AttributValide = False
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : L'attribut INTERSECTE est invalide pour le type de géométrie Point."

                    'Si l'attribut est TROU et que la la géométrie n'est pas un Polygon
                ElseIf gsNomAttribut = "LIMITE" And (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                                                     Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) Then
                    'Définir la valeur par défaut, la contrainte est invalide.
                    AttributValide = False
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : Les types de géométrie Point ou MultiPoint sont invalides pour l'attribut LIMITE."

                    'Si l'attribut est TROU et que la la géométrie n'est pas un Polygon
                ElseIf gsNomAttribut = "TROU" And gpFeatureLayerSelection.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPolygon Then
                    'Définir la valeur par défaut, la contrainte est invalide.
                    AttributValide = False
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : Le type de géométrie Polygon est obligatoire pour l'attribut TROU."

                    'Si l'attribut est ATTRIBUT et que la  liste des attributs est absente
                ElseIf gsNomAttribut = "ATTRIBUT" And gsListeAttribut = "" Then
                    'Définir la valeur par défaut, la contrainte est invalide.
                    AttributValide = False
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : La liste des attributs ne doit pas être vide."

                    'Si l'attribut n'est pas ATTRIBUT et que la  liste des attributs est présente
                ElseIf gsNomAttribut <> "ATTRIBUT" And gsListeAttribut <> "" Then
                    'Définir la valeur par défaut, la contrainte est invalide.
                    AttributValide = False
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : La liste des attributs doit être vide."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont les géométries des éléments respectent ou non l'état VIDE, SIMPLE ou FERMER.
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

            'Si l'attribut est UNIQUE
            If gsNomAttribut = "UNIQUE" Then
                'Vérifier si le type de géométrie est un point
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                    'Si la géométrie à créer n'est pas de type point
                Else
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                End If

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si on doit fusionner
                If gsExpression = "VRAI" Then
                    'Traiter le FeatureLayer
                    Selectionner = TraiterCreerGeometrieUniqueVrai(pFeatureCursor, pTrackCancel, bEnleverSelection)
                Else
                    'Traiter le FeatureLayer
                    Selectionner = TraiterCreerGeometrieUniqueFaux(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si l'attribut est LIMITE
            ElseIf gsNomAttribut = "LIMITE" Then
                'Si la géométrie à créer n'est pas de type point
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)
                End If

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Vérifier si on doit créer un élément par partie de limites d'une géométrie d'élément
                If gsExpression = "VRAI" Then
                    'Vérifier si le type de géométrie est un point
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPoint)
                    End If

                    'Traiter le FeatureLayer
                    Selectionner = TraiterCreerGeometrieLimiteVrai(pFeatureCursor, pTrackCancel, bEnleverSelection)

                    'si on doit créer un élément pour l'ensemble des parties de limites d'une géométrie d'élément
                Else
                    'Vérifier si le type de géométrie est un point
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)
                    End If

                    'Traiter le FeatureLayer
                    Selectionner = TraiterCreerGeometrieLimiteFaux(pFeatureCursor, pTrackCancel, bEnleverSelection)
                End If

                'Si l'attribut est TROU
                ElseIf gsNomAttribut = "TROU" Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                    'Interfaces pour extraire les éléments sélectionnés
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)

                    'Vérifier si on doit créer un élément par trou d'un élément
                    If gsExpression = "VRAI" Then
                        'Traiter le FeatureLayer
                        Selectionner = TraiterCreerGeometrieTrouVrai(pFeatureCursor, pTrackCancel, bEnleverSelection)

                        'si on doit créer un élément pour tous les trous d'un élément
                    Else
                        'Traiter le FeatureLayer
                        Selectionner = TraiterCreerGeometrieTrouFaux(pFeatureCursor, pTrackCancel, bEnleverSelection)
                    End If

                    'Si l'attribut est INTERSECTE
                ElseIf gsNomAttribut = "INTERSECTE" Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                    'Traiter le FeatureLayer
                    Selectionner = TraiterCreerGeometrieIntersecte(pTrackCancel, bEnleverSelection)

                    'Si l'attribut est ATTRIBUT
                ElseIf gsNomAttribut = "ATTRIBUT" Then
                    'Vérifier si le type de géométrie est un point
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                        'Si la géométrie à créer n'est pas de type point
                    Else
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                    End If

                    'Vérifier si on doit fusionner
                    If gsExpression = "VRAI" Then
                        'Traiter le FeatureLayer
                        Selectionner = TraiterCreerGeometrieAttributVrai(pTrackCancel, bEnleverSelection)
                    Else
                        'Traiter le FeatureLayer
                        Selectionner = TraiterCreerGeometrieAttributFaux(pTrackCancel, bEnleverSelection)
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
    ''' Routine qui permet de créer une géométrie pour chaque partie de limite d'une géométrie.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieLimiteVrai(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pNewGeometry As IGeometry = Nothing             'Interface contenant la nouvelle géométrie de travail.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les parties des limites.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour extraire les limites de la géométrie de travail.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieLimiteVrai = New GeometryBag
            TraiterCreerGeometrieLimiteVrai.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieLimiteVrai, IGeometryCollection)

            'Si la géométrie est de type Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.Shape

                    'Interface pour extraire les limites de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)

                    'Interface pour extraire chaque partie des limites
                    pGeomColl = CType(pTopoOp.Boundary, IGeometryCollection)

                    'Vérifier si au moins une partie de limite est présente
                    If pGeomColl.GeometryCount > 0 Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)

                        'Vérifier si la géométrie originale est une surface
                        If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                            'Traiter toutes les lignes
                            For i = 0 To pGeomColl.GeometryCount - 1
                                'Créer la nouvelle géométrie correspondant à une Polyligne
                                pNewGeometry = PathToPolyline(CType(pGeomColl.Geometry(i), IPath))

                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pNewGeometry)

                                'Écrire une erreur de surface correspondant à un anneau intérieur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbLigne=" & pGeomColl.GeometryCount.ToString, pNewGeometry, pFeature.OID)
                            Next i

                            'si la géométrie originale est une ligne
                        ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                            'Traiter tous les points
                            For i = 0 To pGeomColl.GeometryCount - 1
                                'Créer la nouvelle géométrie correspondant à un Point
                                pNewGeometry = pGeomColl.Geometry(i)

                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pNewGeometry)

                                'Écrire une erreur de surface correspondant à un anneau intérieur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbPoint=" & pGeomColl.GeometryCount.ToString, pNewGeometry, pFeature.OID)
                            Next i
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pNewGeometry = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie pour l'ensemble des parties de limite d'une géométrie.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieLimiteFaux(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pNewGeometry As IGeometry = Nothing             'Interface contenant la nouvelle géométrie de travail.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les parties des limites.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour extraire les limites de la géométrie de travail.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieLimiteFaux = New GeometryBag
            TraiterCreerGeometrieLimiteFaux.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieLimiteFaux, IGeometryCollection)

            'Si la géométrie est de type Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.Shape

                    'Interface pour extraire les limites de la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    'Extraire les limites de la géométrie
                    pNewGeometry = pTopoOp.Boundary

                    'Interface pour extraire chaque partie des limites
                    pGeomColl = CType(pNewGeometry, IGeometryCollection)

                    'Vérifier si au moins une partie de limite est présente
                    If pGeomColl.GeometryCount > 0 Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)

                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pNewGeometry)

                        'Vérifier si la géométrie originale est une surface
                        If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                            'Écrire une erreur de surface correspondant à un anneau intérieur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbLigne=" & pGeomColl.GeometryCount.ToString, pNewGeometry, pFeature.OID)

                            'si la géométrie originale est une ligne
                        ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                            'Écrire une erreur de surface correspondant à un anneau intérieur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbPoint=" & pGeomColl.GeometryCount.ToString, pNewGeometry, pFeature.OID)
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pNewGeometry = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie selon chaque trou d'une surface.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieTrouVrai(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieTrouVrai = New GeometryBag
            TraiterCreerGeometrieTrouVrai.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieTrouVrai, IGeometryCollection)

            'Si la géométrie est de type Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Définir le Polygone de l'élément
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Vérifier si aucun polygone extérieur
                    If pPolygon.ExteriorRingCount > 0 Then
                        'Interface pour extraire les anneaux extérieurs
                        pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                        'Traiter tous les anneaux extérieurs
                        For i = 0 To pGeomCollExt.GeometryCount - 1
                            'Définir l'anneau extérieur
                            pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                            'Extraire tous les anneaux intérieurs de l'anneau extérieur
                            pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                            'Vérifier si au moins un anneau intérieur est présent
                            If pGeomCollInt.GeometryCount > 0 Then
                                'Ajouter l'élément dans la sélection
                                pFeatureSel.Add(pFeature)

                                'Traiter tous les anneaux intérieurs
                                For j = 0 To pGeomCollInt.GeometryCount - 1
                                    'Créer la nouvelle géométrie correspondant à l'anneau intérieur
                                    pGeometry = RingToPolygon(CType(pGeomCollInt.Geometry(j), IRing), New GeometryBag)

                                    'Ajouter l'enveloppe de l'élément sélectionné
                                    pGeomSelColl.AddGeometry(pGeometry)

                                    'Écrire une erreur de surface correspondant à un anneau intérieur
                                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbAnneauInt=" & pGeomCollInt.GeometryCount.ToString, pGeometry, pFeature.OID)
                                Next j
                            End If
                        Next i
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pPolygon = Nothing
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pRingExt = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie pour tous les trous d'une surface.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieTrouFaux(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieurs.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les anneaux intérieurs.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieTrouFaux = New GeometryBag
            TraiterCreerGeometrieTrouFaux.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieTrouFaux, IGeometryCollection)

            'Si la géométrie est de type Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Définir le Polygone de l'élément
                    pPolygon = CType(pFeature.Shape, IPolygon4)

                    'Vérifier si aucun polygone extérieur
                    If pPolygon.ExteriorRingCount > 0 Then
                        'Interface pour extraire les anneaux extérieurs
                        pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                        'Traiter tous les anneaux extérieurs
                        For i = 0 To pGeomCollExt.GeometryCount - 1
                            'Définir l'anneau extérieur
                            pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                            'Extraire tous les anneaux intérieurs de l'anneau extérieur
                            pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                            'Traiter tous les anneaux intérieurs
                            If pGeomCollInt.GeometryCount > 0 Then
                                'Créer la nouvelle géométrie correspondant aux anneaux intérieurs
                                pGeometry = New Polygon
                                pGeometry.SpatialReference = TraiterCreerGeometrieTrouFaux.SpatialReference

                                'Interface pour ajouter des anneaux
                                pGeomColl = CType(pGeometry, IGeometryCollection)

                                'Traiter tous les anneaux intérieurs
                                For j = 0 To pGeomCollInt.GeometryCount - 1
                                    'Ajouter un anneau intérieur
                                    pGeomColl.AddGeometry(pGeomCollInt.Geometry(j))
                                Next j

                                'Ajouter l'élément dans la sélection
                                pFeatureSel.Add(pFeature)

                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pGeometry)

                                'Écrire une erreur de surface correspondant à un anneau intérieur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbAnneauInt=" & pGeomCollInt.GeometryCount.ToString, pGeometry, pFeature.OID)
                            End If
                        Next i
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pPolygon = Nothing
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pGeomColl = Nothing
            pRingExt = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie unique (un seul élément contenant toutes les géométries) complexe. Les géométries sont fusionnées.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieUniqueVrai(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface contenant la géométrie unique.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter un point.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier la géométrie de travail.
        Dim iNbGeom As Integer = 0        'Compteur de géométries.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieUniqueVrai = New GeometryBag
            TraiterCreerGeometrieUniqueVrai.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieUniqueVrai, IGeometryCollection)

            'Si la géométrie est de type Point, MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Créer une nouvelle géométrie Multipoint vide
                pGeometry = New GeometryBag

                'Définir la référnce spatiale
                pGeometry.SpatialReference = TraiterCreerGeometrieUniqueVrai.SpatialReference
                pGeometry.SnapToSpatialReference()

                'Interface pour créer la géométrie unique
                pGeometryColl = CType(pGeometry, IGeometryCollection)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Compter les géométries
                    iNbGeom = iNbGeom + 1

                    'Vérifier si la géométrie est un point
                    If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Créer un nouveau multipoint
                        pGeometry = New Multipoint
                        pGeometry.SpatialReference = pFeature.Shape.SpatialReference
                        'Interface pour ajouter le Point
                        pPointColl = CType(pGeometry, IPointCollection)
                        'Ajouter le point
                        pPointColl.AddPoint(CType(pFeature.ShapeCopy, IPoint))

                        'Si la géométrie n'est pas un point
                    Else
                        'Interface pour projeter
                        pGeometry = pFeature.ShapeCopy
                    End If

                    'Projeter la géométrie
                    pGeometry.Project(TraiterCreerGeometrieUniqueVrai.SpatialReference)

                    'Ajouter la géométrie
                    pGeometryColl.AddGeometry(pGeometry)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Afficher le message de sélection des éléments trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Fusionner les géométries dans une seule géométrie ..."

                'Vérifier si le type de géométrie à créer est un Multipoint
                If gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                    'Créer une nouvelle géométrie Multipoint vide
                    pGeometry = New Multipoint

                    'Si le type de géométrie à créer est une Polyline
                ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    'Créer une nouvelle géométrie Polyline vide
                    pGeometry = New Polyline

                    'Si le type de géométrie à créer est un Polygon
                ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Créer une nouvelle géométrie Polygon vide
                    pGeometry = New Polygon
                End If

                'Définir la référnce spatiale
                pGeometry.SpatialReference = TraiterCreerGeometrieUniqueVrai.SpatialReference
                pGeometry.SnapToSpatialReference()

                'Interface pour fusionner la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                'Fusionner la géométrie
                pTopoOp.ConstructUnion(CType(pGeometryColl, IEnumGeometry))
                'Redéfinir la géométrie traitée
                pGeometryColl = CType(pTopoOp, IGeometryCollection)

                'Ajouter la géométrie unique
                pGeomSelColl.AddGeometry(CType(pGeometryColl, IGeometry))

                'Écrire une erreur
                EcrireFeatureErreur("OID=1 #NbGéométries=" & iNbGeom.ToString & " /NbComposantes=" & pGeometryColl.GeometryCount.ToString & " /" & gsNomAttribut & " " & gsExpression, CType(pGeometryColl, IGeometry), iNbGeom)

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometryColl = Nothing
            pPointColl = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie unique (un seul élément contenant toutes les géométries) multiple. Les géométries ne sont pas fusionnées.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterCreerGeometrieUniqueFaux(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface contenant la géométrie unique.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de travail.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier la géométrie de travail.
        Dim iNbGeom As Integer = 0        'Compteur de géométries.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieUniqueFaux = New GeometryBag
            TraiterCreerGeometrieUniqueFaux.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieUniqueFaux, IGeometryCollection)

            'Si la géométrie est de type Point, MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Enlever la sélection
                pFeatureSel.Clear()

                'Vérifier si le type de géométrie à créer est un Multipoint
                If gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                    'Créer une nouvelle géométrie Multipoint vide
                    pGeometry = New Multipoint

                    'Si le type de géométrie à créer est une Polyline
                ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    'Créer une nouvelle géométrie Polyline vide
                    pGeometry = New Polyline

                    'Si le type de géométrie à créer est un Polygon
                ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Créer une nouvelle géométrie Polygon vide
                    pGeometry = New Polygon
                End If

                'Définir la référnce spatiale
                pGeometry.SpatialReference = TraiterCreerGeometrieUniqueFaux.SpatialReference
                pGeometry.SnapToSpatialReference()

                'Interface pour créer la géométrie unique
                pGeometryColl = CType(pGeometry, IGeometryCollection)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Compter les géométries
                    iNbGeom = iNbGeom + 1

                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(TraiterCreerGeometrieUniqueFaux.SpatialReference)

                    'Vérifier si la géométrie est un point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Ajouter la géométrie
                        pGeometryColl.AddGeometry(pGeometry)

                        'Si la géométrie n'est pas un point
                    Else
                        'Ajouter la géométrie
                        pGeometryColl.AddGeometryCollection(CType(pGeometry, IGeometryCollection))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Ajouter la géométrie unique
                pGeomSelColl.AddGeometry(CType(pGeometryColl, IGeometry))

                'Écrire une erreur
                EcrireFeatureErreur("OID=1 #NbGéométries=" & iNbGeom.ToString & " /NbComposantes=" & pGeometryColl.GeometryCount.ToString & " /" & gsNomAttribut & " " & gsExpression, CType(pGeometryColl, IGeometry), iNbGeom)

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pGeometryColl = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer des géométries complexes ou multiples qui s'intersect entre elles.
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
    Private Function TraiterCreerGeometrieIntersecte(ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomBagColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface contenant la géométrie unique.
        Dim pGeometryBag As IGeometry = Nothing             'Interface contenant une géométrie de type BAG.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie polyline ou polygon.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier la géométrie de travail.
        Dim pSet(1) As ISet                                 'Vecteur des OIds des éléments à traiter.
        Dim iOidSel(1) As Integer           'Vecteur des OIds des éléments à traiter.
        Dim iNoSeq As Integer = -1          'Numéro de séquence de la dernière géométrie traitée.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim iNbGeom As Integer = 0          'Compteur de géométries.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieIntersecte = New GeometryBag
            TraiterCreerGeometrieIntersecte.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomBagColl = CType(TraiterCreerGeometrieIntersecte, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
                'Lire les éléments à traiter 
                LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

                'Redimensionner les Set
                ReDim Preserve pSet(pGeomSelColl.GeometryCount)

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement des intersections entre les géométries (" & gpFeatureLayerSelection.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomSelColl, IGeometryBag))

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des liens entre les géométries (" & gpFeatureLayerSelection.Name & ") ..."
                'Afficher la barre de progression
                InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)
                'Traiter tous les éléments
                For i = 0 To pRelResult.RelationElementCount - 1
                    'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                    pRelResult.RelationElement(i, iSel, iRel)

                    'Vérifier si Aucun Set
                    If pSet(iSel) Is Nothing Then
                        'Créer un nouveau Set vide
                        pSet(iSel) = New ESRI.ArcGIS.esriSystem.Set
                    End If

                    'Vérifier si le pointeur de sélection est plus petit que celui en relation
                    If iSel <> iRel Then
                        'Ajouter l'élément en relation
                        pSet(iSel).Add(iRel)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Création des géométries qui s'intersectent (" & gpFeatureLayerSelection.Name & ") ..."
                'Afficher la barre de progression
                InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

                'Traiter toutes les éléments
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Vérifier si la géométrie n'a pas été fusionnée
                    If pSet(i) IsNot Nothing Then
                        'Définir la géométrie
                        pGeometry = pGeomSelColl.Geometry(i)

                        'Redéfinir la géométrie
                        pGeometryColl = CType(pGeometry, IGeometryCollection)

                        'Compter les géométries
                        iNbGeom = pGeometryColl.GeometryCount

                        'Vérifier si aucun lien
                        If pSet(i).Count > 0 Then
                            'Vérifier si on doit fusionner
                            If gsExpression = "VRAI" Then
                                'Créer une nouvelle géométrie vide
                                pGeometryBag = New GeometryBag
                                'Définir la référence spatiale
                                pGeometryBag.SpatialReference = TraiterCreerGeometrieIntersecte.SpatialReference
                                pGeometryBag.SnapToSpatialReference()

                                'Ajouter les géométries en relation
                                AjouterGeometrieRelation(i, pGeometryBag, pGeomSelColl, pSet, iOidSel, True)

                                'Redéfinir la géométrie
                                pGeometryColl = CType(pGeometryBag, IGeometryCollection)

                                'Compter les géométries
                                iNbGeom = pGeometryColl.GeometryCount

                                'Si la géométrie est de type Polyline
                                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                                    'Créer une nouvelle géométrie vide
                                    pGeometry = New Polyline

                                    'Si la géométrie est de type Polygon
                                ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                                    'Créer une nouvelle géométrie vide
                                    pGeometry = New Polygon
                                End If

                                'Définir la référence spatiale
                                pGeometry.SpatialReference = TraiterCreerGeometrieIntersecte.SpatialReference
                                pGeometry.SnapToSpatialReference()

                                'Interface pour fusionner les géométries
                                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                                'Fusionner les géométries
                                pTopoOp.ConstructUnion(CType(pGeometryColl, IEnumGeometry))

                                'Redéfinir la géométrie
                                pGeometry = CType(pTopoOp, IGeometry)

                                'Redéfinir la géométrie
                                pGeometryColl = CType(pGeometry, IGeometryCollection)

                                'Si on ne doit pas fusionner
                            Else
                                'Ajouter les géométries en relation
                                AjouterGeometrieRelation(i, pGeometry, pGeomSelColl, pSet, iOidSel)

                                'Redéfinir la géométrie
                                pGeometryColl = CType(pGeometry, IGeometryCollection)

                                'Compter les géométries
                                iNbGeom = pGeometryColl.GeometryCount
                            End If
                        End If

                        'Ajouter la géométrie créée
                        pGeomBagColl.AddGeometry(pGeometry)

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #NbGéométries=" & iNbGeom.ToString & " /NbComposantes=" & pGeometryColl.GeometryCount.ToString & " /" & gsNomAttribut & " " & gsExpression, pGeometry, iNbGeom)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Cacher la barre de progression
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeomBagColl = Nothing
            pGeomSelColl = Nothing
            pGeometryColl = Nothing
            pGeometryBag = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pSet = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'ajouter les géométries en relation afin de former une seule géométrie.
    '''</summary>
    '''
    '''<param name="i"> Pointeur de géométrie traitée.</param>
    '''<param name="pGeometry"> Interface contenant la géométrie initiale.</param>
    '''<param name="pGeomSelColl"> Interface contenant toutes les géométries à traiter.</param>
    '''<param name="pSet"> Interface contenant les pointeurs des géométries en relation pour toutes les géométries.</param>
    '''<param name="iOidSel"> Contient les OBJECTID pour toutes les géométries à traiter.</param>
    '''<param name="bAjouter"> Indique si on doit ajouter la géométrie traitée dans la géométrie initiale.</param>
    ''' 
    Private Sub AjouterGeometrieRelation(ByVal i As Integer, ByRef pGeometry As IGeometry, ByRef pGeomSelColl As IGeometryCollection, _
                                         ByRef pSet() As ISet, ByRef iOidSel() As Integer, Optional bAjouter As Boolean = False)
        'Déclarer les variables de travail
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface pour ajouter la géométrie.
        Dim pObject As Object = Nothing                     'Interface contenant un objet dans un Set.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Vérifier si les relations de la géométrie n'ont pas été traitées
            If iOidSel(i) > -1 Then
                'Indiquer que la géométrie a tét traité
                iOidSel(i) = iOidSel(i) * -1

                'Extraire le premier
                pSet(i).Reset()
                pObject = pSet(i).Next()

                'Traiter tous les OIDs du Set
                Do Until pObject Is Nothing
                    'Définir la valeur du pointeur
                    iRel = CInt(pObject)

                    'Ajouter les géométries en relation
                    AjouterGeometrieRelation(iRel, pGeometry, pGeomSelColl, pSet, iOidSel, True)

                    'Extraire le premier 
                    pObject = pSet(i).Next()
                Loop

                'Vérifier s'il faut ajouter la géométrie
                If bAjouter Then
                    'Interface pour ajouter la géométrie
                    pGeometryColl = CType(pGeometry, IGeometryCollection)

                    'Vérifier si la géométrie est un Bag
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                        'Ajouter la géométrie dans une autre
                        pGeometryColl.AddGeometry(pGeomSelColl.Geometry(i))

                        'Si la géométrie n'est pas un Bag
                    Else
                        'Ajouter la géométrie dans une autre
                        pGeometryColl.AddGeometryCollection(CType(pGeomSelColl.Geometry(i), IGeometryCollection))
                    End If

                    'Détruire le Set de la géométrie à ajouter
                    pSet(i) = Nothing
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pObject = Nothing
            pGeometryColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de créer une géométrie complexe unique. Les géométries sont fusionnées.
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
    Private Function TraiterCreerGeometrieAttributVrai(ByRef pTrackCancel As ITrackCancel,
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing                  'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing                    'Interface pour sélectionner les éléments.
        Dim pGeometry As IGeometry = Nothing                            'Interface contenant la géométrie de travail.
        Dim pGeometryColl As IGeometryCollection = Nothing              'Interface contenant la géométrie de même valeur d'attribut.
        Dim pGeomSelColl As IGeometryCollection = Nothing               'Interface ESRI contenant les géométries sélectionnées.
        Dim pTopoOp As ITopologicalOperator2 = Nothing                  'Interface utilisé pour simplifier la géométrie de travail.
        Dim pSelColl As Dictionary(Of String, IGeometryBag) = Nothing   'Interface contenant les éléments à traiter.
        Dim pGeometryBag As IGeometryCollection = Nothing               'Interface conteant un ensemble d'éléments.
        Dim sListeAttribut As String = ""           'Contient la liste des noms d'attributs utilisée pour fusionner les géométries.
        Dim sLien As String = ""                    'Contient la valeur du lien.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieAttributVrai = New GeometryBag
            TraiterCreerGeometrieAttributVrai.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieAttributVrai, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Ajouter une virgule à la fin dans la liste des attributs pour corriger le problème de nom d'attributs partiels
                If gsListeAttribut <> "" Then sListeAttribut = gsListeAttribut & ","

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments à traiter (" & gpFeatureLayerSelection.Name & ") ..."
                'Lire les éléments à traiter 
                pSelColl = LireElementListeAttribut(gpFeatureLayerSelection, pTrackCancel, sListeAttribut)

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Création des géométries qui possèdent les mêmes valeurs d'attributs (" & gpFeatureLayerSelection.Name & ") ..."
                'Afficher la barre de progression
                InitBarreProgression(0, pSelColl.Count, pTrackCancel)

                'Extraire le lien de l'élément traité
                For Each sLien In pSelColl.Keys
                    'Définir l'ensemble
                    pGeometryBag = CType(pSelColl.Item(sLien), IGeometryCollection)

                    'Vérifier si le type de géométrie à créer est un Multipoint
                    If gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                        'Créer une nouvelle géométrie Multipoint vide
                        pGeometry = New Multipoint

                        'Si le type de géométrie à créer est une Polyline
                    ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer une nouvelle géométrie Polyline vide
                        pGeometry = New Polyline

                        'Si le type de géométrie à créer est un Polygon
                    ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Créer une nouvelle géométrie Polygon vide
                        pGeometry = New Polygon
                    End If

                    'Définir la référence spatiale
                    pGeometry.SpatialReference = TraiterCreerGeometrieAttributVrai.SpatialReference
                    pGeometry.SnapToSpatialReference()

                    'Interface pour fusionner les géométries
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)

                    'Fusionner les géométries
                    pTopoOp.ConstructUnion(CType(pGeometryBag, IEnumGeometry))

                    'Redéfinir la géométrie
                    pGeometry = CType(pTopoOp, IGeometry)

                    'Interface pour extraire le nombre de composantes
                    pGeometryColl = CType(pGeometry, IGeometryCollection)

                    'Ajouter la géométrie créée
                    pGeomSelColl.AddGeometry(pGeometry)

                    'Écrire une erreur
                    EcrireFeatureErreur(gsListeAttribut & "=" & sLien & " #NbGéométries=" & pGeometryBag.GeometryCount.ToString & " /NbComposantes=" & pGeometryColl.GeometryCount.ToString & " /" & gsNomAttribut & " " & gsExpression & " " & gsListeAttribut, pGeometry, pGeometryBag.GeometryCount)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pTopoOp = Nothing
            pSelColl = Nothing
            pGeometryBag = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer une géométrie multiple unique. Les géométries ne sont pas fusionnées.
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
    Private Function TraiterCreerGeometrieAttributFaux(ByRef pTrackCancel As ITrackCancel,
                                                       Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing                  'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing                    'Interface pour sélectionner les éléments.
        Dim pGeometry As IGeometry = Nothing                            'Interface contenant la géométrie de travail.
        Dim pGeomSelColl As IGeometryCollection = Nothing               'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometryColl As IGeometryCollection = Nothing              'Interface contenant la géométrie de même valeur d'attribut.
        Dim pSelColl As Dictionary(Of String, IGeometryBag) = Nothing   'Interface contenant les éléments à traiter.
        Dim pGeometryBag As IGeometryCollection = Nothing               'Interface conteant un ensemble d'éléments.
        Dim sListeAttribut As String = ""           'Contient la liste des noms d'attributs utilisée pour fusionner les géométries.
        Dim sLien As String = ""                    'Contient la valeur du lien.

        Try
            'Définir la géométrie par défaut
            TraiterCreerGeometrieAttributFaux = New GeometryBag
            TraiterCreerGeometrieAttributFaux.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCreerGeometrieAttributFaux, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Conserver la sélection de départ
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
                pSelectionSet = pFeatureSel.SelectionSet

                'Ajouter une virgule à la fin dans la liste des attributs pour corriger le problème de nom d'attributs partiels
                If gsListeAttribut <> "" Then sListeAttribut = gsListeAttribut & ","

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments à traiter (" & gpFeatureLayerSelection.Name & ") ..."
                'Lire les éléments à traiter 
                pSelColl = LireElementListeAttribut(gpFeatureLayerSelection, pTrackCancel, sListeAttribut)

                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Création des géométries qui possèdent les mêmes valeurs d'attributs (" & gpFeatureLayerSelection.Name & ") ..."
                'Afficher la barre de progression
                InitBarreProgression(0, pSelColl.Count, pTrackCancel)

                'Extraire le lien de l'élément traité
                For Each sLien In pSelColl.Keys
                    'Vérifier si le type de géométrie à créer est un Multipoint
                    If gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                        'Créer une nouvelle géométrie Multipoint vide
                        pGeometry = New Multipoint

                        'Si le type de géométrie à créer est une Polyline
                    ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer une nouvelle géométrie Polyline vide
                        pGeometry = New Polyline

                        'Si le type de géométrie à créer est un Polygon
                    ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Créer une nouvelle géométrie Polygon vide
                        pGeometry = New Polygon
                    End If

                    'Définir la référence spatiale
                    pGeometry.SpatialReference = TraiterCreerGeometrieAttributFaux.SpatialReference
                    pGeometry.SnapToSpatialReference()

                    'Interface pour créer la géométrie avec les mêmes valeurs d'attributs
                    pGeometryColl = CType(pGeometry, IGeometryCollection)

                    'Définir l'ensemble
                    pGeometryBag = CType(pSelColl.Item(sLien), IGeometryCollection)

                    'Traiter toutes les géométries
                    For i = 0 To pGeometryBag.GeometryCount - 1
                        'Définir la géométrie
                        pGeometry = pGeometryBag.Geometry(i)

                        'Ajouter la géométrie
                        pGeometryColl.AddGeometryCollection(CType(pGeometry, IGeometryCollection))
                    Next

                    'Définir la géométrie Multiple
                    pGeometry = CType(pGeometryColl, IGeometry)

                    'Ajouter la géométrie créée
                    pGeomSelColl.AddGeometry(pGeometry)

                    'Écrire une erreur
                    EcrireFeatureErreur(gsListeAttribut & "=" & sLien & " #NbGéométries=" & pGeometryBag.GeometryCount.ToString & " /NbComposantes=" & pGeometryColl.GeometryCount.ToString & " /" & gsNomAttribut & " " & gsExpression & " " & gsListeAttribut, pGeometry, pGeometryBag.GeometryCount)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Vérifier si on doit conserver la sélection
                If bEnleverSelection = False Then pFeatureSel.SelectionSet = pSelectionSet
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeometry = Nothing
            pGeomSelColl = Nothing
            pGeometryColl = Nothing
            pSelColl = Nothing
            pGeometryBag = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de lire les éléments d'un FeatureLayer et de les conserver par la liste des valeurs d'attributs spécifiés.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    '''</summary>
    '''
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="sListeAttribut"> Contient les noms d'attributs utilisés pour fusionner les géométries.</param>
    ''' 
    ''' <return> Dictionary(Of String, IGeometryBag) contenant les éléments par lien.</return>
    ''' 
    Private Function LireElementListeAttribut(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel, _
                                              ByVal sListeAttribut As String) As Dictionary(Of String, IGeometryBag)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pGeometryBag As IGeometryCollection = Nothing   'Interface contenant les géométries avec les mêmes valeurs d'attributs.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter un point.
        Dim sValeurAttribut As String = ""          'Contient les valeurs des attributs spécifiés.

        'Définir la valeur par défaut
        LireElementListeAttribut = New Dictionary(Of String, IGeometryBag)

        Try
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

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Initialiser les valeurs d'attributs
                sValeurAttribut = ""

                'Lire toutes les valeurs d'attributs
                For i = 0 To pFeature.Fields.FieldCount - 1
                    'Vérifier si l'attribut est présent dans la liste
                    If sListeAttribut.Contains(pFeature.Fields.Field(i).Name & ",") Then
                        'Si la liste est vide
                        If sValeurAttribut = "" Then
                            'Ajoute la valeur dans la liste
                            sValeurAttribut = pFeature.Value(i).ToString
                            'Si la liste n'est pas vide
                        Else
                            'Ajoute la valeur dans la liste
                            sValeurAttribut = sValeurAttribut & "," & pFeature.Value(i).ToString
                        End If
                    End If
                Next

                'Vérifier si le lien existe déjà
                If LireElementListeAttribut.ContainsKey(sValeurAttribut) Then
                    'Définir le Bag existant contenu dans le dictionnaire
                    pGeometryBag = CType(LireElementListeAttribut.Item(sValeurAttribut), IGeometryCollection)

                    'Si le lien n'existe pas
                Else
                    'Créer un nouveau Bag vide
                    pGeometry = New GeometryBag
                    'Définir la référence spatiale
                    pGeometry.SpatialReference = pFeature.Shape.SpatialReference
                    'Définir le lien pour ajouter les géométries dans la Bag
                    pGeometryBag = CType(pGeometry, IGeometryCollection)
                    'Ajouter l'ensemble dans le dictionnaire
                    LireElementListeAttribut.Add(sValeurAttribut, CType(pGeometryBag, IGeometryBag))
                End If

                'Vérifier si la géométrie est un Point
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                    'Créer un nouveau multipoint
                    pGeometry = New Multipoint
                    'Définir la référence spatiale
                    pGeometry.SpatialReference = pFeature.Shape.SpatialReference
                    'Interface pour ajouter le Point
                    pPointColl = CType(pGeometry, IPointCollection)
                    'Ajouter le point
                    pPointColl.AddPoint(CType(pFeature.ShapeCopy, IPoint))

                    'Si la géométrie n'est pas un Point
                Else
                    'Définir la géométrie
                    pGeometry = pFeature.ShapeCopy
                End If

                'Ajouter l'élément dans l'ensemble
                pGeometryBag.AddGeometry(pGeometry)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

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
            pGeometryBag = Nothing
            pGeometry = Nothing
            pPointColl = Nothing
        End Try
    End Function
#End Region
End Class
