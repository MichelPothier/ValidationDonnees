Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.ArcMapUI

'**
'Nom de la composante : clsComparerElement.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont les attributs et/ou la géométrie respecte ou non
''' la comparaison avec les éléments des FeatureLayers en lien d’attribut ou de géométrie.
''' 
''' La comparaison se fait sur les attributs ou la géométrie ou les deux.
''' 
''' La classe permet de traiter les deux attributs de comparaison ATTRIBUT et GEOMETRIE.
''' 
''' ATTRIBUT : Le lien entre les éléments à traiter et les éléments en relation est fait via une valeur d'attribut.
''' GEOMETRIE : Le lien entre les éléments à traiter et les éléments en relation est fait via le centroide de la géométrie.
''' 
''' Note : 
''' Pour le lien par attribut, il faut spécifier le nom de l'attribut.
''' Pour le lien par géométrie, il faut spécifier la tolérance de recherche.
''' Par défaut, tous les attributs sont comparés. Sinon, il faut spécifier les attributs à exclure incluant celui de la géométrie.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 20 mai 2015
'''</remarks>
''' 
Public Class clsComparerElement
    Inherits clsValeurAttribut

    'Déclarer les variables globales
    '''<summary>Liste des attributs à exclure du traitement de comparaison.</summary>
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
            FeatureLayersRelation = New Collection
            NomAttribut = "ATTRIBUT"
            Expression = "BDG_ID"
            ExclureAttribut = "OBJECTID"

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
            Nom = "ComparerElement"
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
                    'Vérifier si l'attribut est non éditable
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

                'Définir le paramètre pour comparer tous les attributs sauf le OID
                ListeParametres.Add("ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName)
                'Définir le paramètre pour comparer tous les attributs sauf le OID
                ListeParametres.Add("GEOMETRIE 0.1 " & gpFeatureLayerSelection.FeatureClass.OIDFieldName)
                'Définir le paramètre pour comparer tous les attributs sauf le OID
                ListeParametres.Add("GEOMETRIE/ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName)

                'Définir le paramètre pour comparer seulement les attributs sauf le OID
                ListeParametres.Add("ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName _
                                    & "," & gpFeatureLayerSelection.FeatureClass.ShapeFieldName)
                'Définir le paramètre pour comparer seulement les attributs sauf le OID
                ListeParametres.Add("GEOMETRIE 0.1 " & gpFeatureLayerSelection.FeatureClass.OIDFieldName _
                                    & "," & gpFeatureLayerSelection.FeatureClass.ShapeFieldName)
                'Définir le paramètre pour comparer seulement les attributs sauf le OID
                ListeParametres.Add("GEOMETRIE/ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName _
                                    & "," & gpFeatureLayerSelection.FeatureClass.ShapeFieldName)

                'Définir le paramètre pour comparer tous les attributs sauf les non-éditables
                ListeParametres.Add("ATTRIBUT BDG_ID " & sAttributsNonEditable)
                'Définir le paramètre pour comparer tous les attributs sauf les non-éditables
                ListeParametres.Add("GEOMETRIE 0.1 " & sAttributsNonEditable)
                'Définir le paramètre pour comparer tous les attributs sauf les non-éditables
                ListeParametres.Add("GEOMETRIE/ATTRIBUT BDG_ID " & sAttributsNonEditable)

                'Définir le paramètre pour comparer seulement la géométrie
                ListeParametres.Add("ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName & "," & sAttributs)
                'Définir le paramètre pour comparer seulement la géométrie
                ListeParametres.Add("GEOMETRIE 0.1 " & gpFeatureLayerSelection.FeatureClass.OIDFieldName & "," & sAttributs)
                'Définir le paramètre pour comparer seulement la géométrie
                ListeParametres.Add("GEOMETRIE/ATTRIBUT BDG_ID " & gpFeatureLayerSelection.FeatureClass.OIDFieldName & "," & sAttributs)
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
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function AttributValide() As Boolean
        'Déclarer les variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant un FeatureLayer en relation

        'La contrainte est invalide par défaut.
        AttributValide = False
        gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut

        Try
            'Vérifier si l'attribut est valide
            If gsNomAttribut = "ATTRIBUT" Or gsNomAttribut = "GEOMETRIE" Or gsNomAttribut = "GEOMETRIE/ATTRIBUT" Then
                'Vérifier si la géométrie n'est pas exclut de la comparaison
                If Not gsExclureAttribut.Contains(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) Then
                    'Traiter tous les FeatureLayer en relation
                    For Each pFeatureLayer In gpFeatureLayersRelation
                        'Vérifier si les types de géométrie sont différents
                        If gpFeatureLayerSelection.FeatureClass.ShapeType <> pFeatureLayer.FeatureClass.ShapeType Then
                            'Si la dimension est différente
                            If ((gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                               And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) _
                            Or (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                               And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint)) = False Then
                                'La comparaison est invalide.
                                AttributValide = False
                                gsMessage = "ERREUR : La comparaison est invalide pour ces types de géométries"
                                Exit Function
                            End If
                        End If
                    Next
                End If

                'La contrainte est valide
                AttributValide = True
                gsMessage = "La contrainte est valide"
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
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
        'Déclarer les variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant un FeatureLayer en relation

        Try
            'Retourner l'expression invalide par défaut
            ExpressionValide = False
            gsMessage = "ERREUR : L'expression est invalide"

            'Vérifier si le nom de l'attribut est GEOMETRIE
            If gsNomAttribut = "GEOMETRIE" Then
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

                'Si le nom de l'attribut est ATTRIBUT
            Else
                'Retourner l'expression valide si l'attribut est présent dans la FeatureClass
                ExpressionValide = gpFeatureLayerSelection.FeatureClass.FindField(gsExpression) <> -1
                'Vérifier si l'attribut est absent de la Featureclass
                If ExpressionValide Then
                    gsMessage = "La contrainte est valide"
                    'Traiter tous les FeatureLayer en relation
                    For Each pFeatureLayer In gpFeatureLayersRelation
                        'Vérifier si l'attribut de lien est absent de la Featureclass en relation
                        If pFeatureLayer.FeatureClass.FindField(gsExpression) = -1 Then
                            'Si l'attribut est absent de la Featureclass
                            ExpressionValide = False
                            gsMessage = "ERREUR : L'attribut de lien est absent de la FeatureClass en relation : " & pFeatureLayer.Name
                            Exit Function
                        End If
                    Next

                    'Si l'attribut est absent de la Featureclass
                Else
                    'Retourner le message d'erreur
                    gsMessage = "ERREUR : L'attribut de lien est absent de la FeatureClass de sélection : " & gpFeatureLayerSelection.Name
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont les attributs respecte ou non la condition de comparaison.
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

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Si le nom de l'attribut est ATTRIBUT
            If gsNomAttribut = "ATTRIBUT" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterLienAttribut(pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est GEOMETRIE
            ElseIf gsNomAttribut = "GEOMETRIE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterLienGeometrie(pTrackCancel, bEnleverSelection)
                'Traiter le FeatureLayer
                'Selectionner = TraiterLienCentreGeometrie(pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est GEOMETRIE
            ElseIf gsNomAttribut = "GEOMETRIE/ATTRIBUT" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterLienGeometrieAttribut(pTrackCancel, bEnleverSelection)
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non la comparaison avec ses éléments en relation de lien de géométrie et attribut.
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
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterLienGeometrieAttribut(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant l'élément à comparer.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pGeometryDiff As IGeometry = Nothing            'Interface contenant la différence de géométrie.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim iListOidSel As New List(Of Integer)             'Liste des objectid du Layer à traiter.
        Dim iListOidRel As New List(Of Integer)             'Liste des objectid du Layer à comparer.
        Dim iListPosAttSel As New List(Of Integer)          'Liste des positions d'attributs du Layer à traiter.
        Dim iListPosAttRel As New List(Of Integer)          'Liste des positions d'attributs du Layer à comparer.
        Dim iOidSel(0) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer                           'Vecteur des OIds des éléments en relation.
        Dim sDifference As String = ""                      'Contient les différences entre les éléments.
        Dim sExclureAttribut As String = ""                 'Contient les noms d'attributs à exclure.
        Dim bGeometrie As Boolean = False                   'Indique si on doit comparer la géométrie.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.
        Dim iPosAttSel As Integer = -1                      'Contient la position de l'attribut de lien dans le Layer à traiter.
        Dim iPosAttRel As Integer = -1                      'Contient la position de l'attribut de lien dans le Layer à comparer.

        Try
            'Définir la géométrie par défaut
            TraiterLienGeometrieAttribut = New GeometryBag
            TraiterLienGeometrieAttribut.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterLienGeometrieAttribut, IGeometryCollection)

            'Interface pour conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireCentreGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le FeatureLayer est différent de celui traité
                If Not gpFeatureLayerSelection.Equals(pFeatureLayerRel) Then
                    'Définir la position de  l'attribut de lien dans le Layer à traiter
                    iPosAttSel = gpFeatureLayerSelection.FeatureClass.FindField(gsExpression)
                    'Définir la position de  l'attribut de lien dans le Layer à comparer
                    iPosAttRel = pFeatureLayerRel.FeatureClass.FindField(gsExpression)

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireCentreGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)

                    'Définir la liste des attributs à traiter
                    Call DefinirListeAttribut(gpFeatureLayerSelection.FeatureClass.Fields, pFeatureLayerRel.FeatureClass.Fields, gsExclureAttribut, True,
                                              iListPosAttSel, iListPosAttRel, bGeometrie)

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des liens entre les éléments (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                    'Interface pour traiter la relation spatiale
                    pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                    'Exécuter la recherche et retourner le résultat de la relation spatiale
                    pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Comparaison des éléments à traiter (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                    'Afficher la barre de progression
                    InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

                    'Traiter tous les éléments dont les géométries sont égaux
                    For i = 0 To pRelResult.RelationElementCount - 1
                        'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                        pRelResult.RelationElement(i, iSel, iRel)

                        'Définir l'élément à traiter
                        pFeature = gpFeatureLayerSelection.FeatureClass.GetFeature(iOidSel(iSel))
                        'Définir la géométrie à traiter
                        pGeometry = pGeomSelColl.Geometry(iSel)

                        'Définir l'élément en relation
                        pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(iRel))
                        'Définir la géométrie en relation
                        pGeometryRel = pGeomRelColl.Geometry(iRel)

                        'Vérifier si l'attribut de lien correspond
                        If pFeature.Value(iPosAttSel).ToString = pFeatureRel.Value(iPosAttRel).ToString Then
                            'Comparer les attributs de l'élément à traiter avec celui à comparer
                            sDifference = ComparerAttributElement(pFeature, pFeatureRel, iListPosAttSel, iListPosAttRel)

                            'Vérifier si l'élément a déjà été traité
                            If iListOidSel.Contains(pFeature.OID) Then
                                'Ajouter un problème d'élément dupliqué
                                sDifference = sDifference & " #Élément traité dupliqué"
                            Else
                                'Ajouter le OID traité dans la liste
                                iListOidSel.Add(pFeature.OID)
                            End If

                            'Vérifier si l'élément en relation a déjà été traité
                            If iListOidRel.Contains(pFeatureRel.OID) Then
                                'Ajouter un problème d'élément dupliqué
                                sDifference = sDifference & " #Élément en relation dupliqué"
                            Else
                                'Ajouter le OID traité dans la liste
                                iListOidRel.Add(pFeatureRel.OID)
                            End If

                            'Vérifier si on doit comparer la géométrie
                            If bGeometrie Then
                                'Extraire la différence
                                pGeometryDiff = DifferenceGeometrie(pGeometry, pGeometryRel)

                                'Vérifier s'il n'y a pas une différence
                                If pGeometryDiff.IsEmpty Then
                                    'Définir la géométrie
                                    pGeometryDiff = pFeature.ShapeCopy
                                    's'il y a une différence
                                Else
                                    'Ajouter la différence
                                    sDifference = sDifference & "#" & gpFeatureLayerSelection.FeatureClass.ShapeFieldName & "/" & pFeatureLayerRel.FeatureClass.ShapeFieldName
                                End If
                            End If

                            'Vérifier si on doit sélectionner l'élément
                            If (sDifference = "" And Not bEnleverSelection) Or (sDifference <> "" And bEnleverSelection) Then
                                'Ajouter l'élément dans la sélection
                                pNewSelectionSet.Add(pFeature.OID)
                                'Définir la géométrie de l'élément
                                pGeomResColl.AddGeometry(pGeometry)

                                'S'il n'y a aucune différence
                                If sDifference = "" Then
                                    sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & " #Aucune différence"
                                    'S'il y a une différence
                                Else
                                    sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & sDifference
                                End If

                                'Écrire une erreur
                                EcrireFeatureErreur(sDifference, pGeometryDiff, pFeature.OID)
                            End If
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                    Next

                    'Cacher la barre de progression
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identifier les éléments non traités (" & gpFeatureLayerSelection.Name & ") ..."
                    'Vérifier si tous les éléments à traiter ont été traités 
                    For i = 0 To pGeomSelColl.GeometryCount - 1
                        'Vérifier si l'élément n'a pas été traité
                        If Not iListOidSel.Contains(iOidSel(i)) Then
                            'Ajouter la différence
                            sDifference = "OID=" & iOidSel(i).ToString & " #Aucun élément à comparer"
                            'Ajouter l'élément dans la sélection
                            pNewSelectionSet.Add(iOidSel(i))
                            'Définir l'élément à traiter
                            pFeature = gpFeatureLayerSelection.FeatureClass.GetFeature(iOidSel(i))
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pFeature.ShapeCopy)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pFeature.ShapeCopy, iOidSel(i))
                        End If
                    Next

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identifier les éléments en relation non traités (" & pFeatureLayerRel.Name & ") ..."
                    'Vérifier si tous les éléments en relation ont été traités 
                    For i = 0 To pGeomRelColl.GeometryCount - 1
                        'Vérifier si l'élément en relation n'a été comparé
                        If Not iListOidRel.Contains(iOidRel(i)) Then
                            'Ajouter la différence
                            sDifference = "OID=" & iOidRel(i).ToString & " #Élément en relation non comparé"
                            'Définir l'élément en relation
                            pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(i))
                            'Définir la géométrie de l'élément en relation
                            pGeomResColl.AddGeometry(pFeatureRel.ShapeCopy)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pFeatureRel.ShapeCopy, iOidRel(i))
                        End If
                    Next
                End If
            Next

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Succès du traitement (Nombre de géométries trouvées : " & pGeomResColl.GeometryCount.ToString & ") ..."

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeature = Nothing
            pFeatureLayerRel = Nothing
            pFeatureRel = Nothing
            pGeometry = Nothing
            pGeometryDiff = Nothing
            pGeometryRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            iListOidSel = Nothing
            iListOidRel = Nothing
            iListPosAttSel = Nothing
            iListPosAttRel = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non la comparaison avec ses éléments en relation de lien de géométrie.
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
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterLienGeometrie(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant l'élément à comparer.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pGeometryDiff As IGeometry = Nothing            'Interface contenant la différence de géométrie.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim iListOidSel As New List(Of Integer)             'Liste des objectid du Layer à traiter.
        Dim iListOidRel As New List(Of Integer)             'Liste des objectid du Layer à comparer.
        Dim iListPosAttSel As New List(Of Integer)          'Liste des positions d'attributs du Layer à traiter.
        Dim iListPosAttRel As New List(Of Integer)          'Liste des positions d'attributs du Layer à comparer.
        Dim iOidSel(0) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer                           'Vecteur des OIds des éléments en relation.
        Dim sDifference As String = ""                      'Contient les différences entre les éléments.
        Dim sExclureAttribut As String = ""                 'Contient les noms d'attributs à exclure.
        Dim bGeometrie As Boolean = False                   'Indique si on doit comparer la géométrie.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.

        Try
            'Définir la géométrie par défaut
            TraiterLienGeometrie = New GeometryBag
            TraiterLienGeometrie.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterLienGeometrie, IGeometryCollection)

            'Interface pour conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireCentreGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gdPrecision)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le FeatureLayer est différent de celui traité
                If Not gpFeatureLayerSelection.Equals(pFeatureLayerRel) Then
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    'LireCentreGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)

                    'Définir la liste des attributs à traiter
                    Call DefinirListeAttribut(gpFeatureLayerSelection.FeatureClass.Fields, pFeatureLayerRel.FeatureClass.Fields, gsExclureAttribut, True,
                                              iListPosAttSel, iListPosAttRel, bGeometrie)

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des liens entre les éléments (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                    'Interface pour traiter la relation spatiale
                    pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                    'Exécuter la recherche et retourner le résultat de la relation spatiale
                    pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Comparaison des éléments à traiter (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                    'Afficher la barre de progression
                    InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

                    'Traiter tous les éléments dont les géométries sont égaux
                    For i = 0 To pRelResult.RelationElementCount - 1
                        'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                        pRelResult.RelationElement(i, iSel, iRel)

                        'Définir l'élément à traiter
                        pFeature = gpFeatureLayerSelection.FeatureClass.GetFeature(iOidSel(iSel))
                        'Définir la géométrie à traiter
                        pGeometry = pGeomSelColl.Geometry(iSel)

                        'Définir l'élément en relation
                        pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(iRel))
                        'Définir la géométrie en relation
                        pGeometryRel = pGeomRelColl.Geometry(iRel)

                        'Comparer les attributs de l'élément à traiter avec celui à comparer
                        sDifference = ComparerAttributElement(pFeature, pFeatureRel, iListPosAttSel, iListPosAttRel)

                        'Vérifier si l'élément a déjà été traité
                        If iListOidSel.Contains(pFeature.OID) Then
                            'Ajouter un problème d'élément dupliqué
                            sDifference = sDifference & " #Élément traité dupliqué"
                        Else
                            'Ajouter le OID traité dans la liste
                            iListOidSel.Add(pFeature.OID)
                        End If

                        'Vérifier si l'élément en relation a déjà été traité
                        If iListOidRel.Contains(pFeatureRel.OID) Then
                            'Ajouter un problème d'élément dupliqué
                            sDifference = sDifference & " #Élément en relation dupliqué"
                        Else
                            'Ajouter le OID traité dans la liste
                            iListOidRel.Add(pFeatureRel.OID)
                        End If

                        'Vérifier si on doit comparer la géométrie
                        If bGeometrie Then
                            'Extraire la différence
                            pGeometryDiff = DifferenceGeometrie(pGeometry, pGeometryRel)

                            'Vérifier s'il n'y a pas une différence
                            If pGeometryDiff.IsEmpty Then
                                'Définir la géométrie
                                pGeometryDiff = pFeature.ShapeCopy
                                's'il y a une différence
                            Else
                                'Ajouter la différence
                                sDifference = sDifference & "#" & gpFeatureLayerSelection.FeatureClass.ShapeFieldName & "/" & pFeatureLayerRel.FeatureClass.ShapeFieldName
                            End If
                        End If

                        'Vérifier si on doit sélectionner l'élément
                        If (sDifference = "" And Not bEnleverSelection) Or (sDifference <> "" And bEnleverSelection) Then
                            'Ajouter l'élément dans la sélection
                            pNewSelectionSet.Add(pFeature.OID)
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pGeometry)

                            'S'il n'y a aucune différence
                            If sDifference = "" Then
                                sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & " #Aucune différence"
                                'S'il y a une différence
                            Else
                                sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & sDifference
                            End If

                            'Définir la géométrie de la différence  
                            If pGeometryDiff Is Nothing Then pGeometryDiff = pFeature.ShapeCopy

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pGeometryDiff, pFeature.OID)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                    Next

                    'Cacher la barre de progression
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identifier les éléments non traités (" & gpFeatureLayerSelection.Name & ") ..."
                    'Vérifier si tous les éléments à traiter ont été traités 
                    For i = 0 To pGeomSelColl.GeometryCount - 1
                        'Vérifier si l'élément n'a pas été traité
                        If Not iListOidSel.Contains(iOidSel(i)) Then
                            'Ajouter la différence
                            sDifference = "OID=" & iOidSel(i).ToString & " #Aucun élément à comparer"
                            'Ajouter l'élément dans la sélection
                            pNewSelectionSet.Add(iOidSel(i))
                            'Définir l'élément à traiter
                            pFeature = gpFeatureLayerSelection.FeatureClass.GetFeature(iOidSel(i))
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pFeature.ShapeCopy)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pFeature.ShapeCopy, iOidSel(i))
                        End If
                    Next

                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identifier les éléments en relation non traités (" & pFeatureLayerRel.Name & ") ..."
                    'Vérifier si tous les éléments en relation ont été traités 
                    For i = 0 To pGeomRelColl.GeometryCount - 1
                        'Vérifier si l'élément en relation n'a été comparé
                        If Not iListOidRel.Contains(iOidRel(i)) Then
                            'Ajouter la différence
                            sDifference = "OID=" & iOidRel(i).ToString & " #Élément en relation non comparé"
                            'Définir l'élément en relation
                            pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(i))
                            'Définir la géométrie de l'élément en relation
                            pGeomResColl.AddGeometry(pFeatureRel.ShapeCopy)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pFeatureRel.ShapeCopy, iOidRel(i))
                        End If
                    Next
                End If
            Next

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Succès du traitement (Nombre de géométries trouvées : " & pGeomResColl.GeometryCount.ToString & ") ..."

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeature = Nothing
            pFeatureLayerRel = Nothing
            pFeatureRel = Nothing
            pGeometry = Nothing
            pGeometryDiff = Nothing
            pGeometryRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            iListOidSel = Nothing
            iListOidRel = Nothing
            iListPosAttSel = Nothing
            iListPosAttRel = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non la comparaison avec ses éléments en relation de lien du centre de la géométrie.
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
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterLienCentreGeometrie(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pFeatureRel As IFeature = Nothing               'Interface contenant l'élément à comparer.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pGeometryDiff As IGeometry = Nothing            'Interface contenant la différence de géométrie.
        Dim pGeometryRel As IGeometry = Nothing             'Interface contenant une géométrie.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant la géométrie d'un l'élément en relation.
        Dim pCursor As ICursor = Nothing                    'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les éléments sélectionnés.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant la relation spatiale de base.
        Dim oFeatureColl As Collection = Nothing            'Objet contenant la collection des éléments en relation.
        Dim pPoint As IPoint = New Point                    'Interface contenant le point du centre de la ligne.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour créer un Buffer.
        Dim sDifference As String = ""                      'Contient les différences entre les éléments.
        Dim sExclureAttribut As String = ""                 'Contient les noms d'attributs à exclure.
        Dim bGeometrie As Boolean = False                   'Indique si on doit comparer la géométrie.

        Try
            'Définir la géométrie par défaut
            TraiterLienCentreGeometrie = New GeometryBag
            TraiterLienCentreGeometrie.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterLienCentreGeometrie, IGeometryCollection)

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

            'Indiquer si on doit comparer la géométrie
            bGeometrie = Not sExclureAttribut.Contains(gpFeatureLayerSelection.FeatureClass.ShapeFieldName & ",")

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le FeatureLayer est différent de celui traité
                If Not gpFeatureLayerSelection.Equals(pFeatureLayerRel) Then
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Comparaison des éléments à traiter (" & gpFeatureLayerSelection.Name & ") ..."

                    'Afficher la barre de progression
                    InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

                    'Créer la requête spatiale
                    pSpatialFilter = New SpatialFilterClass
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    'Définir la référence spatiale de sortie dans la requête spatiale
                    pSpatialFilter.OutputSpatialReference(pFeatureLayerRel.FeatureClass.ShapeFieldName) = TraiterLienCentreGeometrie.SpatialReference

                    'Interfaces pour extraire les éléments sélectionnés
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)

                    'Extraire le premier élément
                    pFeature = pFeatureCursor.NextFeature()

                    'Traiter tous les éléments sélectionnés du FeatureLayer
                    Do While Not pFeature Is Nothing
                        'Initialiser les différences
                        sDifference = ""
                        'Définir la géométrie à traiter
                        pGeometry = pFeature.ShapeCopy
                        'Projeter la géométrie à traiter
                        pGeometry.Project(TraiterLienCentreGeometrie.SpatialReference)

                        'Extraire le centre d'une géométrie
                        pPoint = CentreGeometrie(pGeometry)

                        'Vérifier si la précision est spécifié
                        If gdPrecision > 0 Then
                            'Interface pour créer un buffer
                            pTopoOp = CType(pPoint, ITopologicalOperator)
                            'Définir la géométrie utilisée pour la relation spatiale
                            pSpatialFilter.Geometry = pTopoOp.Buffer(gdPrecision)
                        Else
                            'Définir la géométrie utilisée pour la relation spatiale
                            pSpatialFilter.Geometry = pPoint
                        End If

                        'Extraire les éléments en relation
                        oFeatureColl = ExtraireElementsRelation(pSpatialFilter, pFeatureLayerRel)

                        'Vérifier la présence d'éléments en relation
                        If oFeatureColl.Count > 0 Then
                            'Traiter tous les éléments en relation
                            For Each pFeatureRel In oFeatureColl
                                'Définir la géométrie à traiter
                                pGeometryRel = pFeatureRel.ShapeCopy
                                'Projeter la géométrie à traiter
                                pGeometryRel.Project(TraiterLienCentreGeometrie.SpatialReference)

                                'Comparer les attributs de l'élément à traiter avec celui à comparer
                                sDifference = ComparerAttributElement(pFeature, pFeatureRel, sExclureAttribut)

                                'Vérifier si on doit comparer la géométrie
                                If bGeometrie Then
                                    'Extraire la différence
                                    pGeometryDiff = DifferenceGeometrie(pGeometry, pGeometryRel)
                                    'Vérifier s'il n'y a pas une différence
                                    If pGeometryDiff.IsEmpty Then
                                        'Définir la géométrie
                                        pGeometryDiff = pFeature.ShapeCopy
                                        's'il y a une différence
                                    Else
                                        'Ajouter la différence
                                        sDifference = sDifference & "#" & gpFeatureLayerSelection.FeatureClass.ShapeFieldName & "/" & pFeatureLayerRel.FeatureClass.ShapeFieldName
                                    End If
                                End If

                                'Vérifier si on doit sélectionner l'élément
                                If (sDifference = "" And Not bEnleverSelection) Or (sDifference <> "" And bEnleverSelection) Then
                                    'Ajouter l'élément dans la sélection
                                    pNewSelectionSet.Add(pFeature.OID)
                                    'Définir la géométrie de l'élément
                                    pGeomResColl.AddGeometry(pGeometry)

                                    'S'il n'y a aucune différence
                                    If sDifference = "" Then
                                        sDifference = "OID=" & pFeature.OID.ToString & " #Aucune différence"
                                    Else
                                        sDifference = "OID=" & pFeature.OID.ToString & sDifference
                                    End If
                                    'Écrire une erreur
                                    EcrireFeatureErreur(sDifference, pGeometryDiff)
                                End If
                            Next

                            'Si aucun élément en relation
                        Else
                            'Ajouter la différence
                            sDifference = "OID=" & pFeature.OID.ToString & " #Aucun élément à comparer"
                            'Ajouter l'élément dans la sélection
                            pNewSelectionSet.Add(pFeature.OID)
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pGeometry)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pGeometry)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                        'Extraire le prochain élément
                        pFeature = pFeatureCursor.NextFeature()
                    Loop
                End If
            Next

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
            pFeatureLayerRel = Nothing
            pFeatureRel = Nothing
            pGeometry = Nothing
            pGeometryDiff = Nothing
            pGeometryRel = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
            oFeatureColl = Nothing
            pPoint = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non la comparaison avec ses éléments en relation de lien d'attributs.
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
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterLienAttribut(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing              'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing                'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing             'Interface pour sélectionner les éléments.
        Dim pSelColl As Dictionary(Of String, IFeature) = Nothing   'Interface contenant les éléments à traiter.
        Dim pRelColl As Dictionary(Of String, IFeature) = Nothing   'Interface contenant les éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing             'Interface contenant un FeatureLayer en relation.
        Dim pGeomResColl As IGeometryCollection = Nothing           'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pFeature As IFeature = Nothing                          'Interface contenant l'élément à traiter.
        Dim pFeatureRel As IFeature = Nothing                       'Interface contenant l'élément à comparer.
        Dim pGeometry As IGeometry = Nothing                        'Interface contenant une géométrie.
        Dim pGeometryRel As IGeometry = Nothing                     'Interface contenant une géométrie en relation.
        Dim iListPosAttSel As New List(Of Integer)                  'Liste des positions d'attributs du Layer à traiter.
        Dim iListPosAttRel As New List(Of Integer)                  'Liste des positions d'attributs du Layer à comparer.
        Dim sDifference As String = ""                              'Contient les différences entre les éléments.        
        Dim sLien As String = ""                                    'Contient la valeur du lien.
        Dim bGeometrie As Boolean = False                           'Indique si on doit comparer la géométrie.

        Try
            'Définir la géométrie par défaut
            TraiterLienAttribut = New GeometryBag
            TraiterLienAttribut.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterLienAttribut, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments à traiter (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            pSelColl = LireElementLienAttribut(gpFeatureLayerSelection, pTrackCancel, gsExpression)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le FeatureLayer est différent de celui traité
                If Not gpFeatureLayerSelection.Equals(pFeatureLayerRel) Then
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments à comparer (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    pRelColl = LireElementLienAttribut(pFeatureLayerRel, pTrackCancel, gsExpression)

                    'Définir la liste des attributs à traiter
                    Call DefinirListeAttribut(gpFeatureLayerSelection.FeatureClass.Fields, pFeatureLayerRel.FeatureClass.Fields, gsExclureAttribut, True,
                                              iListPosAttSel, iListPosAttRel, bGeometrie)

                    'Afficher le message de traitement de la relation spatiale
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Comparaison en cours (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                    'Afficher la barre de progression
                    InitBarreProgression(0, pSelColl.Count, pTrackCancel)

                    'Extraire le lien de l'élément traité
                    For Each sLien In pSelColl.Keys
                        'Définir l'élément
                        pFeature = pSelColl.Item(sLien)
                        'Définir la géométrie
                        pGeometry = pFeature.ShapeCopy
                        'Projeter la géométrie
                        pGeometry.Project(TraiterLienAttribut.SpatialReference)

                        'Vérifier si le lien est présent dans les éléments à comparer
                        If pRelColl.ContainsKey(sLien) Then
                            'Comparer les attributs de l'élément à traiter avec celui à comparer
                            sDifference = ComparerAttributElement(pSelColl.Item(sLien), pRelColl.Item(sLien), iListPosAttSel, iListPosAttRel)

                            'Vérifier si on doit comparer la géométrie
                            If bGeometrie Then
                                'Définir l'élément
                                pFeatureRel = pRelColl.Item(sLien)
                                'Définir la géométrie en relation
                                pGeometryRel = pFeatureRel.ShapeCopy
                                'Projeter la géométrie
                                pGeometryRel.Project(TraiterLienAttribut.SpatialReference)
                                'Extraire la différence
                                pGeometry = DifferenceGeometrie(pGeometry, pGeometryRel)
                                'Vérifier s'il n'y a pas une différence
                                If pGeometry.IsEmpty Then
                                    'Définir la géométrie
                                    pGeometry = pFeature.ShapeCopy
                                    's'il y a une différence
                                Else
                                    'Ajouter la différence
                                    sDifference = sDifference & " #Géométrie différente (" & gpFeatureLayerSelection.FeatureClass.ShapeFieldName & "/" & pFeatureLayerRel.FeatureClass.ShapeFieldName & ")"
                                End If
                            End If

                            'Vérifier si on doit sélectionner l'élément
                            If (sDifference = "" And Not bEnleverSelection) Or (sDifference <> "" And bEnleverSelection) Then
                                'Ajouter l'élément dans la sélection
                                pNewSelectionSet.Add(pFeature.OID)
                                'Définir la géométrie de l'élément
                                pGeomResColl.AddGeometry(pGeometry)

                                'S'il n'y a aucune différence
                                If sDifference = "" Then
                                    sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & " #Aucune différence"
                                Else
                                    sDifference = "OID=" & pFeature.OID.ToString & "/" & pFeatureRel.OID.ToString & sDifference
                                End If
                                'Écrire une erreur
                                EcrireFeatureErreur(sDifference, pGeometry, pFeature.OID)
                            End If
                            'Retirer l'élément en relation
                            'pRelColl.Remove(sLien)

                            'S'il n'y a pas de lien
                        Else
                            'Ajouter la différence
                            sDifference = "OID=" & pFeature.OID.ToString & " #Aucun élément à comparer"
                            'Ajouter l'élément dans la sélection
                            pNewSelectionSet.Add(pFeature.OID)
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pGeometry)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pGeometry, pFeature.OID)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                    Next

                    'Extraire le lien de l'élément à comparer
                    For Each sLien In pRelColl.Keys
                        'Vérifier si le lien est présent dans les éléments à traiter
                        If Not pSelColl.ContainsKey(sLien) Then
                            'Définir l'élément
                            pFeatureRel = pRelColl.Item(sLien)
                            'Définir la géométrie
                            pGeometryRel = pFeatureRel.ShapeCopy
                            'Projeter la géométrie
                            pGeometryRel.Project(TraiterLienAttribut.SpatialReference)

                            'Ajouter la différence
                            sDifference = "OID=" & pFeatureRel.OID.ToString & " #Élément à comparer non traité"
                            'Définir la géométrie de l'élément
                            pGeomResColl.AddGeometry(pGeometryRel)

                            'Écrire une erreur
                            EcrireFeatureErreur(sDifference, pGeometryRel, pFeatureRel.OID)
                        End If
                    Next
                End If
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pSelColl = Nothing
            pRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pFeature = Nothing
            pFeatureRel = Nothing
            pGeometry = Nothing
            pGeometryRel = Nothing
            iListPosAttSel = Nothing
            iListPosAttRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de lire le centre des géométries et les OIDs des éléments d'un FeatureLayer.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    '''</summary>
    ''' 
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomColl"> Interface contenant les géométries des éléments lus.</param>
    '''<param name="iOid"> Vecteur des OIDs d'éléments lus.</param>
    '''<param name="dTol"> Tolérance utilisée pour créer un tampon autour du cente de la géométrie.</param>
    ''' 
    Private Sub LireCentreGeometrie(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel, ByRef pGeomColl As IGeometryCollection,
                                    Optional ByRef iOid() As Integer = Nothing, Optional ByVal dTol As Double = 0)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément lu.
        Dim pPoint As IPoint = Nothing                      'Interface contenant le centre de la géométrie.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour extraire la limite de la géométrie.

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

            'Interface pour extraire les éléments sélectionnés
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Augmenter le vecteur des Oid selon le nombre d'éléments
            ReDim Preserve iOid(pSelectionSet.Count)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments du FeatureLayer
            For i = 0 To pSelectionSet.Count - 1
                'Vérifier si l'élément est présent
                If pFeature IsNot Nothing Then
                    'Vérifier si la géométrie n'est pas vide
                    If Not pFeature.Shape.IsEmpty Then
                        'Définir la géométrie à traiter
                        pGeometry = pFeature.ShapeCopy

                        'Projeter la géométrie à traiter
                        pGeometry.Project(pFeatureLayer.AreaOfInterest.SpatialReference)
                        pGeometry.SnapToSpatialReference()

                        'Extraire le centre d'une géométrie
                        pPoint = CentreGeometrie(pGeometry)

                        'Vérifier si une tolérance est spécifiée
                        If dTol > 0 Then
                            'Interface pour créer un buffer
                            pTopoOp = CType(pPoint, ITopologicalOperator)
                            'Ajouter la géométrie dans le Bag
                            pGeomColl.AddGeometry(pTopoOp.Buffer(dTol))
                        Else
                            'Ajouter la géométrie dans le Bag
                            pGeomColl.AddGeometry(pPoint)
                        End If

                        'Ajouter le OID de l'élément avec sa séquence 
                        iOid(i) = pFeature.OID
                    End If

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
            pPoint = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de lire les éléments d'un FeatureLayer et de les conserver par lien d'attribut et ObjectId.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="sNomAttributLien"> Nom de l'attribut utilisé pour conserver le lien.</param>
    ''' 
    ''' <return> Dictionary(Of String, IFeature) contenant les éléments par lien.</return>
    '''</summary>
    '''
    Private Function LireElementLienAttribut(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel, _
                                             ByVal sNomAttributLien As String) As Dictionary(Of String, IFeature)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim iPosAtt As Integer = -1                         'Contient la position de l'attribut utilisé pour faire le lien.

        'Définir la valeur par défaut
        LireElementLienAttribut = New Dictionary(Of String, IFeature)

        Try
            'Définir la position de l'attribut en lien
            iPosAtt = pFeatureLayer.FeatureClass.Fields.FindField(sNomAttributLien)

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
                'Vérifier si le lien existe déjà
                If LireElementLienAttribut.ContainsKey(pFeature.Value(iPosAtt).ToString) Then
                    'Conserver le ObjectID de l'élément par le lien
                    LireElementLienAttribut.Add(pFeature.Value(iPosAtt).ToString & "_" & pFeature.OID.ToString, pFeature)
                Else
                    'Conserver le ObjectID de l'élément par le lien
                    LireElementLienAttribut.Add(pFeature.Value(iPosAtt).ToString, pFeature)
                End If

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
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de définir les positions des attributs à comparer et indiquer si la géométrie doit être comparée.
    '''</summary>
    ''' 
    ''' Les OID et les géométries sont toujours exclus.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFields"> Interface contenant les attribut des éléments à traiter.</param>
    '''<param name="pFieldsRel"> Interface contenant les attribut des éléments à comparer.</param>
    '''<param name="sListeAttribut"> Contient les noms des attributs à comparer.</param>
    '''<param name="bExclure"> Indique les noms des attributs à comparer doivent être exclus ou non.</param>
    '''<param name="iListSel"> Liste des positions d'attributs pour le Layer à traiter.</param>
    '''<param name="iListRel"> Liste des positions d'attributs pour le Layer à comparer.</param>
    '''<param name="bGeometrie"> Indique si la géométrie doit être comparée ou non.</param>
    '''
    Private Sub DefinirListeAttribut(ByVal pFields As IFields, ByVal pFieldsRel As IFields, ByVal sListeAttribut As String, ByVal bExclure As Boolean,
                                     ByRef iListSel As List(Of Integer), ByRef iListRel As List(Of Integer), ByRef bGeometrie As Boolean)
        'Déclarer les variables de travail
        Dim bContains As Boolean = False        'Indique si l'attribut est présent dans la liste d'attributs.

        'Par défaut, on ne traite pas la géométrie
        bGeometrie = False

        Try
            'Initialiser la liste d'attributs
            sListeAttribut = sListeAttribut & ","

            'Traiter tous les attributs
            For i = 0 To pFields.FieldCount - 1
                'Indique si l'attribut est présent dans la liste d'attributs
                bContains = sListeAttribut.Contains(pFields.Field(i).Name & ",")

                'Vérifier si l'attribut est à traiter
                If (bContains And Not bExclure) Or (Not bContains And bExclure) Then
                    'Vérifier si l'attribut n'est pas une géométrie
                    If Not (pFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Or pFields.Field(i).Type = esriFieldType.esriFieldTypeOID) Then
                        'Définir la position de l'attribut de l'élément à traiter
                        iListSel.Add(i)

                        'Définir la position de l'attribut de l'élément à comparer
                        iListRel.Add(pFieldsRel.FindField(pFields.Field(i).Name))

                        'Vérifier si l'attribut est une géométrie
                    ElseIf pFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Then
                        'Indiquer que la géométrie doit être traitée
                        bGeometrie = True
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de comparer les attributs spécifiés entre deux éléments.
    '''</summary>
    ''' 
    ''' Les OID et les géométries sont toujours exclus.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFeature"> Interface contenant l'élément à traiter.</param>
    '''<param name="pFeatureRel"> Interface contenant l'élément à comparer.</param>
    '''<param name="iListPosAttSel"> Liste des positions d'attributs pour le Layer à traiter.</param>
    '''<param name="iListPosAttRel"> Liste des positions d'attributs pour le Layer à comparer.</param>
    ''' 
    ''' <return> String contenant le résultat de la comparaison.</return>
    '''
    Private Function ComparerAttributElement(ByVal pFeature As IFeature, ByVal pFeatureRel As IFeature,
                                             ByVal iListPosAttSel As List(Of Integer), ByVal iListPosAttRel As List(Of Integer)) As String
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing        'Interface contenant les attribut des éléments à traiter.
        Dim pFieldsRel As IFields = Nothing     'Interface contenant les attribut des éléments à comparer.

        'Définir la valeur par défaut
        ComparerAttributElement = ""

        Try
            'Interface pour traiter les attributs
            pFields = pFeature.Fields
            pFieldsRel = pFeatureRel.Fields

            'Traiter tous les attributs
            For i = 0 To iListPosAttSel.Count - 1
                'Vérifier si l'attribut à comparer est présent
                If iListPosAttRel.Item(i) > 0 Then
                    'Vérifier si l'attribut n'est pas une géométrie ou un OID
                    If Not (pFields.Field(iListPosAttSel.Item(i)).Type = esriFieldType.esriFieldTypeGeometry Or pFields.Field(iListPosAttSel.Item(i)).Type = esriFieldType.esriFieldTypeOID) Then
                        'Vérifier si la valeur de l'attribut est différent
                        If pFeature.Value(iListPosAttSel.Item(i)).ToString() <> pFeatureRel.Value(iListPosAttRel.Item(i)).ToString() Then
                            'Ajouter la différence
                            ComparerAttributElement = ComparerAttributElement & " #" & pFields.Field(iListPosAttSel.Item(i)).Name & "=" & pFeature.Value(iListPosAttSel.Item(i)).ToString _
                                                                              & "/" & pFeatureRel.Value(iListPosAttRel.Item(i)).ToString
                        End If
                    Else
                        'Ajouter la différence
                        ComparerAttributElement = ComparerAttributElement & " #" & pFields.Field(i).Name & "Aucun attribut correspond"
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
    ''' Routine qui permet de comparer tous les attributs qui ne sont pas exclut entre deux éléments.
    ''' 
    ''' Les OID et les géométries sont toujours exclus.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFeature"> Interface contenant l'élément à traiter.</param>
    '''<param name="pFeatureRel"> Interface contenant l'élément à comparer.</param>
    '''<param name="sNomAttributLien"> Nom de l'attribut utilisé pour conserver le lien.</param>
    ''' 
    ''' <return> String contenant le résultat de la comparaison.</return>
    '''</summary>
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
                    'Vérifier si l'attribut n'est pas une géométrie
                    If Not (pFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Or pFields.Field(i).Type = esriFieldType.esriFieldTypeOID) Then
                        'Définir la position de l'attribut de l'élément en relation
                        iPosAtt = pFeatureRel.Fields.FindField(pFields.Field(i).Name)

                        'Vérifier si l'attribut est présent de la FeatureClass en relation
                        If iPosAtt >= 0 Then
                            'Vérifier si la valeur de l'attribut est différent
                            If pFeature.Value(i).ToString() <> pFeatureRel.Value(iPosAtt).ToString() Then
                                'Ajouter la différence
                                ComparerAttributElement = ComparerAttributElement & " #" & pFields.Field(i).Name & "=" & pFeature.Value(i).ToString _
                                                          & "/" & pFeatureRel.Value(iPosAtt).ToString
                            End If

                            'si l'attribut est absent de la FeatureClass en relation
                        Else
                            'Ajouter la différence
                            ComparerAttributElement = ComparerAttributElement & " #" & pFields.Field(i).Name & "/Aucun attribut correspond"
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
    ''' Fonction qui permet d'extraire les différences entre deux géométries.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à traiter.</param>
    '''<param name="pGeometryRel">Interface ESRI contenant la géométrie à comparer.</param>
    '''<param name="bSommet">Indique qu'on veut vérifier sommet par sommet.</param>
    ''' 
    '''<returns>"IGeometry" contenant les différences. S'il n'y a pas de différence, un multipoint vide est retourné.</returns>
    '''
    Protected Function DifferenceGeometrie(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry, _
                                           Optional ByVal bSommet As Boolean = False) As IGeometry
        'Déclarer les variables de travail
        Dim pPointCollA As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets de la géométrie.
        Dim pPointCollB As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets du clone de la géométrie.
        Dim pMultiPointColl As IPointCollection = Nothing   'Interface qui permet d'accéder aux sommets du multipoint contenant les différences.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface utilisé pour extraire la différence.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier les relations spatiales.

        'Définir la valeur de retour par défaut
        DifferenceGeometrie = New Multipoint

        Try
            'Interface pour vérifier la relation
            pRelOp = CType(pGeometry, IRelationalOperator)

            'Vérifier si la géométrie est différente
            'If Not pGeometry.Equals(pGeometryRel) Then
            If pRelOp.Disjoint(pGeometryRel) Then
                'Vérifier si la géométrie est de type point
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                    'Définir la différence
                    DifferenceGeometrie = pGeometry
                Else
                    'Interface pour trouver la différence
                    pTopoOp = CType(pGeometry, ITopologicalOperator)
                    'Trouver la différence
                    DifferenceGeometrie = pTopoOp.Difference(pGeometryRel)
                End If

                'Vérifier si la géométrie est vide
                If DifferenceGeometrie.IsEmpty And bSommet Then
                    'Vérifier une des géométries n'est pas un point
                    If pGeometry.GeometryType <> esriGeometryType.esriGeometryPoint _
                    And pGeometryRel.GeometryType <> esriGeometryType.esriGeometryPoint Then
                        'Interface pour traiter les sommets de la géométrie à traiter
                        pPointCollA = CType(pGeometry, IPointCollection)

                        'Interface pour traiter les sommets de la géométrie à comparer
                        pPointCollB = CType(pGeometryRel, IPointCollection)

                        'vérifier si le nombre de points est égale
                        If pPointCollA.PointCount = pPointCollB.PointCount Then
                            'Définir la valeur de retour par défaut
                            DifferenceGeometrie = New Multipoint
                            'Interface pour créer un multipoint vide contenant les différences
                            DifferenceGeometrie.SpatialReference = pGeometry.SpatialReference
                            'Interface pour ajouter les sommets dans le multipoints
                            pMultiPointColl = CType(DifferenceGeometrie, IPointCollection)

                            'Ajouter tous les ID sur tous les sommets de la géométrie
                            For i = 0 To pPointCollA.PointCount - 1
                                'Vérifier si le point est égale
                                If Not pPointCollA.Point(i).Equals(pPointCollB.Point(i)) Then
                                    'Ajouter le point différent
                                    pMultiPointColl.AddPoint(pPointCollA.Point(i))
                                End If
                            Next
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPointCollA = Nothing
            pPointCollB = Nothing
            pMultiPointColl = Nothing
            pTopoOp = Nothing
            pRelOp = Nothing
        End Try
    End Function
#End Region
End Class
