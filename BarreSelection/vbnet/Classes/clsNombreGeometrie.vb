Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsNombreGeometrie.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont le nombre de géométrie de composante, 
''' d’anneau extérieur, d’anneau intérieur ou de sommet respecte ou non l’expression régulière spécifiée.
''' 
''' La classe permet de traiter les attributs de nombre de géométries COMPOSANTE, EXTERIEUR, INTERIEUR ET SOMMET.
''' 
''' COMPOSANTE : Le nombre de composante (point, ligne, anneau extérieur ou anneau intérieur) contenue dans la géométrie. 
''' EXTERIEUR : Le nombre d'anneau extérieur contenu dans la géométrie.
''' INTERIEUR : Le nombre d'anneau intérieur par anneau extérieur contenu dans la géométrie.
''' SOMMET : Le nombre de sommet contenu dans la géométrie.
''' 
''' Note : Un Polygone est composé de 1 à N anneau(x) extérieur(s) et 0 à N anneau(x) intérieur(s).
'''        Les anneaux intérieurs sont liés obligatoirement à un anneau extérieur.     
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 21 avril 2015
'''</remarks>
''' 
Public Class clsNombreGeometrie
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
            NomAttribut = "COMPOSANTE"
            Expression = "^1$"

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
            Nom = "NombreGeometrie"
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
                'Définir le paramètre pour trouver les superficies des anneaux (extérieurs ou intérieurs) de géométries
                ListeParametres.Add("COMPOSANTE ^1$")
                'Définir le paramètre pour trouver les superficies des anneaux extérieurs de géométries
                ListeParametres.Add("EXTERIEUR ^1$")
                'Définir le paramètre pour trouver les superficies des anneaux intérieurs de géométries
                ListeParametres.Add("INTERIEUR ^\d$")
                'Définir le paramètre pour trouver les superficies totale des géométries
                ListeParametres.Add("SOMMET ^[1-2]$")
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
                'Vérifier si la FeatureClass est de type MultiPoint, Polyline ou Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Vérifier si le nom de l'attribut est EXTERIEUR OU INTERIEUR
                    If gsNomAttribut = "EXTERIEUR" Or gsNomAttribut = "INTERIEUR" Then
                        'Vérifier si la FeatureClass est de type Polygon
                        If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                            'La contrainte est valide
                            FeatureClassValide = True
                            gsMessage = "La contrainte est valide."
                        Else
                            gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polygon."
                        End If
                    Else
                        'La contrainte est valide
                        FeatureClassValide = True
                        gsMessage = "La contrainte est valide."
                    End If
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type MultiPoint, Polyline ou Polygon."
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
            If gsNomAttribut = "COMPOSANTE" Or gsNomAttribut = "EXTERIEUR" Or gsNomAttribut = "INTERIEUR" Or gsNomAttribut = "SOMMET" Then
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre de géométrie respecte ou non l'expression régulière spécifiée.
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
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Si le nom de l'attribut est COMPOSANTE
            If gsNomAttribut = "COMPOSANTE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterNombreComposante(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est EXTERIEUR
            ElseIf gsNomAttribut = "EXTERIEUR" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterNombreAnneauExterieur(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est INTERIEUR
            ElseIf gsNomAttribut = "INTERIEUR" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterNombreAnneauInterieur(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est SOMMET
            ElseIf gsNomAttribut = "SOMMET" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterNombreSommet(pFeatureCursor, pTrackCancel, bEnleverSelection)
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre de sommet respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterNombreSommet(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire le nombre de sommet.

        Try
            'Définir la géométrie par défaut
            TraiterNombreSommet = New GeometryBag
            TraiterNombreSommet.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreSommet, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pGeometry = pFeature.Shape
                    pGeometry.Project(TraiterNombreSommet.SpatialReference)

                    'Interface pour extraire le nombre de sommet
                    pPointColl = CType(pGeometry, IPointCollection)

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pPointColl.PointCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pGeometry.Envelope)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbSommets=" & pPointColl.PointCount.ToString & " /ExpReg=" & gsExpression, _
                                            pGeometry, pPointColl.PointCount)
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
            pPointColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre de composante (point, ligne, anneau extérieur, ou anneau intérieur)
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
    Private Function TraiterNombreComposante(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                             Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes

        Try
            'Définir la géométrie par défaut
            TraiterNombreComposante = New GeometryBag
            TraiterNombreComposante.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreComposante, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pGeometry = pFeature.Shape
                    pGeometry.Project(TraiterNombreComposante.SpatialReference)

                    'Interface pour extraire les anneaux
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pGeomColl.GeometryCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pGeometry.Envelope)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbComposantes=" & pGeomColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                            pGeometry, pGeomColl.GeometryCount)
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
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'anneau extérieur respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterNombreAnneauExterieur(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                  Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.

        Try
            'Définir la géométrie par défaut
            TraiterNombreAnneauExterieur = New GeometryBag
            TraiterNombreAnneauExterieur.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreAnneauExterieur, IGeometryCollection)

            'Si la géométrie est un Polygone
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour extraire les anneaux extérieurs
                    pPolygon = CType(pFeature.Shape, IPolygon4)
                    pPolygon.Project(TraiterNombreAnneauExterieur.SpatialReference)

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pPolygon.ExteriorRingCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon.Envelope)
                        'Écrire une erreur
                        EcrireFeatureErreur(pFeature.OID.ToString & "#NbAnneauExt=" & pPolygon.ExteriorRingCount.ToString & " /ExpReg=" & gsExpression, _
                                            pPolygon, pPolygon.ExteriorRingCount)
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
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'anneau intérieur respecte ou non l'expression régulière spécifiée.
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
    Private Function TraiterNombreAnneauInterieur(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                  Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)       'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieurs
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection

        Try
            'Définir la géométrie par défaut
            TraiterNombreAnneauInterieur = New GeometryBag
            TraiterNombreAnneauInterieur.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreAnneauInterieur, IGeometryCollection)

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
                    pPolygon.Project(TraiterNombreAnneauInterieur.SpatialReference)

                    'Vérifier si aucun polygone extérieur
                    If pPolygon.ExteriorRingCount > 0 Then
                        'Interface pour extraire les anneaux extérieurs
                        pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                        'Traiter tous les anneaux extérieurs
                        For i = 0 To pGeomCollExt.GeometryCount - 1
                            'Définir l'anneau extérieur
                            pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                            'Extraire tous les anneaux intérieurs
                            pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                            'Valider la valeur d'attribut selon l'expression régulière
                            oMatch = oRegEx.Match(pGeomCollInt.GeometryCount.ToString)

                            'Vérifier si on doit sélectionner l'élément
                            If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                                'Initialiser l'ajout
                                bAjouter = True
                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomSelColl.AddGeometry(pRingExt.Envelope)
                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbAnneauInt=" & pGeomCollInt.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                    RingToPolygon(pRingExt, pGeomCollInt), pGeomCollInt.GeometryCount)
                            End If
                        Next i
                    Else
                        'Initialiser l'ajout
                        bAjouter = True
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolygon.Envelope)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Aucun anneau extérieur /ExpReg=" & gsExpression, pPolygon, 0)
                    End If

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
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pPolygon = Nothing
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pRingExt = Nothing
        End Try
    End Function
#End Region
End Class
