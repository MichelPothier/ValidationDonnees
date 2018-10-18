Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsGeometrieValide.vb
'
'''<summary>
''' Classe qui permet de traiter les géométries valides.
''' 
''' La classe permet de traiter les trois attributs de géométrie VALIDE, VIDE, SIMPLE et FERMER.
''' 
''' VIDE : La valeur de l'attribut de la géométrie est 0:NON VIDE ou 1:VIDE.
''' SIMPLE : La valeur de l'attribut de la géométrie est 0:NON SIMPLE ou 1:SIMPLE.
''' FERMER : La valeur de l'attribut de la géométrie est 0:NON FERMER ou 1:FERMER.
''' VALIDE : La valeur de l'attribut de la géométrie est 0:NON VALIDE ou 1:VALIDE.
'''          VALIDE = NON_VIDE et SIMPLE.
''' 
''' Note : Une géométrie qui est non simple est une géométrie qui ne respecte pas le standard 
'''        de structure de la géométrie. 
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 15 avril 2015
'''</remarks>
''' 
Public Class clsGeometrieValide
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
            NomAttribut = "VALIDE"
            Expression = "VRAI"

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
            Nom = "GeometrieValide"
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
                'Définir le paramètre pour trouver les géométries NON VIDE
                ListeParametres.Add("VALIDE VRAI")
                'Définir le paramètre pour trouver les géométries NON VIDE
                ListeParametres.Add("VIDE FAUX")
                'Définir le paramètre pour trouver les géométries SIMPLE
                ListeParametres.Add("SIMPLE VRAI")
                'Définir le paramètre pour trouver les géométries FERMER
                ListeParametres.Add("FERMER VRAI")
                'Définir le paramètre pour trouver les géométries NON FERMER
                ListeParametres.Add("FERMER FAUX")
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
            If gsNomAttribut = "VIDE" Or gsNomAttribut = "SIMPLE" Or gsNomAttribut = "FERMER" Or gsNomAttribut = "VALIDE" Or gsNomAttribut = "CREER" Then
                'Vérifier si l'attribut est SIMPLE
                If gsNomAttribut = "SIMPLE" And gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    gsMessage = "ERREUR : L'attribut SIMPLE est invalide pour le type de géométrie Point."

                    'Vérifier si l'attribut est FERMER
                ElseIf gsNomAttribut = "FERMER" And gpFeatureLayerSelection.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPolyline Then
                    gsMessage = "ERREUR : L'attribut FERMER est invalide pour le type de géométrie qui n'est pas une Ligne."

                    'Sinon la contrainte est valide
                Else
                    'La contrainte est valide
                    AttributValide = True
                    gsMessage = "La contrainte est valide."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner la valeur de l'attribut sous forme de texte.
    ''' 
    '''<param name="pFeature">Interface contenant l'élément traité.</param>
    ''' 
    '''<return>Texte contenant la valeur de l'attribut 0:Faux ou 1:Vrai.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function ValeurAttribut(ByRef pFeature As IFeature) As String
        Try
            'Définir la valeur par défaut, la valeur est Fausse (0).
            ValeurAttribut = "FAUX"

            'Si le nom de l'attribut est VIDE
            If gsNomAttribut = "VIDE" Then
                'Vérifier si la géométrie est VIDE
                If pFeature.Shape.IsEmpty() Then ValeurAttribut = "VRAI"

                'Si le nom de l'attribut est SIMPLE
            ElseIf gsNomAttribut = "SIMPLE" Then
                'Vérifier si la géométrie est de type point
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                    'La géométrie est SIMPLE
                    ValeurAttribut = "VRAI"

                    'Si la géométrie n'est pas de type point
                Else
                    'Vérifier si la géométrie est simple
                    If fbGeometrieSimple(pFeature.ShapeCopy) Then ValeurAttribut = "VRAI"
                End If

                'Si le nom de l'attribut est FERMER
            ElseIf gsNomAttribut = "FERMER" Then
                'Interface pour vérifier si la géométrie est simple
                Dim pPolyline As IPolyline = CType(pFeature.Shape, IPolyline)
                'Vérifier si la géométrie est FERMER
                If pPolyline.IsClosed Then ValeurAttribut = "VRAI"
                'Vider la mémoire
                pPolyline = Nothing

                'Si le nom de l'attribut est VALIDE
            ElseIf gsNomAttribut = "VALIDE" Then
                'Vérifier si la géométrie est VIDE
                If pFeature.Shape.IsEmpty() Then Exit Function

                'Vérifier si la géométrie est de type point
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                    'La géométrie est SIMPLE
                    ValeurAttribut = "VRAI"

                    'Si la géométrie n'est pas de type point
                Else
                    'Vérifier si la géométrie est simple
                    If fbGeometrieSimple(pFeature.ShapeCopy) Then ValeurAttribut = "VRAI"
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
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

            'Si le nom de l'attribut est VIDE
            If gsNomAttribut = "VALIDE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterGeometrieValide(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est VIDE
            ElseIf gsNomAttribut = "VIDE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterGeometrieVide(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est SIMPLE
            ElseIf gsNomAttribut = "SIMPLE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterGeometrieSimple(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est FERMER
            ElseIf gsNomAttribut = "FERMER" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterGeometrieFermer(pFeatureCursor, pTrackCancel, bEnleverSelection)
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la géométrie respecte ou non l'état VIDE spécifié.
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
    Private Function TraiterGeometrieVide(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                          Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant le centre de la géométrie.
        Dim bSucces As Boolean = False                      'Indique si l'état VIDE recherché est un succès.
        Dim sMessage As String = Nothing                    'Contient le message d'erreur.

        Try
            'Définir la géométrie par défaut
            TraiterGeometrieVide = New GeometryBag
            TraiterGeometrieVide.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterGeometrieVide, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Vérifier si la géométrie est nulle
                    If pFeature.Shape Is Nothing Then
                        'Créer un nouveau point vide
                        pGeometry = New Multipoint
                        pGeometry.SpatialReference = TraiterGeometrieVide.SpatialReference
                        'Définir le message d'erreur
                        sMessage = " #Géométrie nulle="

                        'si la géométrie n'est pas nulle
                    Else
                        'Interface pour projeter
                        pGeometry = pFeature.ShapeCopy
                        pGeometry.Project(TraiterGeometrieVide.SpatialReference)
                        'Définir le message d'erreur
                        sMessage = " #Géométrie vide="
                    End If

                    'Vérifier si l'état recherché est un succès
                    bSucces = (pGeometry.IsEmpty) = (gsExpression = "VRAI")

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Définir le centre de la géomérie et transformer en multipoint
                        pMultipoint = GeometrieToMultiPoint(CentreGeometrie(pGeometry))
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pMultipoint)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & sMessage & pGeometry.IsEmpty.ToString & " /" & gsExpression, pMultipoint)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pGeometry = Nothing
                    pMultipoint = Nothing
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
            pGeometry = Nothing
            pMultipoint = Nothing
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la géométrie respecte ou non l'état SIMPLE spécifié.
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
    Private Function TraiterGeometrieSimple(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                            Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant la géométrie en erreur.
        Dim bSimple As Boolean = False                      'Indique si l'état SIMPLE est vrai.
        Dim bSucces As Boolean = False                      'Indique si l'état SIMPLE recherché est un succès.
        Dim sMessage As String = Nothing                    'Contient le message qui indique la raison de l'erreur.

        Try
            'Définir la géométrie par défaut
            TraiterGeometrieSimple = New GeometryBag
            TraiterGeometrieSimple.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterGeometrieSimple, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pGeometry = pFeature.ShapeCopy
                    pGeometry.Project(TraiterGeometrieSimple.SpatialReference)

                    'Vérifier si la géométrie est simple et retourner les différences
                    bSimple = fbGeometrieSimple(pGeometry, sMessage, pMultipoint)

                    'Vérifier si l'état recherché est un succès
                    bSucces = (bSimple) = (gsExpression = "VRAI")

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)

                        'Vérifier si les erreurs ne sont pas trouvés
                        If pMultipoint Is Nothing Then
                            'Définir le premier sommet de chaque composante par défaut
                            pMultipoint = PremierSommetComposanteGeometrie(pGeometry)
                        End If

                        'Ajouter la différence de la géométrie de l'élément simplifié
                        pGeomSelColl.AddGeometry(pMultipoint)

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Géométrie simple=" & bSimple.ToString & " /" & sMessage & " /" & gsExpression, pMultipoint)
                    End If

                    'Récupération de la mémoire disponble
                    pMultipoint = Nothing
                    GC.Collect()

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
            pMultipoint = Nothing
            bSimple = Nothing
            bSucces = Nothing
            sMessage = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la géométrie respecte ou non l'état FERMER spécifié.
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
    Private Function TraiterGeometrieFermer(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                            Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolyline As IPolyline = Nothing                'Interface pour projeter.
        Dim bSucces As Boolean = False                      'Indique si l'état FERMER recherché est un succès.

        Try
            'Définir la géométrie par défaut
            TraiterGeometrieFermer = New GeometryBag
            TraiterGeometrieFermer.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterGeometrieFermer, IGeometryCollection)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    pPolyline.Project(TraiterGeometrieFermer.SpatialReference)

                    'Vérifier si l'état recherché est un succès
                    bSucces = (pPolyline.IsClosed) = (gsExpression = "VRAI")

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter le dernier sommet de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolyline.ToPoint)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Géométrie fermée=" & pPolyline.IsClosed.ToString & " /" & gsExpression, _
                                            GeometrieToMultiPoint(pPolyline.ToPoint))
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
            bSucces = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la géométrie respecte ou non la validité de la géométrie.
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
    Private Function TraiterGeometrieValide(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                            Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant la géométrie en erreur.
        Dim bSimple As Boolean = False                      'Indique si l'état SIMPLE est vrai.
        Dim bSucces As Boolean = False                      'Indique si l'état VALIDE recherché est un succès.
        Dim sMessage As String = Nothing                    'Contient le message qui indique la raison de l'erreur.

        Try
            'Définir la géométrie par défaut
            TraiterGeometrieValide = New GeometryBag
            TraiterGeometrieValide.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterGeometrieValide, IGeometryCollection)

            'Si la géométrie est de type MultiPoint, Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Vérifier si la géométrie est nulle
                    If pFeature.Shape Is Nothing Then
                        'Créer un nouveau point vide
                        pGeometry = New Multipoint
                        pGeometry.SpatialReference = TraiterGeometrieValide.SpatialReference

                        'si la géométrie n'est pas nulle
                    Else
                        'Interface pour projeter
                        pGeometry = pFeature.ShapeCopy
                        pGeometry.Project(TraiterGeometrieValide.SpatialReference)
                    End If

                    'Vérifier si la géométrie est simple et retourner les différences
                    bSimple = fbGeometrieSimple(pGeometry, sMessage, pMultipoint)

                    'Vérifier si l'état recherché est un succès
                    bSucces = (bSimple And pGeometry.IsEmpty = False) = (gsExpression = "VRAI")

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)

                        'Vérifier si la géométrie est nulle
                        If pFeature.Shape Is Nothing Then
                            'Définir le centre de la géomérie et transformer en multipoint
                            pMultipoint = GeometrieToMultiPoint(CentreGeometrie(pGeometry))

                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Géométrie nulle=True" & " /" & gsExpression, pMultipoint)

                            'si la géométrie est vide
                        ElseIf pGeometry.IsEmpty Then
                            'Définir le centre de la géomérie et transformer en multipoint
                            pMultipoint = GeometrieToMultiPoint(CentreGeometrie(pGeometry))

                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Géométrie vide=" & pGeometry.IsEmpty.ToString & " /" & gsExpression, pMultipoint)

                            'si la géométrie n'est pas vide
                        Else
                            'Vérifier si les erreurs ne sont pas trouvés
                            If pMultipoint Is Nothing Then
                                'Définir le premier sommet de chaque composante par défaut
                                pMultipoint = PremierSommetComposanteGeometrie(pGeometry)
                            End If

                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Géométrie simple=" & bSimple.ToString & " /" & sMessage & " /" & gsExpression, pMultipoint)
                        End If

                        'Ajouter l'erreur de la géométrie de l'élément
                        pGeomSelColl.AddGeometry(pMultipoint)
                    End If

                    'Récupération de la mémoire disponble
                    pMultipoint = Nothing
                    GC.Collect()

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
            pMultipoint = Nothing
            bSimple = Nothing
            bSucces = Nothing
            sMessage = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de vérifier et corriger si la géométrie en paramètre est simple et retourner les différences.
    '''Ce traitement s'effectue seulement sur des Points, Multi-points, lignes et surfaces.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à vérifier et corriger.</param>
    '''<param name="sMessage">Contient la description de l'état de la géométrie.</param>
    '''<param name="pMultipoint">Interface ESRI contenant les sommets ou les différences sont présentes, nothing si aucune différence.</param>
    ''' 
    '''<returns>La fonction va retourner un "Boolean". Si le traitement n'a pas réussi le "Boolean" est à "False".</returns>
    '''
    Private Function fbGeometrieSimple(ByRef pGeometry As IGeometry, ByRef sMessage As String, Optional ByRef pMultipoint As IMultipoint = Nothing) As Boolean
        'Vérifier si la géométrie est présente
        If pGeometry Is Nothing Then Exit Function

        'Déclarer les variables de travail
        Dim pNewGeometry As IGeometry = Nothing             'Interface contenant la nouvelle géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface ESRI utilisée pour simplifier la géométrie.
        Dim pPointCollA As IPointCollection = Nothing       'Interface ESRI pour accéder aux sommets de la géométrie.
        Dim pGeomCollA As IGeometryCollection = Nothing     'Interface ESRI pour accéder aux composantes de la géométrie.
        Dim pPointCollB As IPointCollection = Nothing       'Interface ESRI pour accéder aux sommets de la géométrie.
        Dim pGeomCollB As IGeometryCollection = Nothing     'Interface ESRI pour accéder aux composantes de la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner une géométrie.

        'Définir la valeur de retour par défaut
        fbGeometrieSimple = True

        'Définir le message d'erreur par défaut
        sMessage = "La géométrie est simple"

        Try
            'Vérifier si le type de la géométrie est un multiPoint, ligne ou surface
            If pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour cloner la géométrie
                pClone = CType(pGeometry, IClone)
                'Cloner la nouvelle géométrie
                pNewGeometry = CType(pClone.Clone, IGeometry)
                'Interface pour simplifier le clone
                pTopoOp = CType(pNewGeometry, ITopologicalOperator2)
                'On ne sait pas si la géométrie clonée est simple
                pTopoOp.IsKnownSimple_2 = False
                'Simplifier la géométrie clonée
                pTopoOp.Simplify()
                'Redéfinir la nouvelle géométrie clonée et simplifiée
                pNewGeometry = CType(pTopoOp, IGeometry)

                'Interface pour vérifier la différence entre le nombre de sommets
                pPointCollA = CType(pGeometry, IPointCollection)
                pPointCollB = CType(pNewGeometry, IPointCollection)

                'Interface pour vérifier la différence entre le nombre de géométries
                pGeomCollA = CType(pGeometry, IGeometryCollection)
                pGeomCollB = CType(pNewGeometry, IGeometryCollection)

                'Interface pour vérifier si la géométrie est simple
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                'On ne sait pas si la géométrie est simple
                pTopoOp.IsKnownSimple_2 = False

                'Vérifier si la géométrie ne possède le même nombre de composantes que le clone
                If pGeomCollA.GeometryCount <> pGeomCollB.GeometryCount Then
                    'Indiquer que la géométrie n'est pas simple
                    fbGeometrieSimple = False

                    'Retourner la raison
                    sMessage = "Le nombre de composantes est différent : " & pGeomCollA.GeometryCount.ToString & "<>" & pGeomCollB.GeometryCount.ToString

                    'Définir le premier sommet de chaque composante par défaut
                    pMultipoint = PremierSommetComposanteGeometrie(pGeometry)

                    'Si la géométrie ne possède pas le même nombre de sommets que le clone
                ElseIf pPointCollA.PointCount <> pPointCollB.PointCount Then
                    'Indiquer que la géométrie n'est pas simple
                    fbGeometrieSimple = False

                    'Retourner la raison
                    sMessage = "Le nombre de sommets est différent : " & pPointCollA.PointCount.ToString & "<>" & pPointCollB.PointCount.ToString

                    'Transformer la géométrie en multipoint
                    pPointCollA = CType(GeometrieToMultiPoint(pGeometry), IPointCollection)

                    'Transformer le clone en multipoint
                    pPointCollB = CType(GeometrieToMultiPoint(CType(pTopoOp, IGeometry)), IPointCollection)

                    'Interface pour extraire les sommets différents entre la géométrie et le Clone simplifié
                    pTopoOp = CType(pPointCollA, ITopologicalOperator2)

                    'Extraire les sommets différents entre la géométrie et le Clone simplifié
                    pMultipoint = CType(pTopoOp.SymmetricDifference(CType(pPointCollB, IGeometry)), IMultipoint)

                    'Vérifier si le résultat de la différence est vide
                    If pMultipoint.IsEmpty Then
                        'Définir le premier sommet de chaque composante par défaut
                        pMultipoint = PremierSommetComposanteGeometrie(pGeometry)
                    End If

                    'Si la géométrie n'est pas simple mais possède le même nombre de sommets et de composantes que le clone
                ElseIf pTopoOp.IsSimple = False Then
                    'Transformer la géométrie en multipoint
                    pPointCollA = CType(GeometrieToMultiPoint(pGeometry), IPointCollection)

                    'Transformer le clone en multipoint
                    pPointCollB = CType(GeometrieToMultiPoint(CType(pTopoOp, IGeometry)), IPointCollection)

                    'Interface pour extraire les sommets différents entre la géométrie et le Clone simplifié
                    pTopoOp = CType(pPointCollA, ITopologicalOperator2)

                    'Extraire les sommets différents entre la géométrie et le Clone simplifié
                    pMultipoint = CType(pTopoOp.SymmetricDifference(CType(pPointCollB, IGeometry)), IMultipoint)

                    'Vérifier si le résultat de la différence est vide
                    If Not pMultipoint.IsEmpty Then
                        'Indiquer que la géométrie n'est pas simple
                        fbGeometrieSimple = False

                        'Retourner la raison
                        sMessage = "La géométrie n'est pas simple mais possède le même nombre de sommets et de composantes"
                    End If
                End If
            End If

            'Retourner la géométrie simplifier
            pGeometry = pNewGeometry

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewGeometry = Nothing
            pTopoOp = Nothing
            pPointCollA = Nothing
            pGeomCollA = Nothing
            pPointCollB = Nothing
            pGeomCollB = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de vérifier et corriger si la géométrie en paramètre est simple.
    '''Ce traitement s'effectue seulement sur des Points, Multi-points, lignes et surfaces.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à vérifier et corriger.</param>
    ''' 
    '''<returns>La fonction va retourner un "Boolean". Si le traitement n'a pas réussi le "Boolean" est à "False".</returns>
    '''
    Private Function fbGeometrieSimple(ByVal pGeometry As IGeometry) As Boolean
        'Vérifier si la géométrie est présente
        If pGeometry Is Nothing Then Exit Function

        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface ESRI utilisée pour simplifier la géométrie.
        Dim pPointCollA As IPointCollection = Nothing       'Interface ESRI pour accéder aux sommets de la géométrie.
        Dim pGeomCollA As IGeometryCollection = Nothing     'Interface ESRI pour accéder aux composantes de la géométrie.
        Dim pPointCollB As IPointCollection = Nothing       'Interface ESRI pour accéder aux sommets de la géométrie.
        Dim pGeomCollB As IGeometryCollection = Nothing     'Interface ESRI pour accéder aux composantes de la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner une géométrie.

        'Définir la valeur de retour par défaut
        fbGeometrieSimple = False

        Try
            'Vérifier si le type de la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Indiquer que la géométrie est simple
                fbGeometrieSimple = True

                'Vérifier si le type de la géométrie est un multiPoint, ligne ou surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour cloner la géométrie
                pClone = CType(pGeometry, IClone)
                pClone = pClone.Clone

                'Interface pour simplifier le clone
                pTopoOp = CType(pClone, ITopologicalOperator2)

                'On ne sait pas si la géométrie clonée est simple
                pTopoOp.IsKnownSimple_2 = False

                'Simplifier la géométrie clonée
                pTopoOp.Simplify()

                'Interface pour vérifier la géométrie originale
                pPointCollA = CType(pGeometry, IPointCollection)
                pGeomCollA = CType(pGeometry, IGeometryCollection)

                'Interface pour vérifier la géométrie clonée
                pPointCollB = CType(pTopoOp, IPointCollection)
                pGeomCollB = CType(pTopoOp, IGeometryCollection)
                pClone = CType(pTopoOp, IClone)

                'Vérifier si le clone possède les mêmes propriétés qu'à la géométrie
                If pClone.IsEqual(CType(pGeometry, IClone)) And pPointCollA.PointCount = pPointCollB.PointCount _
                And pGeomCollA.GeometryCount = pGeomCollB.GeometryCount Then
                    'Indiquer que la géométrie est simple
                    fbGeometrieSimple = True
                End If
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointCollA = Nothing
            pGeomCollA = Nothing
            pPointCollB = Nothing
            pGeomCollB = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire les sommets différents après avoir simplifié une géométrie.
    ''' On transforme la géométrie en multipoint et le clone en multipoint et effectue la différence.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à simplifier.</param>
    ''' 
    '''<returns>"IMultipoint" contenant les sommets différents. S'il n'y a pas de différence, le premier sommet est retourné.</returns>
    '''
    Private Function DifferenceSommetSimplifier(ByRef pGeometry As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pPointIdAware As IPointIDAware = Nothing        'Interface qui permet d'ajouter et conserver des identifiants sur les sommets.
        Dim pPointCollA As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets de la géométrie.
        Dim pPointCollB As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets du clone de la géométrie.
        Dim pMultiPointColl As IPointCollection = Nothing   'Interface qui permet d'accéder aux sommets du multipoint contenant les différences.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant les différences de sommets après avoir simplifié.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet de simplifier.
        Dim pClone As IClone = Nothing                      'Interface utilisé pour cloner la géométrie.

        'Définir la valeur de retour par défaut
        DifferenceSommetSimplifier = New Multipoint
        DifferenceSommetSimplifier.SpatialReference = pGeometry.SpatialReference

        Try
            'Conserver une copie de la géométrie originale
            pClone = CType(pGeometry, IClone)
            pClone = pClone.Clone()

            'Simplifier le clone
            pTopoOp = CType(pClone, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Transformer la géométrie en multipoint
            pPointCollA = CType(GeometrieToMultiPoint(pGeometry), IPointCollection)

            'Transformer le clone en multipoint
            pPointCollB = CType(GeometrieToMultiPoint(CType(pTopoOp, IGeometry)), IPointCollection)

            'Extraire les sommets différents de la géométrie du Clone
            pTopoOp = CType(pPointCollA, ITopologicalOperator2)
            pMultipoint = CType(pTopoOp.SymmetricDifference(CType(pPointCollB, IGeometry)), IMultipoint)

            'Définir la valeur de retour des sommets différents
            DifferenceSommetSimplifier = pMultipoint

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPointIdAware = Nothing
            pPointCollA = Nothing
            pPointCollB = Nothing
            pMultipointColl = Nothing
            pMultipoint = Nothing
            pTopoOp = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire les sommets différents après avoir simplifié une géométrie.
    ''' On ajoute un identifiant sur chaque sommet et on vérifie si les mêmes identifiants sont 
    ''' encore présents après avoir simplifié. On enlève les Identifiant à la fin.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à simplifier.</param>
    ''' 
    '''<returns>"IMultipoint" contenant les sommets différents. S'il n'y a pas de différence, il est vide. "Nothing" sinon</returns>
    '''
    Private Function DifferenceIdSommetSimplifier(ByRef pGeometry As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pPointIdAware As IPointIDAware = Nothing        'Interface qui permet d'ajouter et conserver des identifiants sur les sommets.
        Dim pPointCollA As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets de la géométrie.
        Dim pPointCollB As IPointCollection = Nothing       'Interface qui permet d'accéder aux sommets du clone de la géométrie.
        Dim pMultipointColl As IPointCollection = Nothing   'Interface qui permet d'accéder aux sommets du multipoint contenant les différences.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant les différences de sommets après avoir simplifié.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet de simplifier.
        Dim pPoint As IPoint = Nothing                      'Interface contenant un sommet.
        Dim pClone As IClone = Nothing                      'Interface utilisé pour cloner la géométrie.
        Dim qCollection As New Collection                   'Objet contenant les identifiants des sommets traités.
        Dim iID As Integer = Nothing                        'Compteur d'identifiant

        'Définir la valeur de retour par défaut
        DifferenceIdSommetSimplifier = Nothing

        Try
            'Interface pour créer un multipoint vide contenant les différences
            pMultipoint = New Multipoint
            pMultipoint.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter les sommets dans le multipoints
            pMultipointColl = CType(pMultipoint, IPointCollection)

            'Interface pour permettre d'ajouter des IDs aux sommets
            pPointIdAware = CType(pGeometry, IPointIDAware)
            pPointIdAware.PointIDAware = True
            pPointCollA = CType(pGeometry, IPointCollection)

            'Conserver une copie de la géométrie originale
            pClone = CType(pPointCollA, IClone)
            pClone = pClone.Clone()
            pPointCollB = CType(pClone, IPointCollection)

            'Ajouter tous les ID sur tous les sommets de la géométrie
            For i = 0 To pPointCollA.PointCount - 1
                'Interface pour ajouter le ID au sommet
                pPoint = pPointCollA.Point(i)
                'Définir le ID
                pPoint.ID = i + 1
                'Ajouter le ID dans la collection
                qCollection.Add(i, Str(pPoint.ID))
                'Mise à jour du ID
                pPointCollA.UpdatePoint(i, pPoint)
            Next i

            'Simplifier la géométrie
            pTopoOp = CType(pPointCollA, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Traiter tous les sommets conservés et ajoutés
            For i = 0 To pPointCollA.PointCount - 1
                'Interface pour ajouter le ID au sommet
                pPoint = pPointCollA.Point(i)
                'Vérifier si le ID est présent
                If qCollection.Contains(Str(pPoint.ID)) Then
                    'Retirer les numéro de sommets présents dans le multipoint
                    qCollection.Remove(Str(pPoint.ID))
                Else
                    'Ajouter le point ajouté comme étant différent
                    pMultipointColl.AddPoint(pPoint)
                End If
            Next i

            'Ajouter tous les points restants correspondants aux points éliminés
            For Each iID In qCollection
                'Interface pour ajouter le ID au sommet
                pPoint = pPointCollB.Point(iID)
                'Ajouter le point ajouté comme étant différent
                pMultipointColl.AddPoint(pPoint)
            Next

            'Enlever les IDs
            pPointIdAware.DropPointIDs()
            pPointIdAware.PointIDAware = False

            'Définir la valeur de retour des sommets différents
            DifferenceIdSommetSimplifier = pMultipoint

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPointIdAware = Nothing
            pPointCollA = Nothing
            pPointCollB = Nothing
            pMultipointColl = Nothing
            pMultipoint = Nothing
            pTopoOp = Nothing
            pClone = Nothing
            pPoint = Nothing
            qCollection = Nothing
            iID = Nothing
        End Try
    End Function
#End Region
End Class
