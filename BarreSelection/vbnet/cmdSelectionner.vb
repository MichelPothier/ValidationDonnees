Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.DataSourcesRaster
Imports System.IO

'**
'Nom de la composante : cmdSelectionner.vb
'
'''<summary>
''' Commande qui permet de sélectionner les éléments de la classe de sélection selon les paramètres spécifiés.
''' 
''' Seuls les éléments sélectionnés sont traités.
''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Public Class cmdSelectionner
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    'Déclarer les variables de travail
    'Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    'Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        'Définir les variables de travail

        Try
            'Par défaut la commande est inactive
            Enabled = False
            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                m_Application = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Définir le document
                    m_MxDocument = CType(m_Application.Document, IMxDocument)
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim pSpatialRefTol As ISpatialReferenceTolerance    'Interface contenant la tolérance XY de la référence spatiale.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSpatialFilter As ISpatialFilter = Nothing  'Interface contenant la requête spatiale.
        Dim pSelectionSet As ISelectionSet = Nothing    'Interface contenant les éléments sélectionnés.
        Dim oRequete As intRequete = Nothing            'Objet utilisé pour effectuer la sélection selon la requête.
        Dim pPoint As IPoint = Nothing                  'Point utilisé pour centrer l'enveloppe.
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'enveloppe de la géométrie.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour extraire le nombre de composante d'une géométrie.
        Dim pTrackCancel As ITrackCancel = Nothing      'Interface qui permet d'annuler la sélection avec la touche ESC.
        Dim pMouseCursor As IMouseCursor = Nothing      'Interface qui permet de changer le curseur de la souris.
        Dim iMsgBoxStyle As MsgBoxStyle = 0             'Indique le style du MsgBox utilisé.
        Dim bDecoupage As Boolean = False               'Indique si un polygon de découpage est utilisé.
        Dim qStartTime As DateTime = Nothing            'Date de départ.
        Dim qEndTime As DateTime = Nothing              'Date de fin.
        Dim qElapseTime As TimeSpan = Nothing           'Temps de traitement.

        Try
            'Initialiser le temps d'exécution
            qStartTime = DateTime.Now
            'Initialiser le style du MsgBox
            iMsgBoxStyle = MsgBoxStyle.Information

            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Vider la mémoire
            m_GeometrieSelection = Nothing

            'Vérifier si la Featureclass est valide
            If m_FeatureLayer.FeatureClass Is Nothing Then
                'Afficher un message d'erreur
                MsgBox("ERREUR : La FeatureClass de sélection est invalide : " & m_FeatureLayer.Name)
                Exit Sub
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(m_FeatureLayer, IFeatureSelection)
            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Définir la requête spatiale
                pSpatialFilter = New SpatialFilter
                pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                pSpatialFilter.OutputSpatialReference(m_FeatureLayer.FeatureClass.ShapeFieldName) = m_MxDocument.FocusMap.SpatialReference
                pSpatialFilter.Geometry = m_MxDocument.ActiveView.Extent
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver la sélection de départ
            pSelectionSet = pFeatureSel.SelectionSet

            'Permettre d'annuler la sélection avec la touche ESC
            pTrackCancel = New CancelTracker
            pTrackCancel.CancelOnKeyPress = True
            pTrackCancel.CancelOnClick = False

            'Permettre l'affichage de la barre de progression
            m_Application.StatusBar.ProgressBar.Position = 0
            m_Application.StatusBar.ShowProgressBar("Selection en cours ...", 0, pFeatureSel.SelectionSet.Count, 1, True)
            pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar

            'Interface pour extraire la précision de la référence spatiale
            pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
            m_Precision = pSpatialRefTol.XYTolerance
            m_MenuParametresSelection.txtPrecision.Text = m_Precision.ToString("0.0#######")

            'Définir la requête spécifiée
            oRequete = DefinirRequete()

            'Indiquer si un découpage est utilisé
            bDecoupage = oRequete.LimiteDecoupage IsNot Nothing

            'Vérifier si la requête est valide
            If oRequete.EstValide Then
                'Écrire les statistiques d'utilisation
                Call EcrireStatistiqueUtilisation(oRequete.Nom & " " & oRequete.FeatureLayerSelection.Name & " " & oRequete.Parametres)

                'Sélectionner les éléments selon les paramètres spécifiés et définir les géométries trouvées
                m_GeometrieSelection = oRequete.Selectionner(pTrackCancel, m_TypeSelection = "Enlever")

                'Vérifier si des éléments sont sélectionnés
                If m_GeometrieSelection.IsEmpty Then
                    'Rafraichier l'affichage des éléments sélectionnés
                    m_MxDocument.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, m_MxDocument.ActiveView.Extent)

                Else
                    'Vérifier si on doit afficher la classe d'erreur
                    If m_CreerClasseErreur Then
                        'Afficher la FeatureClass d'erreur dans la Map
                        AfficherFeatureClassErreur(oRequete.Map, oRequete.FeatureClassErreur, oRequete.FeatureClassErreur.AliasName & " " & oRequete.Parametres)
                    End If

                    'Vérifier si on doit faire un Zoom selon les géométries d'erreurs
                    If m_ZoomGeometrieErreur Then
                        'Définir l'enveloppe
                        pEnvelope = m_GeometrieSelection.Envelope

                        'Vérifier s'il faut centrer
                        If pEnvelope.Height = 0 And pEnvelope.Width = 0 Then
                            'Définir le point pour centrer l'enveloppe
                            pPoint = pEnvelope.LowerLeft
                            'Définir l'enveloppe
                            pEnvelope = m_MxDocument.ActiveView.Extent
                            'Centrer l'enveloppe courante
                            pEnvelope.CenterAt(pPoint)
                        Else
                            'Agrandir l'enveloppe de 10% de l'élément en erreur
                            pEnvelope.Expand(pEnvelope.Width / 10, pEnvelope.Height / 10, False)
                        End If

                        'Changer l'enveloppe
                        m_MxDocument.ActiveView.Extent = pEnvelope
                        'Afficher tous les éléments sélectionnés
                        m_MxDocument.ActiveView.Refresh()

                        'Si on ne doit pas faire un Zoom selon les géométries d'erreurs
                    Else
                        'Rafraichier l'affichage des éléments sélectionnés
                        m_MxDocument.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, m_MxDocument.ActiveView.Extent)
                    End If
                End If
            Else
                'Créer une géométrie de sélection vide
                m_GeometrieSelection = New GeometryBag
                'Définir le msgBoxStyle d'erreur
                iMsgBoxStyle = MsgBoxStyle.Critical
            End If

            'Cacher la barre de progression
            pTrackCancel.Progressor.Hide()

            'Interface pour extraire le nombre de composantes d'une géométrie
            pGeomColl = CType(m_GeometrieSelection, IGeometryCollection)

            'Définir le temps d'exécution
            qEndTime = DateTime.Now
            qElapseTime = qEndTime.Subtract(qStartTime)

            'Interface pour extraire la précision de la référence spatiale
            pSpatialRefTol = CType(oRequete.SpatialReference, ISpatialReferenceTolerance)

            'Afficher le résultat obtenu
            MsgBox(oRequete.Message & vbCrLf _
                   & "-Référence spatiale : " & oRequete.SpatialReference.FactoryCode.ToString & ":" & oRequete.SpatialReference.Name & vbCrLf _
                   & "-Précision de la référence spatiale : " & pSpatialRefTol.XYTolerance.ToString("0.0#######") & vbCrLf _
                   & "-Limite du découpage utilisée : " & bDecoupage.ToString & vbCrLf _
                   & "-Nombre d'éléments traités : " & pSelectionSet.Count.ToString & vbCrLf _
                   & "-Nombre d'éléments sélectionnés : " & pFeatureSel.SelectionSet.Count.ToString & vbCrLf _
                   & "-Nombre de géométries trouvées : " & pGeomColl.GeometryCount.ToString & vbCrLf _
                   & "-Temps d'exécution : " & qElapseTime.ToString(), iMsgBoxStyle)

        Catch erreur As Exception
            'Vérifier si le traitement a été annulé
            If TypeOf (erreur) Is CancelException Then
                'Vérifier si la barre de progression est active
                If pTrackCancel IsNot Nothing Then
                    'Cacher la barre de progression
                    pTrackCancel.Progressor.Hide()
                    'Afficher le message
                    m_Application.StatusBar.Message(0) = erreur.Message
                End If
                'Rafraichier l'affichage des éléments sélectionnés
                m_MxDocument.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, m_MxDocument.ActiveView.Extent)
            Else
                'Message d'erreur
                MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
            End If

        Finally
            'Vider la mémoire
            pSpatialRefTol = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pSpatialFilter = Nothing
            oRequete = Nothing
            pPoint = Nothing
            pEnvelope = Nothing
            pGeomColl = Nothing
            pTrackCancel = Nothing
            pMouseCursor = Nothing
            iMsgBoxStyle = Nothing
            bDecoupage = Nothing
            qStartTime = Nothing
            qEndTime = Nothing
            qElapseTime = Nothing
            'Récupération de la mémoire disponble
            GC.Collect()
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si le FeatureLayer de sélection est invalide
        If m_FeatureLayer Is Nothing Then
            'Rendre désactive la commande
            Me.Enabled = False

            'Si le FeatureLayer de sélection est invalide
        Else
            'Rendre active la commande
            Me.Enabled = True
        End If
    End Sub

    '''<summary>
    ''' Routine qui permet d'afficher la FeatureClass d'erreur dans la Map et dans le TableWindow d'attributs.
    ''' 
    '''<param name="pMap"> Interface contenant la Map dans lequel le FeatureLayer sera ajouté.</param>
    '''<param name="pFeatureClass"> FeatureClass à ajouter dans la Map.</param>
    '''<param name="sNom"> Nom du FeatureLayer à ajouter dans la Map.</param>
    ''' 
    '''</summary>
    '''
    Protected Sub AfficherFeatureClassErreur(ByVal pMap As IMap, ByVal pFeatureClass As IFeatureClass, ByVal sNom As String)
        'Déclarer les variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface ESRI contenant la Featureclass des éléments en mémoire.
        Dim pTableWindow2 As ITableWindow2 = Nothing        'Interface qui permet de verifier la présence du menu des tables d'attributs et de les manipuler.
        Dim pExistTableWindow As ITableWindow = Nothing     'Interface contenant le menu de la table d'attributs existente.
        'Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing  'Interface pour extraire le Renderer contenant la symbologie.
        'Dim pSimpleRenderer As ISimpleRenderer = Nothing    'Interface contenant la symbologie.

        Try
            'Sortie si la classe d'erreur est absente
            If pFeatureClass Is Nothing Then Exit Sub

            'Créer un nouveau FeatureLayer
            pFeatureLayer = New FeatureLayer
            'Définir le nom du FeatureLayer selon le nom et la date
            pFeatureLayer.Name = sNom
            'Rendre visible le FeatureLayer
            pFeatureLayer.Visible = True
            ''Interface pour changer la symbologie
            'pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
            'pSimpleRenderer = CType(pGeoFeatureLayer.Renderer, ISimpleRenderer)
            'Définir la Featureclass dans le FeatureLayer
            pFeatureLayer.FeatureClass = m_Requete.Requete.FeatureClassErreur
            'Ajouter le FeatureLayer dans la Map
            pMap.AddLayer(pFeatureLayer)

            'Vérifier si on doit afficher la table d'erreur
            If m_Requete.Requete.AfficherTableErreur = False Then Exit Sub

            'Interface pour vérifier la présence du menu des tables d'attributs
            pTableWindow2 = CType(New TableWindow, ITableWindow2)

            'Définir le menu de la table d'attribut de la table s'il est présent
            pExistTableWindow = pTableWindow2.FindViaLayer(pFeatureLayer)

            'Vérifier si le menu de la table d'attribut est absent
            If pExistTableWindow Is Nothing Then
                'Définir le FeatureLayer à afficher
                pTableWindow2.Layer = pFeatureLayer
            End If

            'Vérifier le menu de la table d'attribut est absent
            If pExistTableWindow Is Nothing Then
                'Définir les paramètre d'affichage du menu des tables d'attributs
                pTableWindow2.TableSelectionAction = esriTableSelectionActions.esriSelectFeatures
                pTableWindow2.ShowSelected = False
                pTableWindow2.ShowAliasNamesInColumnHeadings = True
                pTableWindow2.Application = m_Requete.Requete.Application

                'Si le menu de la table d'attribut est présent
            Else
                'Redéfinir le menu des tables d'attributs pour celui existant
                pTableWindow2 = CType(pExistTableWindow, ITableWindow2)
            End If

            'Afficher le menu des tables d'attributs s'il n'est pas affiché
            If Not pTableWindow2.IsVisible Then pTableWindow2.Show(True)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            pTableWindow2 = Nothing
            pExistTableWindow = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les statistiques d'utilisation d'un usager.
    ''' 
    '''<param name="sCommande"> Commande à écrire dans le fichier de statistique d'utilisation.</param>
    '''<param name="sNomRepertoire"> Nom du répertoire dans lequel le fichier de statistique est présent.</param>
    ''' 
    '''</summary>
    '''
    Protected Sub EcrireStatistiqueUtilisation(ByVal sCommande As String, Optional ByVal sNomRepertoire As String = "S:\Developpement\geo\")
        'Déclarer les variables de travail
        Dim oStreamWriter As StreamWriter = Nothing     'Objet utilisé pour écrire dans un fichier text.
        Dim sNomFichier As String = ""                  'Nom complet du fichier de statistique d'utilisation.
        Dim sNomUsager As String = ""                   'Nom de l'usager.

        Try
            'Définir le nom de l'usager
            sNomUsager = Environment.GetEnvironmentVariable("USERNAME")

            'Définir le nom complet du fichier
            sNomFichier = sNomRepertoire & sNomUsager & ".txt"

            'Vérifier si le fichier existe
            If File.Exists(sNomFichier) Then
                'Définir l'objet pour écrire à la fin du fichier
                oStreamWriter = File.AppendText(sNomFichier)

                'Si le fichier n'existe pas
            Else
                'Définir l'objet pour écrire dans un nouveau fichier créé
                oStreamWriter = File.CreateText(sNomFichier)

                'Écrire l'entête du fichier
                oStreamWriter.WriteLine("Date, 	 Env, 	 Usager, 	 UsagerBD, 	 UsagerSIB, 	 Outil")
            End If

            'Écrire la commande utilisée
            oStreamWriter.WriteLine(DateTime.Now.ToString & "," & vbTab & "ARCMAP" & "," & vbTab & sNomUsager & "," & vbTab & "NONE," & vbTab & "NONE," & vbTab & sCommande)

            'Fermer le fichier
            oStreamWriter.Close()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oStreamWriter = Nothing
        End Try
    End Sub
End Class
