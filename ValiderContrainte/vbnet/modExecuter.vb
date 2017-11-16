Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

'**
'Nom de la composante : modExecuter.vb 
'
'''<summary>
''' Module principale permettant l'exécution de la validation des contraintes d'intégrité présentes dans une table de contraintes.
''' 
''' Les 11 paramètres de la ligne de commande du programme sont les suivants:
''' -------------------------------------------------------------------------
''' 1-Geodatabase des classes: Nom de la Géodatabase contenant les classes à traiter.
'''                            Obligatoire, Défaut : 
''' 2-Table des contraintes  : Nom de la table contenant l'information sur les contraintes d'intégrité spatiales.
'''                            Obligatoire, Défaut : CONTRAINTE_INTEGRITE_SPATIALE
''' 3-Filtre des contraintes : Filtre effectué sur la table des contraintes pour lesquelles on veut effectuer le traitement.
'''                            Optionnel, Défaut : # (Aucun filtre)
''' 4-Layer de découpage     : Nom du Layer de découpage dans lequel les éléments de découpage sélectionnés seront utilisés pour effectuer le traitement.
'''                            Optionnel, Défaut : 
''' 5-Attribut de découpage  : Nom de l'attribut du Layer de découpage dans lequel les identifiants seront utilisés pour effectuer le traitement.
'''                            Optionnel, Défaut : 
''' 6-Répertoire des erreurs : Nom du répertoire dans lequel les erreurs seront écrites.
'''                            Optionnel, Défaut : 
''' 7-Rapport d'erreurs      : Nom du rapport dans lequel l'information sur les erreurs sera écrite.
'''                            Optionnel, Défaut : 
''' 8-Fichier journal        : Nom du fichier journal dans lequel l'information sur l'exécution du traitement sera écrite.
'''                            Optionnel, Défaut : 
''' 9-Courriel               : Addresses courriels des personnes à qui on veut envoyer le rapport d'erreurs.
'''                            Optionnel, Défaut : 
''' 10-Table des statistiques: Nom de la table contenant l'information sur les statistiques d'erreurs et de traitements.
'''                            Optionnel, Défaut : 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 09 mai 2016
''' 
''' Pour augmenter la mémoire en 32-bits après avoir compilé le programme:
''' "C:\Program Files (x86)\Microsoft Visual Studio 12.0\VC\bin\editbin.exe" /LARGEADDRESSAWARE S:\applications\gestion_bdg\pro\Geotraitement\exe\ValiderContrainte.exe
'''</remarks>
''' 
Module modExecuter
    '''<summary> Interface d'initialisation des licences ESRI.</summary>
    Private m_AOLicenseInitializer As LicenseInitializer = New ValiderContrainte.LicenseInitializer()
    '''<summary>Nom de la Géodatabase contenant les classes à valider.</summary>
    Private m_NomGeodatabaseClasses As String = "Database Connections\BDRS_PRO_BDG.sde"
    '''<summary>Nom de la table contenant l'information sur les contraintes d'intégrité.</summary>
    Private m_NomTableContraintes As String = "CONTRAINTE_INTEGRITE_SPATIALE"
    '''<summary>Filter des contraintes à traiter.</summary>
    Private m_FiltreContraintes As String = ""
    '''<summary>Nom de la classe de découpage dans lequel les éléments de découpage sélectionnés seront utilisés pour effectuer le traitement.</summary>
    Private m_NomClasseDecoupage As String = ""
    '''<summary>Nom de l'attribut du Layer de découpage dans lequel les identifiants seront utilisés pour effectuer le traitement.</summary>
    Private m_NomAttributDecoupage As String = ""
    '''<summary>Nom du répertoire dans lequel les erreurs seront écrites.</summary>
    Private m_NomRepertoireErreurs As String = ""
    '''<summary>Nom du rapport d'erreurs dans lequel seront écrites les statistiques d'erreurs.</summary>
    Private m_NomRapportErreurs As String = ""
    '''<summary>Nom du fichier journal dans lequel l'exécution du traitement sera inscrit.</summary>
    Private m_NomFichierJournal As String = ""
    '''<summary>Addresses courriels des personnes à qui on veut envoyer le rapport d'exécution du traitement.</summary>
    Private m_Courriel As String = ""
    '''<summary>Nom de la table contenant l'information sur les statistiques d'erreurs et de traitements.</summary>
    Private m_NomTableStatistiques As String = ""
    '''<summary>Interface qui permet d'annuler l'exécution du traitement en inteactif.</summary>
    Private m_TrackCancel As ITrackCancel = Nothing

    <STAThread()> _
    Sub Main()
        'Déclarer les variables de travail
        Dim oTableContraintes As clsTableContraintes = Nothing  'Objet qui permet d'exécuter la validation des contraintes d'intégrité.

        Try
            'ESRI License Initializer generated code.
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeStandard}, _
                                                         New esriLicenseExtensionCode() {})

            'Valider les parametres de la ligne de commandes
            Call ValiderParametres()

            'Permettre d'annuler la sélection avec la touche ESC
            m_TrackCancel = New CancelTracker
            m_TrackCancel.CancelOnKeyPress = False
            m_TrackCancel.CancelOnClick = False

            'Objet qui permet d'exécuter le traitement de validation des contraintes d'intégrité
            oTableContraintes = New clsTableContraintes(m_NomGeodatabaseClasses, m_NomTableContraintes, m_FiltreContraintes)
            oTableContraintes.NomClasseDecoupage = m_NomClasseDecoupage
            oTableContraintes.NomAttributDecoupage = m_NomAttributDecoupage
            oTableContraintes.NomRepertoireErreurs = m_NomRepertoireErreurs
            oTableContraintes.NomRapportErreurs = m_NomRapportErreurs
            oTableContraintes.NomFichierJournal = m_NomFichierJournal
            oTableContraintes.Courriel = m_Courriel
            oTableContraintes.NomTableStatistiques = m_NomTableStatistiques
            oTableContraintes.TrackCancel = m_TrackCancel
            oTableContraintes.Commande = System.Environment.CommandLine

            'Exécuter le traitement de validation des contraintes d'intégrité
            oTableContraintes.Executer()

            'Retourner le code d'exéution du traitement
            System.Environment.Exit(oTableContraintes.CodeErreur)

        Catch ex As Exception
            'Afficher l'erreur
            Console.WriteLine(ex.Message)
            'Écrire le message d'erreur
            'File.AppendAllText("D:\Erreur_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log", ex.Message & vbCrLf)
            'Retourner le code d'échec du traitement
            System.Environment.Exit(-1)
        Finally
            'Vider la mémoire
            oTableContraintes = Nothing
            'Fermer adéquatement l'application utilisant les licences ESRI
            m_AOLicenseInitializer.ShutdownApplication()
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider et définir les paramètres du programme.
    '''</summary>
    '''
    Sub ValiderParametres()
        'Déclaration des variables de travail
        Dim args() As String = System.Environment.GetCommandLineArgs()  'Contient les paramètres de la ligne de commandes

        Try
            'Valider le paramètre de la Geodatabase des classes à valider.
            Call ValiderParametreGeodatabaseClasses(args)

            'Valider le paramètre de la table contenant les contraintes d'intégrité.
            Call ValiderParametreTableContraintes(args)

            'Valider le paramètre du filtre des contraintes.
            Call ValiderParametreFiltreContraintes(args)

            'Valider le paramètre contenant les identifiants de découpage à valider
            Call ValiderParametreClasseDecoupage(args)

            'Valider le paramètre de l'attribut contenant l'identifiant de découpage à valider
            Call ValiderParametreAttributDecoupage(args)

            'Valider le paramètre du répertoire dans lequel les erreurs seront écrites
            Call ValiderParametreRepertoireErreurs(args)

            'Valider le paramètre du rapport dans lequel les statistiques d'erreurs seront écrites
            Call ValiderParametreRapportErreurs(args)

            'Valider le paramètre du fichier journal dans lequel l'information sur l'exécution du traitement sera écrite.
            Call ValiderParametreFichierJournal(args)

            'Valider le paramètre des courriels qui recevront le rapport d'exécution du traitement
            Call ValiderParametreCourriels(args)

            'Valider le paramètre de la table contenant les statistiques.
            Call ValiderParametreTableStatistiques(args)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            args = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du nom de la Geodatabase des classes à valider.
    '''</summary>
    '''
    Sub ValiderParametreGeodatabaseClasses(ByVal args() As String)
        Try
            'Vérifier si le paramètre de la Geodatabase des classes est présent.
            If args.Length - 1 > 0 Then
                'Définir le nom de la Géodatbase des classes
                m_NomGeodatabaseClasses = args(1)

            Else
                'Retourner l'erreur
                Err.Raise(-1, "ValiderParametreGeodatabaseClasses", "ERREUR : Le paramètre de la Geodatabase des classes est absent !")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre de la table des contraintes d'intégrité.
    '''</summary>
    '''
    Sub ValiderParametreTableContraintes(ByVal args() As String)
        Try
            'Vérifier si le paramètre de la table des contraintes d'intégrité est présent.
            If args.Length - 1 > 1 Then
                'Définir le nom de la table des contraintes
                m_NomTableContraintes = args(2)

            Else
                'Retourner l'erreur
                Err.Raise(-1, "ValiderParametreTableContraintes", "ERREUR : Le paramètre de la table des contraintes est absent !")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du filtre des contraintes à valider.
    '''</summary>
    '''
    Sub ValiderParametreFiltreContraintes(ByVal args() As String)
        Try
            'Vérifier si le paramètre de la liste des groupes est présent.
            If args.Length - 1 > 2 Then
                'Vérifier si le paramètre contient un filtre des contraintes
                If args(3) <> "#" Then
                    'Définir le filtre des contraintes
                    m_FiltreContraintes = args(3)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du nom du Layer contenant les identifiants de découpage à valider.
    '''</summary>
    '''
    Sub ValiderParametreClasseDecoupage(ByVal args() As String)
        Try
            'Valider le paramètre du nom du Layer contenant les identifiants de découpage à valider
            If args.Length - 1 > 3 Then
                'Vérifier si le paramètre contient le nom de la classe de découpage
                If args(4) <> "#" Then
                    'Définir le nom de la classe de découpage 
                    m_NomClasseDecoupage = args(4)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du nom de l'attribut contenant l'identifiant de découpage à valider.
    '''</summary>
    '''
    Sub ValiderParametreAttributDecoupage(ByVal args() As String)
        Try
            'Valider le paramètre du nom de l'attribut contenant l'identifiant de découpage à valider
            If args.Length - 1 > 4 Then
                'Vérifier si le paramètre contient le nom de l'attribut de découpage
                If args(5) <> "#" Then
                    'Définir le nom de l'attribut de découpage
                    m_NomAttributDecoupage = args(5)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du répertoire dans lequel les erreurs seront écrites.
    '''</summary>
    '''
    Sub ValiderParametreRepertoireErreurs(ByVal args() As String)
        Try
            'Valider le paramètre du répertoire ou de la géodatabase dans lequel les erreurs seront écrites
            If args.Length - 1 > 5 Then
                'Vérifier si le paramètre contient le nom du répertoire ou de la géodatabase d'erreurs
                If args(6) <> "#" Then
                    'Définir le nom du répertoire ou de la géodatabase d'erreurs
                    m_NomRepertoireErreurs = args(6)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du rapport dans lequel les statistiques d'erreurs seront écrites.
    '''</summary>
    '''
    Sub ValiderParametreRapportErreurs(ByVal args() As String)
        Try
            'Valider le paramètre du rapport d'erreurs
            If args.Length - 1 > 6 Then
                'Vérifier si le paramètre contient le nom du rapport d'erreurs
                If args(7) <> "#" Then
                    'Définir le nom du rapport d'erreurs
                    m_NomRapportErreurs = args(7)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre du fichier journal dans lequel l'information sur l'exécution du traitement sera écrite.
    '''</summary>
    '''
    Sub ValiderParametreFichierJournal(ByVal args() As String)
        Try
            'Valider le paramètre du répertoire dans lequel les erreurs seront écrites
            If args.Length - 1 > 7 Then
                'Vérifier si le paramètre contient le nom du fichier journal
                If args(8) <> "#" Then
                    'Définir le nom du fichier journal
                    m_NomFichierJournal = args(8)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre des courriels qui recevront le rapport d'exécution du traitement.
    '''</summary>
    '''
    Sub ValiderParametreCourriels(ByVal args() As String)
        Try
            'Valider le paramètre des courriels qui recevront le rapport d'exécution du traitement
            If args.Length - 1 > 8 Then
                'Vérifier si le paramètre contient les courriels pour lesquels on veut envoyer le rapport d'erreurs
                If args(9) <> "#" Then
                    'Définir les courriels pour lesquels on veut envoyer le rapport d'erreurs
                    m_Courriel = args(9)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le paramètre de la table des statistiques.
    '''</summary>
    '''
    Sub ValiderParametreTableStatistiques(ByVal args() As String)
        Try
            'Valider le paramètre de la table des statistiques
            If args.Length - 1 > 9 Then
                'Vérifier si le paramètre contient le nom de la table des statistiques
                If args(10) <> "#" Then
                    'Définir le nom de la table des statistiques
                    m_NomTableStatistiques = args(10)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub
End Module
