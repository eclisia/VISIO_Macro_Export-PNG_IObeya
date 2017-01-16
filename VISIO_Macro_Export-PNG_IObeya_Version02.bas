Attribute VB_Name = "Export"
'Am�lioration � apporter sur cette macro :
'   1 - D�finir les chemins des images en tant que Variable
'   2 - Reconnaitre la page automatiquement avec le chemin de l'image afin
'       d 'avoir une v�rification si l'on applique la mauvaise macro � la mauvaise page/image
'   3 - Inclure dans une page de garde du visio, des formes de types " texte " qui serviront � :
'           - stocker dans un texte, le chemin relative du lieu de stockage des images
'           - stocker dans un texte, le nom des images utilis�es dans les macro.
'       L 'objectif de ce point 3, est de permettre � n'importe qui d'utiliser la macro sans avoir � ouvrer l'�diteur VBA pour modifier des variables.
'   4 - Ouvrir l'explorateur Windows avec le chemin vers ces images pour v�rifier le job
'   5 - Mettre un petit gestionnaire d'erreur et msgbox de Succ�s ?

' Version 02 de la macro
' Date de l'�volution : 16/12/2016
' Evolutions :
'
'   1-  Factorisation du code :
'       Cr�ation d'une fonction "export_png" g�n�rique.
'       Cette fonction attend en argument, le chemin, le nom du fichier et le nom de l'onglet Visio dont l'image est � exporter.
'       Puis la fonction r�alise l'export.
'   2-  Modification des autres fonctions Export_d�di�es, afin d'int�grer la factorisation ajout�e.



Sub Export_Planning_PNG()
    
    'Definition of the variable
    Dim exportmyPath As String
    Dim exportmyName As String
    Dim exportmyPage As String

    'Parameter for the sub
    exportmyPath = "W:\Commun\Affaires\HREOS\Image\BancImageG�n�rique\2-Management\Compte rendu avancement & revues\11_Obeya\KPI-image-iObeya"
    exportmyName = "KPI-Planning.png"
    exportmyPage = "Export_KPI_Planning"
    
    'Call the generic export procedure with the right parameters.
    Export_PNG exportmyPage, exportmyName, exportmyPath

   

End Sub

Public Sub Export_PNG(strDesiredPage As String, strExportName As String, strPath As String)
'This is the generic function to export all the shapes of the given page, as PNG file.

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140

    Application.ActiveWindow.SelectAll
    
    'Resolution export setup
    Application.Settings.SetRasterExportResolution visRasterUseScreenResolution, 96#, 96#, visRasterPixelsPerInch
    Application.Settings.SetRasterExportSize visRasterFitToSourceSize, 19.479167, 7.614583, visRasterInch
    Application.Settings.RasterExportDataFormat = visRasterInterlace
    Application.Settings.RasterExportColorFormat = visRaster24Bit
    Application.Settings.RasterExportRotation = visRasterNoRotation
    Application.Settings.RasterExportFlip = visRasterNoFlip
    Application.Settings.RasterExportBackgroundColor = 16777215
    Application.Settings.RasterExportTransparencyColor = 16777215
    Application.Settings.RasterExportUseTransparencyColor = False
    ActiveWindow.DeselectAll
    
    
    Dim myPageColl As Pages
    Dim myPageItem As Page
    Dim myFlag As Boolean
    'Default and initial status of the flag
    myFlag = False
    'List all page of the activedocument
    Set myPageColl = Application.ActiveDocument.Pages
    For Each myPageItem In myPageColl
        Debug.Print myPageItem.Name
        'Test if the page exist
        If myPageItem.Name <> strDesiredPage Then
            Debug.Print "page ne correspond pas au param�tre pass� en appel de proc�dure"
        Else
            Debug.Print "page existe et correspond"
            myFlag = True
        End If
    Next myPageItem
    
    If myFlag = False Then
        MsgBox "Export non r�ussi car la page demand�e est inexistante - Erreur de param�trage de la macro VBA", vbCritical, "Erreur Macro VBA"
        Exit Sub
    End If

    'set the page for the export as the desired page
    Set myPageItem = myPageColl.Item(strDesiredPage)


    'change the page to be the desired page (i.e. mypageItem = "strDesiredPage")
    Application.ActiveWindow.Page = myPageItem


    'Select all the shape and export them
    Application.ActiveWindow.SelectAll
    Application.ActiveWindow.Selection.Export strPath & "\" & strExportName
    Debug.Print "PNG file export to : " & strPath & "\" & strExportName

    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
    Application.ActiveWindow.DeselectAll


End Sub
Sub Export_SonarJava_PNG()
    
    'Definition of the variable
    Dim exportmyPath As String
    Dim exportmyName As String
    Dim exportmyPage As String

    'Parameter for the sub
    exportmyPath = "W:\Commun\Affaires\HREOS\Image\BancImageG�n�rique\2-Management\Compte rendu avancement & revues\11_Obeya\KPI-image-iObeya"
    exportmyName = "KPI-SONAR-Java.png"
    exportmyPage = "SONAR"
    
    'Call the generic export procedure with the right parameters.
    Export_PNG exportmyPage, exportmyName, exportmyPath

End Sub



Sub Export_iObeyaFond_PNG()
    
    'Definition of the variable
    Dim exportmyPath As String
    Dim exportmyName As String
    Dim exportmyPage As String

    'Parameter for the sub
    exportmyPath = "W:\Commun\Affaires\HREOS\Image\BancImageG�n�rique\2-Management\Compte rendu avancement & revues\11_Obeya\KPI-image-iObeya"
    exportmyName = "iOBEYA-Fond.png"
    exportmyPage = "Fond1"
    
    'Call the generic export procedure with the right parameters.
    Export_PNG exportmyPage, exportmyName, exportmyPath
    

End Sub
