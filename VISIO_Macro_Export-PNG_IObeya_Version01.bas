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





Sub Export_Planning_PNG()

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140

    Application.ActiveWindow.SelectAll

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
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(1), visSelect
    '
    Application.ActiveWindow.SelectAll
    Application.ActiveWindow.Selection.Export "W:\Commun\Affaires\HREOS\Image\BancImageG�n�rique\2-Management\Compte rendu avancement & revues\11_Obeya\KPI-image-iObeya\KPI-Planning.png"

    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
    Application.ActiveWindow.DeselectAll

End Sub
Sub Export_SonarJava_PNG()
    
    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140

    Application.ActiveWindow.SelectAll

    Application.Settings.SetRasterExportResolution visRasterUseScreenResolution, 96#, 96#, visRasterPixelsPerInch
    Application.Settings.SetRasterExportSize visRasterFitToSourceSize, 12.4375, 7.322917, visRasterInch
    Application.Settings.RasterExportDataFormat = visRasterInterlace
    Application.Settings.RasterExportColorFormat = visRaster24Bit
    Application.Settings.RasterExportRotation = visRasterNoRotation
    Application.Settings.RasterExportFlip = visRasterNoFlip
    Application.Settings.RasterExportBackgroundColor = 16777215
    Application.Settings.RasterExportTransparencyColor = 16777215
    Application.Settings.RasterExportUseTransparencyColor = False
    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(1), visSelect
    '
    Application.ActiveWindow.SelectAll
    Application.ActiveWindow.Selection.Export "W:\Commun\Affaires\HREOS\Image\BancImageG�n�rique\2-Management\Compte rendu avancement & revues\11_Obeya\KPI-image-iObeya\KPI-SONAR-Java.png"

    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
    Application.ActiveWindow.DeselectAll

End Sub
