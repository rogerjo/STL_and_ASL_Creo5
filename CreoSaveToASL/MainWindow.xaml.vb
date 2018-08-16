Imports pfcls

Class MainWindow

    Dim asyncConnection As IpfcAsyncConnection = Nothing
    Dim model As IpfcModel
    Dim solid As IpfcSolid
    Dim activeserver As IpfcServer
    Dim paramval As IpfcParamValue
    Dim session As IpfcBaseSession

    Sub Creo_Connect()

        Dim asyncConnection As IpfcAsyncConnection = Nothing

        Try
            myInfo.Text = "Connecting..."
            asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
            session = asyncConnection.Session
            activeserver = session.GetActiveServer
            model = session.CurrentModel
            myInfo.Text = " "
            Call SetWhiteBackMacro()

        Catch ex As Exception
            MsgBox(ex.Message.ToString + Chr(13) + ex.StackTrace.ToString)
            If Not asyncConnection Is Nothing AndAlso asyncConnection.IsRunning Then
                asyncConnection.Disconnect(1)
            End If
            myInfo.Text = "Error occurred while connecting"
        End Try
    End Sub

    Private Sub MyWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles myWindow.Loaded
        myInfo.Text = ""

        Call Creo_Connect()

    End Sub

    Private Sub UpdateWindowInfo(v As String)
        MessageBox.Show(v)
    End Sub

    Private Sub MyWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles myWindow.Closing
        Try
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MyButton_Click(sender As Object, e As RoutedEventArgs) Handles myButton.Click
        Try
            ExportImageToASL()
            'ExportImageCreo5()
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try

        Close()

    End Sub

    Private Sub ExportImageToASL()
        Dim ASLDir As String = "\\galaxis.axis.com\DavWWWRoot\portfolios\NewVideoProducts\PublishingImages\"
        'Dim ASLDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString() + "\"
        Dim CompleteFilename As String
        Dim ModelName As String
        Dim CurrWindow As IpfcWindow
        myInfo.Text = "*****"

        If model Is Nothing Then
            MsgBox("Model is not present",, "Script message")
            asyncConnection.Disconnect(1)
            Environment.Exit(0)
        End If

        ModelName = model.FullName
        CompleteFilename = ASLDir + ModelName
        CurrWindow = session.GetModelWindow(model)

        Try
            'Call SetWhiteBackMacro()
            Call OutputImageWindow(CurrWindow, EpfcRasterType.EpfcRASTER_JPEG, CompleteFilename)

        Catch ex As Exception

        End Try



    End Sub

    Private Sub SetWhiteBackMacro()

        Try

            Dim macrostring As String
            macrostring = "~ Trail `UI Desktop` `UI Desktop` `PREVIEW_POPUP_TIMER` \"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "`main_dlg_w1:PHTLeft.AssyTree:<NULL>`;~ Select `main_dlg_cur` `appl_casc`;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Close `main_dlg_cur` `appl_casc`;~ Command `ProCmdRibbonOptionsDlg` ;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Select `ribbon_options_dialog` `PageSwitcherPageList` 1 `colors_layouts`;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Open `ribbon_options_dialog` `colors_layouts.Color_scheme_optMenu`;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Close `ribbon_options_dialog` `colors_layouts.Color_scheme_optMenu`;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Select `ribbon_options_dialog` `colors_layouts.Color_scheme_optMenu` 1 `2`;\"
            macrostring = macrostring & vbCrLf
            macrostring = macrostring & "~ Activate `ribbon_options_dialog` `OkPshBtn`;"


            session.RunMacro(macrostring)

        Catch ex As Exception

        End Try

    End Sub
    Private Function GetRasterInstructions(ByVal type As Integer,
                                       ByVal rasterWidth As Double,
                                       ByVal rasterHeight As Double,
                                       ByVal dotsPerInch As Integer,
                                       ByVal imageDepth As Integer) As _
                                       IpfcRasterImageExportInstructions

        Dim instructions As IpfcRasterImageExportInstructions

        Select Case type

            Case EpfcRasterType.EpfcRASTER_BMP
                Dim bmpInstrs As IpfcBitmapImageExportInstructions
                bmpInstrs = (New CCpfcBitmapImageExportInstructions).Create(rasterWidth, rasterHeight)
                instructions = bmpInstrs

            Case EpfcRasterType.EpfcRASTER_TIFF
                Dim tiffInstrs As IpfcTIFFImageExportInstructions
                tiffInstrs = (New CCpfcTIFFImageExportInstructions).Create(rasterWidth, rasterHeight)
                instructions = tiffInstrs

            Case EpfcRasterType.EpfcRASTER_JPEG
                Dim jpegInstrs As IpfcJPEGImageExportInstructions
                jpegInstrs = (New CCpfcJPEGImageExportInstructions).Create(rasterWidth, rasterHeight)
                instructions = jpegInstrs

            Case EpfcRasterType.EpfcRASTER_EPS
                Dim epsInstrs As IpfcEPSImageExportInstructions
                epsInstrs = (New CCpfcEPSImageExportInstructions).Create(rasterWidth, rasterHeight)
                instructions = epsInstrs

            Case Else
                Throw New Exception("Unsupported Raster Type")
        End Select

        instructions.DotsPerInch = dotsPerInch
        instructions.ImageDepth = imageDepth

        Return instructions
    End Function

    Public Sub OutputImageWindow(ByRef window As IpfcWindow,
                                 ByVal type As Integer,
                                 ByVal imageName As String)
        Dim instructions As IpfcRasterImageExportInstructions
        Dim imageExtension As String
        Dim rasterHeight As Double = 7.5
        Dim rasterWidth As Double = 10.0
        Dim dotsPerInch As Integer
        Dim imageDepth As Integer

        Try
            dotsPerInch = EpfcDotsPerInch.EpfcRASTERDPI_100
            imageDepth = EpfcRasterDepth.EpfcRASTERDEPTH_24

            instructions = GetRasterInstructions(type, rasterWidth,
                                                 rasterHeight, dotsPerInch,
                                                 imageDepth)

            imageExtension = GetRasterExtension(type)

            window.ExportRasterImage(imageName + imageExtension, instructions)

        Catch ex As Exception
            MsgBox(ex.Message.ToString + Chr(13) + ex.StackTrace.ToString)
        End Try
    End Sub

    Private Function GetRasterExtension(ByVal type As Integer) As String

        Select Case type

            Case EpfcRasterType.EpfcRASTER_BMP
                Return ".bmp"

            Case EpfcRasterType.EpfcRASTER_TIFF
                Return ".tiff"

            Case EpfcRasterType.EpfcRASTER_JPEG
                'Return ".jpg"
                Return ".png"

            Case EpfcRasterType.EpfcRASTER_EPS
                Return ".eps"

            Case Else
                Throw New Exception("Unsupported Raster Type")
        End Select

    End Function
End Class
