Imports pfcls

Class MainWindow

    Dim asyncConnection As IpfcAsyncConnection = Nothing
    Dim model As IpfcModel
    Dim solid As IpfcSolid
    Dim activeserver As IpfcServer
    Dim paramval As IpfcParamValue
    Dim session As IpfcBaseSession
    Dim Moditem As CMpfcModelItem
    Dim FileNameComplete As String
    Dim AmountInteger As Integer = 1

    Sub Creo_Connect()

        Dim asyncConnection As IpfcAsyncConnection = Nothing

        Try
            myInfo.Text = "Connecting..."

            asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
            session = asyncConnection.Session
            activeserver = session.GetActiveServer
            model = session.CurrentModel
            myInfo.Text = " "

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

    Sub TestForDir(workdir As String)
        Try
            If Dir(workdir, vbDirectory) = "" Then
                MkDir(workdir)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ExportFileToDisc()
        Dim Workdir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() & "\STL Files\"
        Dim CompleteFileName As String
        Dim TodaysDate As String
        Dim Amount As String
        Dim User As String
        Dim ModelName As String

        myInfo.Text = "*****"

        TestForDir(Workdir)

        If model Is Nothing Then
            MsgBox("Model is not present",, "Script message")
            asyncConnection.Disconnect(1)
            Environment.Exit(0)
        End If

        TodaysDate = Date.Now.ToShortDateString().Replace("-", "").Remove(0, 2)

        ModelName = model.FullName
        User = Environment.UserName
        Amount = myTextBox.Text

        If Amount = 1 Then
            CompleteFileName = TodaysDate + "_" + User + "_" + ModelName + ".stl"
        Else CompleteFileName = TodaysDate + "_" + User + "_" + ModelName + "_x" + Amount + ".stl"
        End If

        Dim Destination As String = Workdir + CompleteFileName


        Try
            Dim cDesExSTL As CCpfcSTLBinaryExportInstructions
            Dim DesEx As IpfcExportInstructions
            Dim DesExSTL As IpfcSTLBinaryExportInstructions
            Dim DesSTLEx As IpfcCoordSysExportInstructions


            cDesExSTL = New CCpfcSTLBinaryExportInstructions
            DesExSTL = cDesExSTL.Create("CSYS")
            DesSTLEx = DesExSTL

            DesSTLEx.Quality = 8
            DesSTLEx.AngleControl = Nothing
            DesSTLEx.StepSize = Nothing
            DesSTLEx.CsysName = "CSYS"
            DesSTLEx.MaxChordHeight = Nothing

            DesEx = DesSTLEx

            session.CurrentModel.Export(Destination, DesEx)
            Call UpdateWindowInfo(CompleteFileName + " exported")

        Catch ex As Exception

        End Try




    End Sub

    Private Sub ExportFileToDisc2()
        Dim Workdir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() & "\Fileshuffler Files\"
        Dim CompleteFileName As String
        Dim TodaysDate As String
        Dim Amount As String
        Dim User As String
        Dim ModelName As String

        myInfo.Text = "*****"

        TestForDir(Workdir)

        If model Is Nothing Then
            MsgBox("Model is not present",, "Script message")
            asyncConnection.Disconnect(1)
            Environment.Exit(0)
        End If

        TodaysDate = Date.Now.ToShortDateString().Replace("-", "").Remove(0, 2)

        ModelName = model.FullName
        User = Environment.UserName
        Amount = myTextBox.Text

        If (LTHCheckbox.IsChecked = True) Then
            CompleteFileName = ModelName + "_LTH_" + User + "_" + TodaysDate + "_x" + Amount + ".stl"
        Else CompleteFileName = ModelName + "_Axis_" + User + "_" + TodaysDate + "_x" + Amount + ".stl"
        End If


        Dim Destination As String = Workdir + CompleteFileName


        Try
            Dim cDesExSTL As CCpfcSTLBinaryExportInstructions
            Dim DesEx As IpfcExportInstructions
            Dim STLInstructions As IpfcSTLBinaryExportInstructions
            Dim DesSTLEx As IpfcCoordSysExportInstructions
            cDesExSTL = New CCpfcSTLBinaryExportInstructions

            STLInstructions = cDesExSTL.Create("CSYS")

            DesSTLEx = STLInstructions
            DesSTLEx.AngleControl = 0
            DesSTLEx.MaxChordHeight = 0.01
            DesSTLEx.StepSize = 0
            DesSTLEx.Quality = Nothing

            DesEx = DesSTLEx

            session.CurrentModel.Export(Destination, DesEx)
            Call UpdateWindowInfo(CompleteFileName + " exported")

        Catch ex As Exception

        End Try

    End Sub


    Private Sub UpdateWindowInfo(v As String)
        MessageBox.Show(v)
    End Sub

    Private Sub MyButton_Click(sender As Object, e As RoutedEventArgs) Handles myButton.Click
        Try
            ExportFileToDisc2()

            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try
        Close()
    End Sub

    Private Sub myWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles myWindow.Closing
        Try
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try

    End Sub

End Class

