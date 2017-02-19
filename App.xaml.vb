''' <summary>
''' Provides application-specific behavior to supplement the default Application class.
''' </summary>
NotInheritable Class App
    Inherits Application

    ''' <summary>
    ''' Initializes a new instance of the App class.
    ''' </summary>
    Public Sub New()
        Microsoft.ApplicationInsights.WindowsAppInitializer.InitializeAsync(
            Microsoft.ApplicationInsights.WindowsCollectors.Metadata Or
            Microsoft.ApplicationInsights.WindowsCollectors.Session)
        InitializeComponent()
    End Sub

    ''' <summary>
    ''' Invoked when the application is launched normally by the end user.  Other entry points
    ''' will be used when the application is launched to open a specific file, to display
    ''' search results, and so forth.
    ''' </summary>
    ''' <param name="e">Details about the launch request and process.</param>
    Protected Overrides Sub OnLaunched(e As Windows.ApplicationModel.Activation.LaunchActivatedEventArgs)
#If DEBUG Then
        ' Show graphics profiling information while debugging.
        If System.Diagnostics.Debugger.IsAttached Then
            ' Display the current frame rate counters
            Me.DebugSettings.EnableFrameRateCounter = True
        End If
#End If

        Dim rootFrame As Frame = TryCast(Window.Current.Content, Frame)

        ' Do not repeat app initialization when the Window already has content,
        ' just ensure that the window is active

        If rootFrame Is Nothing Then
            ' Create a Frame to act as the navigation context and navigate to the first page
            rootFrame = New Frame()

            AddHandler rootFrame.NavigationFailed, AddressOf OnNavigationFailed

            If e.PreviousExecutionState = ApplicationExecutionState.Terminated Then
                ' TODO: Load state from previously suspended application
            End If
            ' Place the frame in the current Window
            Window.Current.Content = rootFrame
        End If

        If e.PrelaunchActivated = False Then
            If rootFrame.Content Is Nothing Then
                ' When the navigation stack isn't restored navigate to the first page,
                ' configuring the new page by passing required information as a navigation
                ' parameter
                rootFrame.Navigate(GetType(MainPage), e.Arguments)
            End If

            ' Ensure the current window is active
            Window.Current.Activate()
        End If
    End Sub

    ''' <summary>
    ''' Invoked when Navigation to a certain page fails
    ''' </summary>
    ''' <param name="sender">The Frame which failed navigation</param>
    ''' <param name="e">Details about the navigation failure</param>
    Private Sub OnNavigationFailed(sender As Object, e As NavigationFailedEventArgs)
        Throw New Exception("Failed to load Page " + e.SourcePageType.FullName)
    End Sub

    ''' <summary>
    ''' Invoked when application execution is being suspended.  Application state is saved
    ''' without knowing whether the application will be terminated or resumed with the contents
    ''' of memory still intact.
    ''' </summary>
    ''' <param name="sender">The source of the suspend request.</param>
    ''' <param name="e">Details about the suspend request.</param>
    Private Sub OnSuspending(sender As Object, e As SuspendingEventArgs) Handles Me.Suspending
        Dim deferral As SuspendingDeferral = e.SuspendingOperation.GetDeferral()
        ' TODO: Save application state and stop any background activity
        deferral.Complete()
    End Sub

End Class

' Proof spelling here: https://github.com/jaylagorio/CareLink-USB-for-Windows-Apps

'For the next release
' Band integration                                              '89ee4f44-39d0-4d7b-93da-553a4305efad
'   First page: Arrow, Number (delta), Last Updated (x mins)    'b42218b5-a1a2-4b97-a5d0-2e538d2f1a92
'   Second page: IOB, Current Basal, Last Loop run (x mins)     'bd810550-75c2-4ab2-b00b-eff65b99e4c6
'   Third page: Last Treatment size, carbs if so, time delta    '031d1ee0-00a6-4903-83bd-bcfca64abb66
'   Fourth page: Refresh button                                 '01229f41-8c31-425e-9a8a-bd84108eb81a
' Lock Screen: 
'   https://msdn.microsoft.com/en-us/library/windows/apps/dn934800.aspx
'   https://msdn.microsoft.com/en-us/library/windows/apps/hh779720.aspx
'   https://msdn.microsoft.com/en-us/library/windows/apps/dn934782.aspx
' App insights? How do I get more info from this and add more telemetry
' https://github.com/openaps/dexcom_reader/pull/6, https://github.com/bewest/dexcom_reader/pull/1, https://github.com/bewest/dexcom_reader/pull/1#issuecomment-157866179
' Care Portal shows a dialog when entering items but that dialog doesn't show in the app - https://social.msdn.microsoft.com/Forums/en-US/2a075255-39f6-4b1c-97b1-a62e7c566633/how-can-i-handle-confirm-dialog-in-webview-uwp-windows-10-app-c?forum=wpdevelop; http://stackoverflow.com/questions/35405827/uwp-webview-calls-dynamic-javascript (modify the function not to need the confirm);
