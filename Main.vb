''Import Notes
'Import Management for WMI query to find com port device names
'Import Threading to run reading of hypot register from seperate thread

Imports System.Management
Imports System.Threading
Imports System.IO


Public Class Main

    ''Variable Notes
    'A sperate thread is required to query the hypot tester because this tread is a continious loop. The main thread 
    'of the application is used to handle the windows form events. Using this same thread to send rapid update requests
    'to the hypot tester will cause the windows forms to freeze
    '
    'WMI is the Windows Managment Instrumentation. It allows programatic access of computer managmenet information. This
    'lets us scan all the com port devices and pull thier names. We looks for the USB serial device to automatically
    'detect to and connect to our USB-Serial converter and into the Hypot tester
    ''

    Dim comPortName As String                               ''Name of com port to connect to, set dynamically during form load
    Dim com1 As IO.Ports.SerialPort = Nothing               ''Declare and instantiate com port
    Dim hypotUpdate As New Thread(AddressOf readRegister)   ''Declare and instantiate thread to run fucntion 'readRegister'
    Dim devices As ArrayList = New ArrayList                ''Arraylsit to hold names of devices found by WMI query
    Dim tryConnect As Boolean = True

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''WMI uses an SQL query to select devices which have the specified GUID id. The id noted below returns all the com ports
        'present in the system. The Win32_PnPEntity represents all the attached hardware devices to the system.
        'The Chr(34) is used to substitute for a quotation mark (") as is required by an sql statement when defining a field. If
        'we used quotes in the actual string definition it would cause a conflict with the complier due to how it defines strings
        'The variable sql will hold our wmi sql query
        Dim readStr As String = ""
        Dim strIndex As Integer
        Dim sql As String
        Dim tempComPort As String
        sql = "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=" & Chr(34) & "{4d36e978-e325-11ce-bfc1-08002be10318}" & Chr(34)

        'writeToErrorLog("<Application Started>")

        'Try cases let you attempt to run a line of code which can cause a runtime error. A try-catch statement will try a line
        'of code and attempt to catch any errors before a runtime error can cause a program crash. We can then handle the error safetly
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", sql)     ''Define a searcher using our SQL query

            For Each queryObj As ManagementObject In searcher.Get()             ''Loop through each object found by the searcher
                devices.Add(queryObj("Name"))                                   ''Add the name of the found device to an array list
            Next
            'writeToErrorLog("WMI query successful")
        Catch err As Exception                                        ''Cathc a management excepetion or runtime error
            MessageBox.Show("An error occurred while querying for WMI data: " & err.Message & " closing application")
            writeToErrorLog("An error occurred while querying for WMI data: " & err.Message & " closing application")
            Me.Close()
        End Try

        ''Pulling data from the WMI query
        'The WMI query will return a port name string as so: "USB Serial Device (COM8)
        'We need to seperate the com port address and the name because while the name of the device will not change, the com port will
        'We need to ensure that the accurate com port address number (COM1, COM2... etc) will be provided to the connect function
        comPortName = "none"                                                            ''Set default name if no device is found
        For i As Integer = 0 To devices.Count - 1                                       ''Loop through the array list of found device names
            If getComDescription(devices(i)) = "Prolific USB-to-Serial Comm Port" Then  ''Use string function to get the name and compare
                'MessageBox.Show("USB Device Found: Click ok to Connect")                ''If the USB Serial Port is found display a msg to user
                tempComPort = getComPort(devices(i))                                    ''Extract and save the comport name to a variable
                connect(tempComPort)                                                    ''Call the function to connect to the found com port
                If tryconnect Then
                    If com1.IsOpen Then
                        Try
                            com1.WriteLine("*IDN?")             ''Write "*IDN?" command to hypot to request name
                            'writeToErrorLog("IDN write success to device " & i)
                        Catch ex As Exception                   ''These exceptions are tried to make sure port may not be randomly disconnected during testing
                            MsgBox("Error: " & ex.Message & " closing application")
                            writeToErrorLog("Error writing IDN string request on startup device: " & i & " Error: " & ex.Message & " Closing application")
                            Me.Close()
                        End Try
                        System.Threading.Thread.Sleep(200)  ''Delay to allow hypot to process command
                        Try
                            readStr = com1.ReadLine        ''Read response from Hypot
                            strIndex = InStr(readStr, ",")
                            readStr = Mid(readStr, 1, strIndex - 1)
                            'writeToErrorLog("Response recieved: " & readStr)
                            If readStr = "ASSOCIATED RESEARCH" Or readStr = "ASSOCIATED RESEARCH INC." Then 'Possible names of Hypot tester
                                'writeToErrorLog("Hypot tester found at " & comPortName)
                                MsgBox("Hypot tester found, click OK to connect")
                                comPortName = getComPort(devices(i))
                            End If
                        Catch ex As Exception
                            'writeToErrorLog("Error: Serial Port read timed out, checking next device")
                            'Me.Close()
                        End Try
                        com1.Close()
                    End If
                Else
                    tryConnect = True
                End If
            End If
        Next

        If comPortName = "none" Then                                          ''If the com port has not been set aka no matching devices foun then display msg and exit program
            MsgBox("Could not find Hypot, application will exit")
            writeToErrorLog("Hypot tester not found on startup, closing application")
            Me.Close()
        Else
            Me.CheckForIllegalCrossThreadCalls = False                        ''This function disables warnings for multithread access issues, not the best solution but the program is too simple for a more complex work around
            connect(comPortName)
            If com1.IsOpen Then
                'writeToErrorLog("Successfully connected to Hypot tester")
                MsgBox("Successfully connected to Hypot tester")
                com1.ReadTimeout = 2000
                hypotUpdate.Start()                                               ''Once the connect function finishes, start querying the hypot tester
            Else
                MsgBox("Could not connect, application will exit")
                writeToErrorLog("Could not connect to Hypot tester, closing application")
                Me.Close()
            End If

        End If
    End Sub


    Private Sub connect(portName As String)
        Try
            com1 = My.Computer.Ports.OpenSerialPort(portName)       ''Try to open serial port at specified com port name 
            'writeToErrorLog("Connected to COM port at " & portName)
        Catch ex As Exception
            If ex.GetType.ToString = "System.UnauthorizedAccessException" Then
                'writeToErrorLog("UnauthorizedAccessException on COM port " & portName & " connect, skipping device")
                tryConnect = False
            Else
                MsgBox("Exception:" & vbCrLf & ex.Message & " closing application")              ''Catch any exception that might occur
                writeToErrorLog("Error connecting to COM port " & portName & " Error: " & ex.Message & " Closing application")
                Me.Close()                                              ''Close the form on error
            End If
        End Try
    End Sub

    Private Sub readRegister()
        Dim readLine As String = ""             ''Incoming data read from Hypot
        Dim readIndex As Integer                ''Dynamically set based on number of allowed results
        Dim strIndex As Integer                 ''Index used to locate ' using string functions to pull pass/fail state from comma delimited hypot return data
        Dim readytoRead As Boolean              ''Is set when the state of the hypot changes from in process to finished

        'writeToErrorLog("-Starting Hypot update thread loop-")
        Do While (com1 IsNot Nothing)           ''Runs the update loop while the com port is open
            Try
                com1.WriteLine("*STB?")             ''Write "*STB?" command to hypot to request status register
            Catch ex As Exception                   ''These exceptions are tried to make sure port may not be randomly disconnected during testing
                MsgBox("Error: " & ex.Message & " closing application")
                writeToErrorLog("Error writing status request to Hypot: " & ex.Message & " closing application")
                Me.Close()
            End Try
            System.Threading.Thread.Sleep(200)  ''Delay to allow hypot to process command
            Try
                readLine = com1.ReadLine        ''Read response from Hypot
            Catch ex As Exception
                MsgBox("Error: Serial Port read timed out, closing application")
                writeToErrorLog("Error reading status response from Hypot: " & ex.Message & " closing application")
                Me.Close()
            End Try
            If readLine = "40" Or readLine = "8" Then             ''40 or 8 is returned when hypot is in process depending on model
                readytoRead = True              ''Set readyToRead when hypot is in process
                lstOutput.Items.Clear()         ''Clear list box of previous results
            End If

            If readLine <> "40" And readLine <> "8" And readytoRead = True Then ''When the hypot registers shows it is finished and coming out of an in process step, read the results
                readytoRead = False                         ''Reset ready to read bit
                com1.WriteLine("TD?")                       ''Query hypot about last step in test to determined number of steps (last step = number of steps)
                Try
                    readLine = com1.ReadLine                ''Read response
                Catch ex As Exception
                    MsgBox("Error: Serial Port read timed out.")
                    writeToErrorLog("Error reading last test result from Hypot: " & ex.Message & " closing application")
                    Me.Close()
                End Try
                strIndex = InStr(readLine, ",")             ''Data is comma delimited, step number is before first comma. Get index location of first comma
                readLine = Mid(readLine, 1, strIndex - 1)              ''Pull string out from start of original string to index of first comma, this returns the last step number
                readIndex = Convert.ToInt32(readLine)       ''Convert that string step number into an integer format to be used in a for loop for enumaration
                For i As Integer = 1 To readIndex           ''For loop to run as many times as the returned number of steps
                    com1.WriteLine("RD " & i & "?")         ''The for loop index 'i' is inserted into the read step number command for the hypot
                    Try
                        readLine = com1.ReadLine
                    Catch ex As Exception
                        MsgBox("Error: Serial Port read timed out, closing application")
                        writeToErrorLog("Error reading test result " & i & " from Hypot: " & ex.Message & " closing application")
                        Me.Close()
                    End Try
                    For j As Integer = 1 To 2                   ''This for loop runs twice to drop off sections of comma delimited data from the hypot return
                        strIndex = InStr(readLine, ",")
                        readLine = Mid(readLine, strIndex + 1, readLine.Length)
                    Next
                    strIndex = InStr(readLine, ",")             ''The final string function returns of the pass/fail state pulled from the comma delimited data
                    If strIndex - 1 < 1 Then
                        strIndex = 2
                    End If
                    readLine = Mid(readLine, 1, strIndex - 1)
                    lstOutput.Items.Add(i & ": " & readLine)    ''Add the results into the listbox
                    If readLine = "Pass" Then                  ''Depenging on the pass/fail state of results change color to match 
                        lstOutput.Items.Item(i - 1).ForeColor = Color.Green
                    Else
                        If readLine = "Abort" Then
                            lstOutput.Items.Item(i - 1).ForeColor = Color.Black
                        Else
                            lstOutput.Items.Item(i - 1).ForeColor = Color.Red
                        End If
                    End If
                Next                                            ''Increment the for loop
            End If
            System.Threading.Thread.Sleep(200)                  ''Thread dealy to allow hypot time to respond, slow down requests
        Loop
    End Sub


    ''String functions to pull the name of the com port from the return of the WMI query
    ''A little bit of trail and error to get the indexing correct. 
    ''Both Functions look for ( to determine the location of the description and the actual portname in the WMI query return string
    Function getComPort(inputString As String) As String
        Dim indexStart As Integer
        Dim indexEnd As Integer
        Dim returnString As String
        indexStart = InStr(inputString, "(")
        indexEnd = InStr(inputString, ")")
        returnString = Mid(inputString, indexStart + 1, indexEnd - (indexStart + 1))

        Return returnString             ''These mehtods are return methods, when they called they return a string
    End Function

    ''String functions to pull the description of the com port from the return of the WMI query
    Function getComDescription(inputString As String) As String
        Dim index As String
        Dim returnString As String
        index = InStr(inputString, "(")
        returnString = Mid(inputString, 1, index - 2)

        Return returnString
    End Function

    ''Function that handles closing of the application. The com port and Hypot update thread must be stopped before the application to avoid an error
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If hypotUpdate.IsAlive Then     ''If the hypot update thread is still running then abort it
            hypotUpdate.Abort()
            writeToErrorLog("Running Hypot update thread stopped")
        Else
            writeToErrorLog("No running Hypot thread")
        End If

        If com1 IsNot Nothing Then      ''If the comport has been set aka is not nothing (lol) then close the comport
            com1.Close()
            writeToErrorLog("Open COM port closed at " & comPortName)
        Else
            writeToErrorLog("No open COM port")
        End If
        writeToErrorLog("Application Ended Successfully")

        ''As soon as this function ends the application will close
        ''Me.Close is used throughout the program to close the application in the event of any errors
        ''Anytime Me.Close is called it will run this function first before closing 
    End Sub

    Private Sub writeToErrorLog(ByVal msg As String)

        ''Check and make the directory if necessary; this is set to look in the Application folder
        If Not System.IO.Directory.Exists(Application.StartupPath & "\ErrorLog\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\ErrorLog\")
        End If

        ''Check the file
        Dim fs As FileStream = New FileStream(Application.StartupPath &
        "\ErrorLog\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        ''Log it
        Dim fs1 As FileStream = New FileStream(Application.StartupPath & "\ErrorLog\errlog.txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.Write(DateTime.Now.ToString() & ": ")
        s1.Write(msg & vbCrLf)
        s1.Close()
        fs1.Close()

    End Sub
End Class
