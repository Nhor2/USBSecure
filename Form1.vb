Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Management
Imports System.Text
Imports System.Net.Mime.MediaTypeNames
Imports System.Security
Imports System.Net
Imports System.Security.Cryptography

Public Class Form1

    'USBSecure
    'un modo originale di utilizzare la pendrive usb.
    '13-04-2024
    'https://bytes.com/topic/visual-basic/answers/822709-get-usb-serial-number

    'il numero di serie usb/hdd scelto
    Dim serialusb_secure As String = ""


    Public Const IOCTL_STORAGE_GET_DEVICE_NUMBER As Integer = &H2D1080
    Public Const IOCTL_STORAGE_QUERY_PROPERTY As Integer = &H2D1400

    <DllImport("Kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSECURITY_ATTRIBUTES As IntPtr, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As IntPtr) As IntPtr
    End Function

    <DllImport("Kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function CloseHandle(hObject As IntPtr) As Boolean
    End Function

    Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Object, ByVal nInBufferSize As Long, lpOutBuffer As Object, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Object) As Long


    Public Const FILE_SHARE_READ As Integer = 1
    Public Const FILE_SHARE_WRITE As Integer = 2
    Public Const FILE_SHARE_DELETE As Integer = 4
    Public Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
    Public Const GENERIC_READ As Integer = &H80000000
    Public Const CREATE_NEW As Integer = 1
    Public Const CREATE_ALWAYS As Integer = 2
    Public Const OPEN_EXISTING As Integer = 3
    Public Const OPEN_ALWAYS As Integer = 4
    Public Const TRUNCATE_EXISTING As Integer = 5

    <DllImport("Kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function DeviceIoControl(hDevice As IntPtr,
                                            dwIoControlCode As Integer,
                                            lpInBuffer As IntPtr,
                                            nInBufferSize As Integer,
                                            lpOutBuffer As IntPtr,
                                            nOutBufferSize As Integer,
                                            ByRef lpBytesReturned As Integer,
                                            ByRef lpOverlapped As System.Threading.NativeOverlapped) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure STORAGE_DEVICE_DESCRIPTOR
        Public Version As UInteger
        Public Size As UInteger
        Public DeviceType As Byte
        Public DeviceTypeModifier As Byte
        <MarshalAs(UnmanagedType.I1)>
        Public RemovableMedia As Boolean
        <MarshalAs(UnmanagedType.I1)>
        Public CommandQueueing As Boolean
        Public VendorIdOffset As UInteger
        Public ProductIdOffset As UInteger
        Public ProductRevisionOffset As UInteger
        Public SerialNumberOffset As UInteger
        Public BusType As STORAGE_BUS_TYPE
        Public RawPropertiesLength As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
        Public RawDeviceProperties As Byte()
    End Structure

    Public Enum STORAGE_BUS_TYPE
        BusTypeUnknown = &H0
        BusTypeScsi = &H1
        BusTypeAtapi = &H2
        BusTypeAta = &H3
        BusType1394 = &H4
        BusTypeSsa = &H5
        BusTypeFibre = &H6
        BusTypeUsb = &H7
        BusTypeRAID = &H8
        BusTypeiScsi = &H9
        BusTypeSas = &HA
        BusTypeSata = &HB
        BusTypeSd = &HC
        BusTypeMmc = &HD
        BusTypeVirtual = &HE
        BusTypeFileBackedVirtual = &HF
        BusTypeSpaces = &H10
        BusTypeNvme = &H11
        BusTypeSCM = &H12
        BusTypeMax = &H13
        BusTypeMaxReserved = &H7F
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure STORAGE_PROPERTY_QUERY
        Public PropertyId As STORAGE_PROPERTY_ID
        Public QueryType As STORAGE_QUERY_TYPE
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
        Public AdditionalParameters As Byte()
    End Structure

    Public Enum STORAGE_PROPERTY_ID
        StorageDeviceProperty = 0
        StorageAdapterProperty = 1
        StorageDeviceIdProperty = 2
        StorageDeviceUniqueIdProperty = 3
        StorageDeviceWriteCacheProperty = 4
        StorageMiniportProperty = 5
        StorageAccessAlignmentProperty = 6
        StorageDeviceSeekPenaltyProperty = 7
        StorageDeviceTrimProperty = 8
        StorageDeviceWriteAggregationProperty = 9
        StorageDeviceDeviceTelemetryProperty = 10
        StorageDeviceLBProvisioningProperty = 11
        StorageDevicePowerProperty = 12
        StorageDeviceCopyOffloadProperty = 13
        StorageDeviceResiliencyProperty = 14
        StorageDeviceMediumProductType = 15
        StorageAdapterRpmbProperty = 16
        StorageDeviceIoCapabilityProperty = 48
        StorageAdapterProtocolSpecificProperty = 49
        StorageDeviceProtocolSpecificProperty = 50
        StorageAdapterTemperatureProperty = 51
        StorageDeviceTemperatureProperty = 52
        StorageAdapterPhysicalTopologyProperty = 53
        StorageDevicePhysicalTopologyProperty = 54
        StorageDeviceAttributesProperty = 55
        StorageDeviceManagementStatus = 56
        StorageAdapterSerialNumberProperty = 57
        StorageDeviceLocationProperty = 58
    End Enum

    Public Enum STORAGE_QUERY_TYPE
        PropertyStandardQuery = 0 ' Retrieves the descriptor                                 
        PropertyExistsQuery = 1 ' Used To test whether the descriptor Is supported           
        PropertyMaskQuery = 2 ' Used To retrieve a mask Of writeable fields In the descriptor
        PropertyQueryMaxDefined = 3 ' use To validate the value                              
    End Enum

    Private Structure MEDIA_SERIAL_NUMBER_DATA
        Dim SerialNumberLength As Long
        Dim Result As Long
        Dim Reserved() As Long
        Dim SerialNumberData() As Byte
    End Structure

    Private Const INVALID_HANDLE_VALUE = -1
    Private Const FILE_FLAG_DELETE_ON_CLOSE = 67108864
    Private Const GENERIC_WRITE = &H40000000
    Private Const IOCTL_STORAGE_GET_MEDIA_SERIAL_NUMBER As Long = &H2D0C10

    Private Function GetPhysicalDriveSerial(PhysicalDrive As String) As String
        Dim hDrive As Long
        Dim DummyReturnedBytes As Long
        Dim si As MEDIA_SERIAL_NUMBER_DATA = Nothing

        hDrive = CreateFile(PhysicalDrive, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
        If hDrive <> INVALID_HANDLE_VALUE Then
            Call DeviceIoControl(hDrive, IOCTL_STORAGE_GET_MEDIA_SERIAL_NUMBER, 0, 0, si, Len(si), DummyReturnedBytes, 0)
            Call CloseHandle(hDrive)
            Return si.SerialNumberData.ToString
        End If
        Return Nothing
    End Function


    Private Function GetSerialNumberDisk(drive As String) As String
        Dim nBytesReturned As Integer = 0
        Dim hVolume As IntPtr = IntPtr.Zero
        Dim sDriveName As String = String.Format("\\.\{0}", drive)
        hVolume = CreateFile(sDriveName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
        If (hVolume <> -1) Then
            Dim DeviceDescriptor As STORAGE_DEVICE_DESCRIPTOR = New STORAGE_DEVICE_DESCRIPTOR()
            Dim PropertyQuery As STORAGE_PROPERTY_QUERY = New STORAGE_PROPERTY_QUERY()

            PropertyQuery.PropertyId = STORAGE_PROPERTY_ID.StorageDeviceProperty
            PropertyQuery.QueryType = STORAGE_QUERY_TYPE.PropertyStandardQuery
            Dim nBytesPropertyQuery As Integer = Marshal.SizeOf(PropertyQuery)
            Dim pPropertyQuery As IntPtr = Marshal.AllocHGlobal(nBytesPropertyQuery)
            Marshal.StructureToPtr(PropertyQuery, pPropertyQuery, False)

            Dim nBytesDeviceDescriptor As Integer = Marshal.SizeOf(DeviceDescriptor) + 2048
            Dim pDeviceDescriptor As IntPtr = Marshal.AllocHGlobal(nBytesDeviceDescriptor)

            Dim bDeviceIoControl As Boolean = DeviceIoControl(hVolume, IOCTL_STORAGE_QUERY_PROPERTY, pPropertyQuery, Marshal.SizeOf(GetType(STORAGE_PROPERTY_QUERY)), pDeviceDescriptor, nBytesDeviceDescriptor, nBytesReturned, Nothing)
            Dim DeviceDescriptorRet As STORAGE_DEVICE_DESCRIPTOR = Marshal.PtrToStructure(pDeviceDescriptor, GetType(STORAGE_DEVICE_DESCRIPTOR))

            Dim nSerialNumberOffset As UInteger = DeviceDescriptorRet.SerialNumberOffset
            Dim pSerialNumber As IntPtr = IntPtr.Add(pDeviceDescriptor, CInt(nSerialNumberOffset))
            Dim sSerialNumber As String = Marshal.PtrToStringAnsi(pSerialNumber)

            Console.WriteLine("SerialNumber : {0}", sSerialNumber)

            Marshal.FreeHGlobal(pDeviceDescriptor)
            Marshal.FreeHGlobal(pPropertyQuery)
            CloseHandle(hVolume)

            Return sSerialNumber
        Else
            Return Nothing
        End If
    End Function

    Private Function GetSerialNumberDisk2(PhysicalDrive As String) As String
        Dim nBytesReturned As Integer = 0
        Dim hVolume As IntPtr = IntPtr.Zero
        Dim sDriveName As String = PhysicalDrive
        hVolume = CreateFile(sDriveName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
        If (hVolume <> -1) Then
            Dim DeviceDescriptor As STORAGE_DEVICE_DESCRIPTOR = New STORAGE_DEVICE_DESCRIPTOR()
            Dim PropertyQuery As STORAGE_PROPERTY_QUERY = New STORAGE_PROPERTY_QUERY()

            PropertyQuery.PropertyId = STORAGE_PROPERTY_ID.StorageDeviceProperty
            PropertyQuery.QueryType = STORAGE_QUERY_TYPE.PropertyStandardQuery
            Dim nBytesPropertyQuery As Integer = Marshal.SizeOf(PropertyQuery)
            Dim pPropertyQuery As IntPtr = Marshal.AllocHGlobal(nBytesPropertyQuery)
            Marshal.StructureToPtr(PropertyQuery, pPropertyQuery, False)

            Dim nBytesDeviceDescriptor As Integer = Marshal.SizeOf(DeviceDescriptor) + 2048
            Dim pDeviceDescriptor As IntPtr = Marshal.AllocHGlobal(nBytesDeviceDescriptor)

            Dim bDeviceIoControl As Boolean = DeviceIoControl(hVolume, IOCTL_STORAGE_QUERY_PROPERTY, pPropertyQuery, Marshal.SizeOf(GetType(STORAGE_PROPERTY_QUERY)), pDeviceDescriptor, nBytesDeviceDescriptor, nBytesReturned, Nothing)
            Dim DeviceDescriptorRet As STORAGE_DEVICE_DESCRIPTOR = Marshal.PtrToStructure(pDeviceDescriptor, GetType(STORAGE_DEVICE_DESCRIPTOR))

            Dim nSerialNumberOffset As UInteger = DeviceDescriptorRet.SerialNumberOffset
            Dim pSerialNumber As IntPtr = IntPtr.Add(pDeviceDescriptor, CInt(nSerialNumberOffset))
            Dim sSerialNumber As String = Marshal.PtrToStringAnsi(pSerialNumber)

            Console.WriteLine("SerialNumber : {0}", sSerialNumber)

            Marshal.FreeHGlobal(pDeviceDescriptor)
            Marshal.FreeHGlobal(pPropertyQuery)
            CloseHandle(hVolume)

            Return sSerialNumber
        Else
            Return Nothing
        End If
    End Function

    Private Sub GetDisks()
        Dim DriveList As String() = Directory.GetLogicalDrives
        Dim Drive As String

        For Each Drive In DriveList
            If Drive.Contains("\") Then
                Drive = Drive.Replace("\", "")
            End If
            ComboBox1.Items.Add("" & Drive.ToUpper)
        Next
    End Sub

    Private Function GetWMISerial(PhysicalDriveTag As String) As String
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")
        For Each wmi_HD As ManagementObject In searcher.Get()
            Dim serialNumber As String = (wmi_HD("SerialNumber")).ToString()
            Dim tag As String = wmi_HD("Tag").ToString
            If tag = PhysicalDriveTag Then
                Return tag & "," & serialNumber
            Else
                Return ""
            End If
        Next
    End Function

    Public Function GetInfo(Win32_stringIn As String) As String

        Dim sbOutput As New StringBuilder(String.Empty)

        Try
            Dim mcInfo As New ManagementClass(Win32_stringIn)

            Dim mcInfoCol As ManagementObjectCollection =
                mcInfo.GetInstances()

            Dim pdInfo As PropertyDataCollection = mcInfo.Properties

            For Each objMng As ManagementObject In mcInfoCol

                For Each prop As PropertyData In pdInfo

                    Try

                        sbOutput.AppendLine(prop.Name + ":  " +
                         objMng.Properties(prop.Name).Value)

                    Catch

                    End Try

                Next

                sbOutput.AppendLine()

            Next

        Catch

        End Try

        Return sbOutput.ToString()

    End Function

    Private Sub GetAllDrives()
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
        For Each wmi_HD As ManagementObject In searcher.Get()

            Dim HDModel As String = (wmi_HD("Model")).ToString()
            Dim HDType As String = wmi_HD("InterfaceType").ToString
            Dim Caption As String = wmi_HD("Caption").ToString
            Dim Description As String = wmi_HD("Description").ToString
            Dim Manufacturer As String = wmi_HD("Manufacturer").ToString
            Dim Name As String = wmi_HD("Name").ToString
            Dim SerialNumber As String = wmi_HD("SerialNumber").ToString
            Dim Status As String = wmi_HD("Status").ToString
            ComboBox2.Items.Add(Name & "   " & HDModel & "   " & HDType)
        Next
    End Sub

    Private Function GetAllDrivesByName(PhysicalDrive As String) As String
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
        For Each wmi_HD As ManagementObject In searcher.Get()

            Dim HDModel As String = (wmi_HD("Model")).ToString()
            Dim HDType As String = wmi_HD("InterfaceType").ToString
            Dim Caption As String = wmi_HD("Caption").ToString
            Dim Description As String = wmi_HD("Description").ToString
            Dim Manufacturer As String = wmi_HD("Manufacturer").ToString
            Dim PhysicalName As String = wmi_HD("Name").ToString
            Dim SerialNumber As String = wmi_HD("SerialNumber").ToString
            Dim Status As String = wmi_HD("Status").ToString
            If PhysicalDrive = PhysicalName Then
                Return SerialNumber
            Else
                Return ""
            End If
        Next
        Return Nothing
    End Function

    Private Function GetUsbInfo()
        Dim mos As New ManagementObjectSearcher("SELECT * FROM Win32_UsbController")

        For Each mo As ManagementObject In mos.Get()
            Dim DeviceID As String = mo.Properties.Item("DeviceID").Value.ToString
            Dim Caption As String = mo.Properties.Item("Caption").Value.ToString
            Dim Description As String = mo.Properties.Item("Description").Value.ToString
            Dim Manufacturer As String = mo.Properties.Item("Manufacturer").Value.ToString
            Dim Name As String = mo.Properties.Item("Name").Value.ToString
            Dim Status As String = mo.Properties.Item("Status").Value.ToString
            Dim SystemName As String = mo.Properties.Item("SystemName").Value.ToString
            RichTextBox1.Text &= DeviceID & " " & Caption & " " & Name & " " & SystemName
        Next
        Return Nothing
    End Function

    Private Function GetSerialFSOhex(drive As String) As String
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Return Hex(fso.GetDrive(drive & "\").SerialNumber)
    End Function

    Private Function GetSerialViaPath() As String
        Dim theSearcher As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE InterfaceType='USB'")
        Dim serialnum = ""
        For Each currentObject As ManagementObject In theSearcher.Get()
            Dim theSerialNumberObjectQuery As ManagementObject = New ManagementObject()
            theSerialNumberObjectQuery.Path = New ManagementPath("Win32_PhysicalMedia.Tag='" & currentObject("DeviceID") & "'")
            serialnum &= theSerialNumberObjectQuery("SerialNumber").ToString() & vbCrLf
        Next
        Return serialnum
    End Function



    Function GetUSBSerialNo(ByVal DriveLetter As String)
        Dim PnPID As String
        PnPID = USBSerialNo(DriveLetter)

        If Not Trim(PnPID) = "" Then
            GetUSBSerialNo = formatSerialNo(PnPID)
        Else
            GetUSBSerialNo = ""
        End If

    End Function


    Function USBSerialNo(ByVal DriveLetter As String)
        Dim SerialNo As String

        Dim ComputerName = "."
        Dim wmiDiskDrive, query, wmiDiskPartition, wmiLogicalDisk

        Dim wmiServices = GetObject("winmgmts:{impersonationLevel=Impersonate}!//" & ComputerName)

        Dim wmiDiskDrives = wmiServices.ExecQuery("SELECT Caption, DeviceID,PNPDeviceID FROM Win32_DiskDrive")

        For Each wmiDiskDrive In wmiDiskDrives

            SerialNo = wmiDiskDrive.PNPDeviceID '1

            query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & wmiDiskDrive.deviceid & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
            Dim wmiDiskPartitions = wmiServices.ExecQuery(query)

            For Each wmiDiskPartition In wmiDiskPartitions
                Dim wmiLogicalDisks = wmiServices.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & wmiDiskPartition.deviceid & "'} WHERE AssocClass = Win32_LogicalDiskToPartition")

                For Each wmiLogicalDisk In wmiLogicalDisks

                    If (wmiLogicalDisk.deviceid = DriveLetter) And (wmiLogicalDisk.DriveType = 2) Then '2
                        USBSerialNo = SerialNo
                        Exit Function
                    End If

                Next
            Next
        Next
    End Function


    Function formatSerialNo(ByVal PnPID As String)
        Dim arrSerialNo
        Dim arrSerialNo1
        arrSerialNo = Split(PnPID, "\")
        Dim i
        arrSerialNo1 = Split(arrSerialNo(UBound(arrSerialNo)), "&")

        If UBound(arrSerialNo1) > 0 Then
            formatSerialNo = arrSerialNo1(UBound(arrSerialNo1) - 1)
        Else
            formatSerialNo = arrSerialNo1(UBound(arrSerialNo1))
        End If

    End Function





    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetDisks()
        Dim Info As String = GetInfo("Win32_PhysicalMedia")
        GetAllDrives()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'serial USB
        TextBox1.Text = ""

        'Dim serialnum As String = GetSerialNumberDisk(ComboBox1.SelectedItem.ToString.ToUpper)
        'RichTextBox1.Text &= GetSerialFSOhex(ComboBox1.SelectedItem.ToString.ToUpper) & vbCrLf
        'Label4.Text = serialnum

        Dim serialUSB As String = USBSerialNo(ComboBox1.SelectedItem.ToString)
        Label2.Text = serialUSB.ToString

        Dim numerodiserie() As String = Strings.Split(serialUSB.ToString, "\")
        Label8.Text = numerodiserie(2)

        'attivo la scelta USB
        CheckBox1.Enabled = True
        'attivo il copia
        Button1.Enabled = True
        TextBox1.Enabled = True
        'Sicurezza
        Label16.ForeColor = Color.White
        'Gui
        Label1.ForeColor = Color.White
        Label5.ForeColor = Color.White
        Label7.ForeColor = Color.White
        Label9.ForeColor = Color.White
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'Scelta Drive fisico
        'Label4.Text = GetWMISerial(PhysicalDrive)

        Dim drives() As String = Strings.Split(ComboBox2.SelectedItem.ToString, "   ")
        Dim PhysicalDrive As String = drives(0).ToString

        Label4.Text = GetSerialNumberDisk2(PhysicalDrive)
        Label8.Text = GetSerialNumberDisk2(PhysicalDrive)

        'attivo la scelta HDD
        CheckBox2.Enabled = True
        'attivo il copia
        Button1.Enabled = True
        TextBox1.Enabled = True
        'Sicurezza
        Label16.ForeColor = Color.White
        'Gui
        Label1.ForeColor = Color.White
        Label5.ForeColor = Color.White
        Label7.ForeColor = Color.White
        Label9.ForeColor = Color.White
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Copia usb label
        Label8.Text = ""
        TextBox1.Text = Label2.Text
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        'usb
        RichTextBox1.Text = ""
        If Label8.Text = "" Or Len(Label8.Text) < 5 Then
            If TextBox1.Text = "" Or Len(TextBox1.Text) < 5 Then
                RichTextBox1.Text = "Purtroppo il numero di serie USB sembra nullo o troppo corto" & vbCrLf
            Else
                'aggiungo al seriale quello trovato o creato a mano
                RichTextBox1.Text = "Numero di Serie USB = " & TextBox1.Text & vbCrLf
                serialusb_secure = serialusb_secure & TextBox1.Text
            End If
        Else
            'aggiungo al seriale quello trovato
            RichTextBox1.Text = "Numero di Serie USB = " & Label8.Text & vbCrLf
            serialusb_secure = serialusb_secure & Label8.Text
        End If

        RichTextBox2.Text &= "Verrà usato SN: " & serialusb_secure & vbCrLf

        'disattivo il copia e attivo pulisci
        Button1.Enabled = False
        Button2.Enabled = True
        TextBox1.Enabled = False
        Button3.Enabled = True
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'hdd
        RichTextBox1.Text = ""
        If Label4.Text = "" Or Len(Label4.Text) < 5 Then
            If TextBox1.Text = "" Or Len(TextBox1.Text) < 5 Then
                RichTextBox1.Text = "Purtroppo il numero di serie HDD sembra nullo o troppo corto" & vbCrLf
            Else
                'aggiungo al seriale quello trovato o creato a mano
                RichTextBox1.Text = "Numero di Serie HDD = " & TextBox1.Text & vbCrLf
                serialusb_secure = serialusb_secure & TextBox1.Text
            End If
        Else
            'aggiungo al seriale quello trovato 
            RichTextBox1.Text = "Numero di Serie HDD = " & Label4.Text & vbCrLf
            serialusb_secure = serialusb_secure & Label4.Text
        End If

        RichTextBox2.Text &= "Verrà usato SN: " & serialusb_secure & vbCrLf

        'disattivo il copia e attivo pulisci
        Button1.Enabled = False
        Button2.Enabled = True
        TextBox1.Enabled = False
        Button3.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Pulisci
        serialusb_secure = ""
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""

        'disattivo il copia
        Button1.Enabled = False
        TextBox1.Enabled = False
        'disattivo scelte
        CheckBox2.Enabled = False
        CheckBox1.Enabled = False

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        'Disattivo OK
        Button3.Enabled = False
        'Disattivo Crypto
        Button4.Enabled = False
        Button5.Enabled = False
        RichTextBox3.Enabled = False
        RichTextBox3.Text = ""
        RichTextBox4.Enabled = False
        RichTextBox4.Text = ""
        Label20.ForeColor = Color.Black
        Label21.ForeColor = Color.Black

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'OK Vai
        Label19.ForeColor = Color.White
        Button4.Enabled = True
        Button5.Enabled = True
        Label20.ForeColor = Color.White
        Label21.ForeColor = Color.White
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox3.Enabled = True
        RichTextBox4.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Secure
        RichTextBox4.Text = ""

        ' SecureKey dalla pendrive + hard disk
        Dim secureKey As String = serialusb_secure

        '1. String to SecureString
        Dim theSecureStringInput As SecureString = New NetworkCredential("", secureKey).SecurePassword
        '2. Encrypt 3DES
        Dim wrapper As New Simple3Des(secureKey)
        Dim cipherText3DES As String = wrapper.EncryptData(RichTextBox3.Text)
        '3. Instantiate the secure string.
        Dim securePwd As New SecureString()
        '4. Create secure string from result AES
        '   Ignore any key out of range
        For i = 1 To Len(cipherText3DES)
            Dim car As String = Strings.Mid(cipherText3DES, i, 1)
            If CInt(Asc(car)) >= 32 And CInt(Asc(car)) <= 126 Then
                ' Append the character to the password.
                securePwd.AppendChar(car)
                Console.Write("*")
            End If
        Next
        '5. Secure string to string
        Dim theStringPlain As String = New NetworkCredential("", securePwd).Password
        '6. Show result
        RichTextBox4.Text = cipherText3DES
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'dec
        RichTextBox4.Text = ""

        ' SecureKey dalla pendrive + hard disk
        Dim secureKey As String = serialusb_secure

        '2. deEncrypt 3DES
        Dim wrapper As New Simple3Des(secureKey)
        '   DecryptData throws if the wrong password is used.
        Dim plainText As String = ""
        Try
            plainText = wrapper.DecryptData(RichTextBox3.Text)
        Catch ex As System.Security.Cryptography.CryptographicException
            MsgBox("Non posso decriptare i dati con questa password.")
        End Try
        '3. Creo la secure string con la securekey
        Dim theSecureStringInput As SecureString = New NetworkCredential("", secureKey).SecurePassword
        '4. Risultato in chiaro
        RichTextBox4.Text = plainText
    End Sub
End Class

Public NotInheritable Class Simple3Des
    Private TripleDes As New TripleDESCryptoServiceProvider

    Private Function TruncateHash(
    ByVal key As String,
    ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte =
        System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function


    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub


    Public Function EncryptData(
    ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte =
        System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
        TripleDes.CreateEncryptor(),
        System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function


    Public Function DecryptData(
    ByVal encryptedtext As String) As String

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
        TripleDes.CreateDecryptor(),
        System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function
    'https://learn.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/strings/walkthrough-encrypting-and-decrypting-strings
End Class
