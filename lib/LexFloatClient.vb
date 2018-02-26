
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Cryptlex
    Public Class LexFloatClient
        Private Const DLL_FILE_NAME As String = "LexFloatClient.dll"

        '
        '           In order to use "Any CPU" configuration, rename 64 bit LexFloatClient.dll to LexFloatClient64.dll and uncomment 
        '	        "LF_ANY_CPU" conditional compiler constant.
        '
        '
#Const LF_ANY_CPU = 1

#If LF_ANY_CPU Then
        Private Const DLL_FILE_NAME_X64 As String = "LexFloatClient64.dll"
#End If

        Private vGUID As String = Nothing
        Private handle As UInteger = 0

        '
        '            FUNCTION: SetProductFile()
        '
        '            PURPOSE: Sets the path of the Product.dat file. This should be
        '            used if your application and Product.dat file are in different
        '            folders or you have renamed the Product.dat file.
        '
        '            If this function is used, it must be called on every start of
        '            your program before any other functions are called.
        '
        '            PARAMETERS:
        '            * filePath - path of the product file (Product.dat)
        '
        '            RETURN CODES: LF_OK, LF_E_FPATH, LF_E_PFILE
        '
        '            NOTE: If this function fails to set the path of product file, none of the
        '            other functions will work.
        '        


        Public Shared Function SetProductFile(filePath As String) As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.SetProductFile_x64(filePath), Native.SetProductFile(filePath))
#Else
            Return Native.SetProductFile(filePath)
#End If

        End Function

        '
        '            FUNCTION: SetVersionGUID()
        '
        '            PURPOSE: Sets the version GUID of your application.
        '
        '            PARAMETERS:
        '            * versionGUID - the unique version GUID of your application as mentioned
        '              on the product version page of your application in the dashboard.
        '
        '            RETURN CODES: LF_OK, LF_E_PFILE, LF_E_GUID
        '        


        Public Function SetVersionGUID(versionGUID As String) As Integer
            Me.vGUID = versionGUID
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.GetHandle_x64(versionGUID, handle), Native.GetHandle(versionGUID, handle))
#Else
            Return Native.GetHandle(versionGUID, handle)
#End If
        End Function

        '
        '            FUNCTION: SetFloatServer()
        '
        '            PURPOSE: Sets the network address of the LexFloatServer.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '            * hostAddress - hostname or the IP address of the LexFloatServer
        '            * port - port of the LexFloatServer
        '
        '            RETURN CODES: LF_OK, LF_E_HANDLE, LF_E_GUID, LF_E_SERVER_ADDRESS
        '        


        Public Function SetFloatServer(hostAddress As String, port As UShort) As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.SetFloatServer_x64(handle, hostAddress, port), Native.SetFloatServer(handle, hostAddress, port))
#Else
            Return Native.SetFloatServer(handle, hostAddress, port)
#End If

        End Function

        '
        '            FUNCTION: SetLicenseCallback()
        '
        '            PURPOSE: Sets refresh license error callback function.
        '
        '            Whenever the lease expires, a refresh lease request is sent to the
        '            server. If the lease fails to refresh, refresh license callback function
        '            gets invoked with the following status error codes: LF_E_LICENSE_EXPIRED,
        '            LF_E_LICENSE_EXPIRED_INET, LF_E_SERVER_TIME, LF_E_TIME.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '            * callback - name of the callback function
        '
        '            RETURN CODES: LF_OK, LF_E_HANDLE, LF_E_GUID
        '        


        Public Function SetLicenseCallback(callback As CallbackType) As Integer
            Dim wrappedCallback = callback
            Dim syncTarget = TryCast(callback.Target, System.Windows.Forms.Control)
            If syncTarget IsNot Nothing Then
                wrappedCallback = Function(v) syncTarget.Invoke(callback, New Object() {v})
            End If
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.SetLicenseCallback_x64(handle, wrappedCallback), Native.SetLicenseCallback(handle, wrappedCallback))
#Else
            Return Native.SetLicenseCallback(handle, wrappedCallback)
#End If

        End Function

        '
        '            FUNCTION: RequestLicense()
        '
        '            PURPOSE: Sends the request to lease the license from the LexFloatServer.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_GUID, LF_E_SERVER_ADDRESS,
        '            LF_E_CALLBACK, LF_E_LICENSE_EXISTS, LF_E_INET, LF_E_NO_FREE_LICENSE, LF_E_TIME,
        '            LF_E_PRODUCT_VERSION, LF_E_SERVER_TIME
        '        


        Public Function RequestLicense() As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.RequestLicense_x64(handle), Native.RequestLicense(handle))
#Else
            Return Native.RequestLicense(handle)
#End If

        End Function

        '
        '            FUNCTION: DropLicense()
        '
        '            PURPOSE: Sends the request to drop the license from the LexFloatServer.
        '
        '            Call this function before you exit your application to prevent zombie licenses.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_GUID, LF_E_SERVER_ADDRESS,
        '            LF_E_CALLBACK, LF_E_INET, LF_E_TIME, LF_E_SERVER_TIME
        '        


        Public Function DropLicense() As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.DropLicense_x64(handle), Native.DropLicense(handle))
#Else
            Return Native.DropLicense(handle)
#End If

        End Function

        '
        '            FUNCTION: HasLicense()
        '
        '            PURPOSE: Checks whether any license has been leased or not. If yes,
        '            it retuns LF_OK.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_GUID, LF_E_SERVER_ADDRESS,
        '            LF_E_CALLBACK
        '        


        Public Function HasLicense() As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.HasLicense_x64(handle), Native.HasLicense(handle))
#Else
            Return Native.HasLicense(handle)
#End If

        End Function

        '
        '            FUNCTION: GetCustomLicenseField()
        '
        '            PURPOSE: Get the value of the custom field associated with the float server key.
        '
        '            PARAMETERS:
        '            * handle - handle for the version GUID
        '            * fieldId - id of the custom field whose value you want to get
        '            * fieldValue - pointer to a buffer that receives the value of the string
        '            * length - size of the buffer pointed to by the fieldValue parameter
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_GUID, LF_E_SERVER_ADDRESS,
        '            LF_E_CALLBACK, LF_E_BUFFER_SIZE, LF_E_CUSTOM_FIELD_ID, LF_E_INET, LF_E_TIME,
        '            LF_E_SERVER_TIME
        '        


        Public Function GetCustomLicenseField(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.GetCustomLicenseField_x64(handle, fieldId, fieldValue, length), Native.GetCustomLicenseField(handle, fieldId, fieldValue, length))
#Else
            Return Native.GetCustomLicenseField(handle, fieldId, fieldValue, length)
#End If

        End Function

        '
        '            FUNCTION: GlobalCleanUp()
        '
        '            PURPOSE: Releases the resources acquired for sending network requests.
        '
        '            Call this function before you exit your application.
        '
        '            RETURN CODES: LF_OK
        '
        '            NOTE: This function does not drop any leased license on the LexFloatServer.
        '        


        Public Shared Function GlobalCleanUp() As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.GlobalCleanUp_x64(), Native.GlobalCleanUp())
#Else
            Return Native.GlobalCleanUp()
#End If

        End Function

        '** Return Codes **


        Public Const LF_OK As Integer = &H0

        Public Const LF_FAIL As Integer = &H1

        '
        '            CODE: LF_E_INET
        '
        '            MESSAGE: Failed to connect to the server due to network error.
        '        


        Public Const LF_E_INET As Integer = &H2

        '
        '            CODE: LF_E_CALLBACK
        '
        '            MESSAGE: Invalid or missing callback function.
        '        


        Public Const LF_E_CALLBACK As Integer = &H3

        '
        '            CODE: LF_E_NO_FREE_LICENSE
        '
        '            MESSAGE: No free license is available
        '        


        Public Const LF_E_NO_FREE_LICENSE As Integer = &H4

        '
        '            CODE: LF_E_LICENSE_EXISTS
        '
        '            MESSAGE: License has already been leased.
        '        


        Public Const LF_E_LICENSE_EXISTS As Integer = &H5

        '
        '            CODE: LF_E_HANDLE
        '
        '            MESSAGE: Invalid handle.
        '        


        Public Const LF_E_HANDLE As Integer = &H6

        '
        '            CODE: LF_E_LICENSE_EXPIRED
        '
        '            MESSAGE: License lease has expired. This happens when the
        '            request to refresh the license fails due to license been taken
        '            up by some other client.
        '        


        Public Const LF_E_LICENSE_EXPIRED As Integer = &H7

        '
        '            CODE: LF_E_LICENSE_EXPIRED_INET
        '
        '            MESSAGE: License lease has expired due to network error. This 
        '            happens when the request to refresh the license fails due to
        '            network error.
        '        


        Public Const LF_E_LICENSE_EXPIRED_INET As Integer = &H8

        '
        '            CODE: LF_E_SERVER_ADDRESS
        '
        '            MESSAGE: Missing server address.
        '        


        Public Const LF_E_SERVER_ADDRESS As Integer = &H9

        '
        '            CODE: LF_E_PFILE
        '
        '            MESSAGE: Invalid or corrupted product file.
        '        


        Public Const LF_E_PFILE As Integer = &HA

        '
        '            CODE: LF_E_FPATH
        '
        '            MESSAGE: Invalid product file path.
        '        


        Public Const LF_E_FPATH As Integer = &HB

        '
        '            CODE: LF_E_PRODUCT_VERSION
        '
        '            MESSAGE: The version GUID of the client and server don't match.
        '        


        Public Const LF_E_PRODUCT_VERSION As Integer = &HC

        '
        '            CODE: LF_E_GUID
        '
        '            MESSAGE: The version GUID doesn't match that of the product file.
        '        


        Public Const LF_E_GUID As Integer = &HD

        '
        '            CODE: LF_E_SERVER_TIME
        '
        '            MESSAGE: System time on Server Machine has been tampered with. Ensure 
        '            your date and time settings are correct on the server machine.
        '        


        Public Const LF_E_SERVER_TIME As Integer = &HE

        '
        '            CODE: LF_E_TIME
        '
        '            MESSAGE: The system time has been tampered with. Ensure your date
        '            and time settings are correct.
        '        


        Public Const LF_E_TIME As Integer = &H10

        '
        '            CODE: LF_E_CUSTOM_FIELD_ID
        '
        '            MESSAGE: Invalid custom field id.
        '        


        Public Const LF_E_CUSTOM_FIELD_ID As Integer = &H11

        '
        '            CODE: LF_E_BUFFER_SIZE
        '
        '            MESSAGE: The buffer size was smaller than required.
        '        


        Public Const LF_E_BUFFER_SIZE As Integer = &H12

        <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
        Public Delegate Sub CallbackType(status As UInteger)

        ' To prevent garbage collection of delegate, need to keep a reference 


        Shared leaseCallback As CallbackType

        Private NotInheritable Class Native
            Private Sub New()
            End Sub
            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetProductFile(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetHandle(versionGUID As String, ByRef handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetFloatServer(handle As UInteger, hostAddress As String, port As UShort) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetLicenseCallback(handle As UInteger, callback As CallbackType) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function RequestLicense(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function DropLicense(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function HasLicense(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetCustomLicenseField(handle As UInteger, fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GlobalCleanUp() As Integer
            End Function

#If LF_ANY_CPU Then
            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetProductFile", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetProductFile_x64(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetHandle", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetHandle_x64(versionGUID As String, ByRef handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetFloatServer", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetFloatServer_x64(handle As UInteger, hostAddress As String, port As UShort) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetLicenseCallback", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetLicenseCallback_x64(handle As UInteger, callback As CallbackType) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="RequestLicense", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function RequestLicense_x64(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="DropLicense", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function DropLicense_x64(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="HasLicense", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function HasLicense_x64(handle As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetCustomLicenseField", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetCustomLicenseField_x64(handle As UInteger, fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GlobalCleanUp", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GlobalCleanUp_x64() As Integer
            End Function

#End If
        End Class
    End Class
End Namespace
