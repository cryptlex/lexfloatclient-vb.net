
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

        Private productId As String = Nothing
        Private handle As UInteger = 0

        '
        '            FUNCTION: SetProductId()
        '
        '            PURPOSE: Sets the product id of your application.
        '
        '            PARAMETERS:
        '            * productId - the unique product id of your application as mentioned
        '              on the product version page of your application in the dashboard.
        '
        '            RETURN CODES: LF_OK, LF_E_PFILE, LF_E_PRODUCT_ID
        '        
        Public Function SetProductId(productId As String) As Integer
            Me.productId = productId
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.GetHandle_x64(productId, handle), Native.GetHandle(productId, handle))
#Else
            Return Native.GetHandle(productId, handle)
#End If
        End Function

        '
        '            FUNCTION: SetFloatServer()
        '
        '            PURPOSE: Sets the network address of the LexFloatServer.
        '
        '            PARAMETERS:
        '            * handle - handle for the product id
        '            * hostAddress - hostname or the IP address of the LexFloatServer
        '            * port - port of the LexFloatServer
        '
        '            RETURN CODES: LF_OK, LF_E_HANDLE, LF_E_PRODUCT_ID, LF_E_SERVER_ADDRESS
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
        '            * handle - handle for the product id
        '            * callback - name of the callback function
        '
        '            RETURN CODES: LF_OK, LF_E_HANDLE, LF_E_PRODUCT_ID
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
        '            * handle - handle for the product id
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_PRODUCT_ID, LF_E_SERVER_ADDRESS,
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
        '            * handle - handle for the product id
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_PRODUCT_ID, LF_E_SERVER_ADDRESS,
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
        '            * handle - handle for the product id
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_PRODUCT_ID, LF_E_SERVER_ADDRESS,
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
        '            FUNCTION: GetLicenseMetadata()
        '
        '            PURPOSE: Get the value of the license metadata field associated with the float server key.
        '
        '            PARAMETERS:
        '            * handle - handle for the product id
        '            * key - key of the metadata field whose value you want to get
        '            * value - pointer to a buffer that receives the value of the string
        '            * length - size of the buffer pointed to by the value parameter
        '
        '            RETURN CODES: LF_OK, LF_FAIL, LF_E_HANDLE, LF_E_PRODUCT_ID, LF_E_SERVER_ADDRESS,
        '            LF_E_CALLBACK, LF_E_BUFFER_SIZE, LF_E_METADATA_KEY_NOT_FOUND, LF_E_INET, LF_E_TIME,
        '            LF_E_SERVER_TIME
        '        


        Public Function GetLicenseMetadata(key As String, value As StringBuilder, length As Integer) As Integer
#If LF_ANY_CPU Then
            Return If(IntPtr.Size = 8, Native.GetLicenseMetadata_x64(handle, key, value, length), Native.GetLicenseMetadata(handle, key, value, length))
#Else
            Return Native.GetLicenseMetadata(handle, key, value, length)
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

        Public Enum StatusCodes As UInteger

            '
            '    CODE: LF_OK

            '    MESSAGE: Success code.
            '
            LF_OK = 0

            '
            '    CODE: LF_FAIL

            '    MESSAGE: Failure code.
            '
            LF_FAIL = 1

            '
            '    CODE: LF_E_PRODUCT_ID

            '    MESSAGE: The product id is incorrect.
            LF_E_PRODUCT_ID = 40

            '
            '    CODE: LF_E_CALLBACK

            '    MESSAGE: Invalid or missing callback function.
            '
            LF_E_CALLBACK = 41

            '
            '    CODE: LF_E_HANDLE

            '    MESSAGE: Invalid handle.
            LF_E_HANDLE = 42

            '
            '    CODE: LF_E_SERVER_ADDRESS

            '    MESSAGE: Missing or invalid server address.
            LF_E_SERVER_ADDRESS = 43

            '
            '    CODE: LF_E_SERVER_TIME

            '    MESSAGE: System time on Server Machine has been tampered with. Ensure
            '    your date and time settings are correct on the server machine.
            LF_E_SERVER_TIME = 44

            '
            '    CODE: LF_E_TIME

            '    MESSAGE: The system time has been tampered with. Ensure your date
            '    and time settings are correct.
            LF_E_TIME = 45

            '
            '    CODE: LF_E_INET

            '    MESSAGE: Failed to connect to the server due to network error.
            LF_E_INET = 46

            '
            '    CODE: LF_E_NO_FREE_LICENSE

            '    MESSAGE: No free license is available
            LF_E_NO_FREE_LICENSE = 47

            '
            '    CODE: LF_E_LICENSE_EXISTS

            '    MESSAGE: License has already been leased.
            LF_E_LICENSE_EXISTS = 48

            '
            '    CODE: LF_E_LICENSE_EXPIRED

            '    MESSAGE: License lease has expired. This happens when the
            '    request to refresh the license fails due to license been taken
            '    up by some other client.
            LF_E_LICENSE_EXPIRED = 49

            '
            '    CODE: LF_E_LICENSE_EXPIRED_INET

            '    MESSAGE: License lease has expired due to network error. This
            '    happens when the request to refresh the license fails due to
            '    network error.
            LF_E_LICENSE_EXPIRED_INET = 50

            '
            '    CODE: LF_E_BUFFER_SIZE

            '    MESSAGE: The buffer size was smaller than required.
            LF_E_BUFFER_SIZE = 51

            '
            '    CODE: LF_E_METADATA_KEY_NOT_FOUND

            '    MESSAGE: The metadata key does not exist.
            LF_E_METADATA_KEY_NOT_FOUND = 52

            '
            '    CODE: LF_E_SERVER

            '    MESSAGE: Server error.
            LF_E_SERVER = 70

            '
            '    CODE: LF_E_CLIENT

            '    MESSAGE: Client error.
            LF_E_CLIENT = 71
            
        End Enum

        <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
        Public Delegate Sub CallbackType(status As UInteger)

        ' To prevent garbage collection of delegate, need to keep a reference 


        Shared leaseCallback As CallbackType

        Private NotInheritable Class Native
            Private Sub New()
            End Sub
            
            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetHandle(productId As String, ByRef handle As UInteger) As Integer
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
            Public Shared Function GetLicenseMetadata(handle As UInteger, key As String, value As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GlobalCleanUp() As Integer
            End Function

#If LF_ANY_CPU Then

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetHandle", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetHandle_x64(productId As String, ByRef handle As UInteger) As Integer
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

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetLicenseMetadata", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetLicenseMetadata_x64(handle As UInteger, key As String, value As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GlobalCleanUp", CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GlobalCleanUp_x64() As Integer
            End Function

#End If
        End Class
    End Class
End Namespace
