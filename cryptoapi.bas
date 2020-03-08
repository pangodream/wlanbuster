Attribute VB_Name = "Module2"
'  Cryptography API Functions (CryptoAPI)
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
'  The CryptDecrypt function is not required for the RC4 algorithm since the algorithm uses a symmetrical key.
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal phProv As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwLength As Long, ByRef pbData As Any) As Long

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
' Microsoft CryptoAPI Function Constants
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////

'  Constants for CryptoAPI functions.
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const PP_VERSION As Long = 5
Private Const PP_NAME   As Long = 4

'  Provider Types
Private Const MS_BASE_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1

'  Constants for CryptoAPI functions.
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768

Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048

Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MAC = 5
Private Const ALG_SID_SSL3SHAMD5 = 8

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

'  Encryption Algorithm Constants used by CryptoAPI.
Private Const CALG_RC4 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)   ' RC4 Encryption Algorithm

'  Hashing Algorithm Constants  used by CryptoAPI.
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
        ' MD2 Hash Algorithm
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
        ' MD5 Hash Algorithm
Private Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
        ' US DSA Secure Hash Algorithm (SHA)

'  Public Enumerated Hashing Algorithms
Public Enum HashAlgorithms
   MD2 = CALG_MD2
   MD5 = CALG_MD5
   SHA = CALG_SHA
End Enum

' To invoke DSS on W2K
'CryptAcquireContext(&hProv, pszContainer, MS_DEF_DSS_DH_PROV, PROV_DSS,dwFlags)

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
'  [out] RC4Cipher(arg1 [in], arg2 [in])
'
'  Function Description:
'
'  This function utilizes the CryptoAPI provided in Microsoft NT Windows 4.0.
'  The CryptoAPI supports RC4 encryption by use of either a 42-bit or 128-bit
'  encryption key.  The 128-bit encryption key length is available with the
'  'Microsoft Enhanced Cryptographic Provider.'
'
'  ----------   --------------------------------------------------------------
'  Argument     Data Type       Description
'  ----------   --------------------------------------------------------------
'  arg1         String          Data to be encrypted/decypted.
'  arg2         String          Key (password) used to encrypt/decrypt.
'  return       String          Encrypted/decrypted data processed through
'                               the RC4 algorithm
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function RC4Cipher(ByVal data As String, ByVal Key As String) As String

   Dim lngHKey          As Long             ' Handle to the Hash Key.
   Dim lngHCryptProvider As Long   ' Handle to the CSP.
   Dim lngHHash         As Long            ' Handle to the Hash Object
   Dim lngResult        As Long           ' General API result variable.

   Const cstFuncName = "RCCipher"

   On Error GoTo ErrorProc

   RC4Cipher = vbNullString

   '  Test to make sure we have data to encrypt.
   If Len(data) = 0 Then
      Exit Function
   End If

   '  Acquire the CSP.
   '  Note:  The CRYPT_VERIFYCONTEXT flag is used with the PROV_RSA_FULLalgorithm flag.  This flag tells the CSP
   '  that the RC4 algorithm does not utilize public/private key support.  Weare only interested in private key algorithms.
   lngResult = CryptAcquireContext(lngHCryptProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Create a hash object using the MD5 algorithm for our CSP.
   lngResult = CryptCreateHash(lngHCryptProvider, CALG_MD5, 0, 0, lngHHash)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Hash in the password text using the MD5 hash object.
   '  Note:  The 4th parameter is ignored by the microsoft CSP and can be setto '0'.
   lngResult = CryptHashData(lngHHash, Key, Len(Key), 0)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Derive a key (hash it) for RC4 cyphering from the hash object.  Get ahandle to the key.
   lngResult = CryptDeriveKey(lngHCryptProvider, CALG_RC4, lngHHash, 0, lngHKey)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Destroy the MD5 hash object and clear the handle.  We only needed it toprime the RC4 algorithm.
   Call CryptDestroyHash(lngHHash)
   lngHHash = 0

   '  Encrypt/Decrypt the text data.
   '
   '  The 'Data' variable which holds the data to be cyphered will return with
   '  the cryphered data in its contents.
   '
   '  Note: The Encrypt routine also decrypts with the RC4 algorithm.
   lngResult = CryptEncrypt(lngHKey, 0, 1, 0, data, Len(data), Len(data))

   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Return the encrypted data.
   RC4Cipher = data

ExitProc:

   '  Destroy then session key.
   If lngHKey Then
      lngResult = CryptDestroyKey(lngHKey)
   End If

   '  Destroy hash object.
   If lngHHash Then
      lngResult = CryptDestroyHash(lngHHash)
   End If

   '  Release Context provider handle.
   If lngHCryptProvider Then
      lngResult = CryptReleaseContext(lngHCryptProvider, 0)
   End If

   Exit Function
ErrorProc:
   Err.Raise Err.Number, mcstClassName & "." & cstFuncName, Err.Description
   Resume ExitProc

End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
'  [out] ProviderVersion()
'
'  Property Description:
'
'  This property utilizes the CryptoAPI provided in Microsoft NT Windows 4.0.
'  The CryptoAPI supports encryption by use of either a 42-bit or 128-bit
'  encryption key.  The 128-bit encryption key length is available with the
'  'Microsoft Enhanced Cryptographic Provider.'  The 42-bit encryption key length
'  is available in the 'Microsoft Base Cryptographic Provider.'
'
'  ----------   --------------------------------------------------------------
'  Argument     Data Type       Description
'  ----------   --------------------------------------------------------------
'  return       Long            1 - Microsoft Base Cryptographic Provider(42-bit key length)
'                               2 - Microsoft Enhanced Cryptographic Provider(128-bit key length)
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get ProviderVersion() As Long

   Dim lngHCryptProvider As Long       '  Handle to the CSP.
   Dim abytBuffer(10)   As Byte          '  Byte array buffer.
   Dim lngResult        As Long               '  General API result value.

   Const cstBufferSize As Long = 10
   Const cstVersionByte As Long = 1
   Const cstFuncName As String = "ProviderVersion"

   On Error GoTo ErrorProc

   '  Acquire the CSP.
   '
   '  Note:  The CRYPT_VERIFYCONTEXT flag is used with the PROV_RSA_FULLalgorithm flag.  This flag tells the CSP
   '  that the RC4 algorithm does not utilize public/private key support.  Weare only interested in private key algorithms.
   lngResult = CryptAcquireContext(lngHCryptProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Use the get provider parameter function to acquire the version of theprovider installed.
   lngResult = CryptGetProvParam(lngHCryptProvider, PP_VERSION, abytBuffer(0), cstBufferSize, 0)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  The version number is stored in the 2nd byte of the array.  Return theversion value.
   ProviderVersion = abytBuffer(cstVersionByte)

ExitProc:

   '  Release Context provider handle.
   If lngHCryptProvider Then
      lngResult = CryptReleaseContext(lngHCryptProvider, 0)
   End If

   Exit Function
ErrorProc:
   Err.Raise Err.Number, mcstClassName & "." & cstFuncName, Err.Description
   Resume ExitProc
End Property

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
'  [out] MD5Hash(arg1 [in], arg2 [in])
'
'  Function Description:
'
'  This function utilizes the CryptoAPI provided in Microsoft NT Windows 4.0.
'  The CryptoAPI supports several hashing algorithms.  The hashing algorithm
'  is specified in 'arg2.'  The data to hash is passed in 'arg1.'  The return
'  is a string containing the hashed data.
'
'  ----------   --------------------------------------------------------------
'  Argument     Data Type       Description
'  ----------   --------------------------------------------------------------
'  arg1         String          Data to be hashed.
'
'  return       String          Hashed data.
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function MD5Hash(ByVal Key As String) As String

   Dim lngResult        As Long               '  General API result value.
   Dim lngHHash         As Long                '  Handle to the Hash object.
   Dim lngHCryptProvider As Long       '  Handle to the CSP.
   Dim abytBuffer()     As Byte            '  Byte array buffer.
   Dim lngHashLength    As Long           '  Length of the Hash.
   Dim strHash          As String               '  Hash converted to a string.
   Dim intI             As Integer                 '  Counter variable.
   Dim lngKeyHashAlgorithm As Long

   Const cstBufferSize As Long = 256   '  Buffer Size for the Byte Array.
   Const cstFuncName As String = "Hash"

   On Error GoTo ErrorProc

   MD5Hash = vbNullString
   lngKeyHashAlgorithm = MD5

   '  Resize the byte array buffer.
   ReDim abytBuffer(cstBufferSize)

   '  Acquire the CSP.
   '
   '  Note:  The CRYPT_VERIFYCONTEXT flag is used with the PROV_RSA_FULLalgorithm flag.  This flag tells the CSP
   '  that the RC4 algorithm does not utilize public/private key support.  Weare only interested in private key algorithms.
   lngResult = CryptAcquireContext(lngHCryptProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Create a hash object using the specified algorithm supported by our CSP.
   lngResult = CryptCreateHash(lngHCryptProvider, lngKeyHashAlgorithm, 0, 0, lngHHash)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Hash in the password text using the new hash object.
   '  Note:  The 4th parameter is ignored by the microsoft CSP.
   lngResult = CryptHashData(lngHHash, Key, Len(Key), 0)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Get the hash size (HP_HASHSIZE).
   '  The hash value size. The pbData buffer will contain a DWORD valueindicating the number of bytes
   '  in the hash value. This value will usually be 16 or 20, depending on thehash algorithm.
   lngResult = CryptGetHashParam(lngHHash, HP_HASHSIZE, abytBuffer(0), cstBufferSize, 0)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  The length of the hash can be found in the first byte array position ofthe buffer.
   lngHashLength = abytBuffer(0)

   '  Get the hash value (HP_HASHVAL).
   '  The hash value. The pbData buffer will contain the hash value or messagehash for the
   '  hash object specified by hHash. This value is generated based on the datasupplied earlier to the
   '  hash object through the CryptHashData and CryptHashSessionKey functions.
   lngResult = CryptGetHashParam(lngHHash, HP_HASHVAL, abytBuffer(0), cstBufferSize, 0)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Convert the byte array to a VB string.
   For intI = 0 To lngHashLength
      strHash = strHash & Chr(abytBuffer(intI))
   Next intI

   '  Return the hashed data as a VB string.
   MD5Hash = strHash

ExitProc:

   '  Destroy hash object.
   If lngHHash Then
      lngResult = CryptDestroyHash(lngHHash)
   End If

   '  Release Context provider handle.
   If lngHCryptProvider Then
      lngResult = CryptReleaseContext(lngHCryptProvider, 0)
   End If

   Exit Function
ErrorProc:
   Err.Raise Err.Number, mcstClassName & "." & cstFuncName, Err.Description
   Resume ExitProc
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
'  [out] GenerateRandomNumbers(arg1 [in])
'
'  Property Description:
'
'  This function utilizes the CryptoAPI provided in Microsoft NT Windows 4.0.
'  The purpose of this function is to provide access to the superior random Number
'  generator provided by the Microsoft CryptoAPI.
'
'  ----------   --------------------------------------------------------------
'  Argument     Data Type       Description
'  ----------   --------------------------------------------------------------
'  arg1         Long            The length of the array of random numbers to bereturned.
'  return       Byte Array      A byte array containing the random generatednumbers.
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GenerateRandomNumbers(ByVal LengthOfArray As Long) As Byte()

   Dim lngResult        As Long
   Dim lngHCryptProvider As Long
   Dim abytBuffer()     As Byte
   Dim intI             As Integer
   Dim lngRandomNumber  As Long

   Const cstFuncName As String = "GenerateRandomNumbers"

   On Error GoTo ErrorProc

   '  Prepare the buffer size. (zero based array)
   ReDim abytBuffer(LengthOfArray - 1)

   '  Acquire the CSP.
   '
   '  Note:  The CRYPT_VERIFYCONTEXT flag is used with the PROV_RSA_FULLalgorithm flag.  This flag tells the CSP
   '  that the RC4 algorithm does not utilize public/private key support.  Weare only interested in private key algorithms.
   lngResult = CryptAcquireContext(lngHCryptProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   lngResult = CryptGenRandom(lngHCryptProvider, LengthOfArray, abytBuffer(0))
   If lngResult <> 1 Then
      Err.Raise Err.LastDllError, mcstClassName, Err.Description
   End If

   '  Return the byte array.
   GenerateRandomNumbers = abytBuffer

ExitProc:

   '  Release Context provider handle.
   If lngHCryptProvider Then
      lngResult = CryptReleaseContext(lngHCryptProvider, 0)
   End If

   Exit Function
ErrorProc:
   Err.Raise Err.Number, mcstClassName & "." & cstFuncName, Err.Description
   Resume ExitProc
End Function

