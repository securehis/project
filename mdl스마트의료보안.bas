Attribute VB_Name = "mdl스마트의료보안"
Option Explicit

'170425 케이사인 모듈 연계-------------------------------------------------------------------------
Private Declare Function C_GetCertToMobile Lib "ComKSignHIS.dll" () As Long
Private Declare Function C_GenSignedData Lib "ComKSignHIS.dll" (ByVal x As String, y&) As Long
Private Declare Function C_VerifySignedDataForHIS Lib "ComKSignHIS.dll" (ByVal x As String, y&) As Long
Private Declare Function C_VerifyCert Lib "ComKSignHIS.dll" (ByVal x As String, ret&) As Long
Private Declare Function C_GetDn Lib "ComKSignHIS.dll" () As Long

Private Declare Function SysFreeString Lib "OleAut32.dll" (source As Any)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

'170425 소프트캠프--------------------------------------------------------------------------------
Private Declare Function SCDoEncryptFile Lib "SCHISPRequest.dll" (ByVal FilePath As Long) As Integer
Private Declare Function SCADDCaptureProtect Lib "SCHISPRequest.dll" (ByVal ProcessName As Long, ByVal Title As Long) As Integer
Private Declare Function SCRELEASECaptureProtect Lib "SCHISPRequest.dll" (ByVal ProcessName As Long, ByVal Title As Long) As Integer
'170522 소프트캠프 연동 함수 추가
Private Declare Function SCDoHISPLogin Lib "SCHISPLogin.dll" (ByVal UserID As Long, ByVal FullPath As Long, ByVal Position As Integer) As Integer
Private Declare Function SCDoHISPLogout Lib "SCHISPLogin.dll" (ByVal FullPath As Long) As Integer
'180111 소프트캠프 DB암호화 기능 함수 추가
Private Declare Function SCDoEncryptDBFile Lib "SCHISPRequest.dll" (ByVal FilePath As Long) As Integer
Private Declare Function SCDoDecryptDBFile Lib "SCHISPRequest.dll" (ByVal FilePath As Long) As Integer

Public lngD_KS인증주소 As Long
Public blnD_KS인증결과 As Boolean
Public lngD_KS인증결과 As Long
Public strD_KS전자서명값 As String
Public strD_KSDN값 As String
Public strD_SC보안경로 As String
Public strD_SC프로세스 As String
Public strD_SC타이틀 As String
Public intD_SC결과 As Integer

Function StringFromBSTR(ByVal pointer As Long) As String
    Dim temp As String
    CopyMemory ByVal VarPtr(temp), pointer, 4
    StringFromBSTR = temp
    CopyMemory ByVal VarPtr(temp), 0&, 4
End Function

'케이사인 모바일 저장된 인증서 요청
Public Function fD_인증서요청() As Boolean

On Error GoTo fD_인증서요청_Error:

    fD_인증서요청 = False

    lngD_KS인증결과 = C_GetCertToMobile

    If lngD_KS인증결과 = 0 Then
        fD_인증서요청 = True
    Else
        MsgBox "인증서 로그인이 실패되었습니다. 입력하신 정보를 확인해주세요.", vbCritical, "인증서 로그인 실패"
    End If

    Exit Function

fD_인증서요청_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_인증서요청 of Module mdl스마트의료보안"

End Function

'케이사인 전자서명생성
Public Function fD_전자서명생성()
On Error GoTo fD_전자서명생성_Error:
Dim s전자서명값 As String

    fD_전자서명생성 = ""

    'lngD_KS인증결과 = C_GenSignedData(strR_전자서명데이터, lngD_KS인증주소&)     '실제
    lngD_KS인증결과 = C_GenSignedData("planData", lngD_KS인증주소&)             '테스트

    If lngD_KS인증결과 = 0 Then s전자서명값 = StringFromBSTR(lngD_KS인증주소)


    Exit Function

fD_전자서명생성_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_전자서명생성 of Module mdl스마트의료보안"

End Function

'케이사인 전자서명검증
Public Function fD_전자서명검증()
On Error GoTo fD_전자서명검증_Error:
Dim s전자서명값 As String

    lngD_KS인증결과 = C_VerifySignedDataForHIS(s전자서명값, lngD_KS인증주소&)

    If lngD_KS인증결과 = 0 Then
'        MsgBox "정상적으로 검증되었습니다"
    Else
        MsgBox "서명검증 오류"
    End If

    Exit Function

fD_전자서명검증_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_전자서명검증 of Module mdl스마트의료보안"

End Function

'케이사인 인증서상태검증
Public Sub fD_인증상태검증()
On Error GoTo fD_인증상태검증_Error:
Dim s전자서명값 As String

    lngD_KS인증결과 = C_VerifyCert(s전자서명값, lngD_KS인증주소&)

    If lngD_KS인증결과 = 0 Then
'        MsgBox "정상적으로 검증되었습니다"
    Else
        MsgBox "유효한 인증서 입니다"
    End If

    Exit Sub

fD_인증상태검증_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_인증상태검증 of Module mdl스마트의료보안"

End Sub

'케이사인 DN값 읽어오기
Public Function fD_DN값() As String
On Error GoTo fD_DN값_Error:
Dim Dn As Long


    fD_DN값 = ""
    Dn = C_GetDn

    If StringFromBSTR(Dn) <> "" Then
        fD_DN값 = StringFromBSTR(Dn)
    Else
        MsgBox "DN값이 없습니다"
    End If

Exit Function
fD_DN값_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_DN값 of Module mdl스마트의료보안"
End Function


'소프트캠프 파일암복호화
Public Function fD_보안경로(strR_경로 As String) As Boolean
On Error GoTo fD_보안경로_Error:

    intD_SC결과 = SCDoEncryptFile(StrPtr(strR_경로))

    If Not intD_SC결과 = 1 Then
        MsgBox "파일보안 실패" & "[ 경로" & strR_경로 & "]", vbCritical, "소프트캠프 연동"
    End If

    Exit Function

fD_보안경로_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_보안경로 of Module mdl스마트의료보안"

End Function

'소프트캠프 화면보안 설정
Public Function fD_보안설정(strR_프로세스 As String, strR_타이틀 As String) As Boolean
On Error GoTo fD_보안설정_Error:


    intD_SC결과 = SCADDCaptureProtect(StrPtr(strR_프로세스), StrPtr(strR_타이틀))

    If Not intD_SC결과 = 1 Then
        MsgBox "화면보안설정 실패" & vbCrLf & "[ Process:" & StrPtr(strR_프로세스) & vbCrLf & "Title:" & StrPtr(strR_타이틀) & " ]", vbCritical, "소프트캠프 연동"
    End If

    Exit Function

fD_보안설정_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_보안설정 of Module mdl스마트의료보안"

End Function
'소프트캠프 화면보안 풀기
Public Function fD_보안해제(strR_프로세스 As String, strR_타이틀 As String) As Boolean
On Error GoTo fD_보안해제_Error:


    intD_SC결과 = SCRELEASECaptureProtect(StrPtr(strR_프로세스), StrPtr(strR_타이틀))

    If Not intD_SC결과 = 1 Then
        MsgBox "화면보안해제 실패" & vbCrLf & "[ Process:" & StrPtr(strR_프로세스) & vbCrLf & "Title:" & StrPtr(strR_타이틀) & " ]", vbCritical, "소프트캠프 연동"
    End If

    Exit Function

fD_보안해제_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_보안해제 of Module mdl스마트의료보안"

End Function
'소프트캠프 로그인 인증
Public Function fD_로그인(strR_아이디 As String, strR_경로 As String, lngG_권한 As Long) As Boolean
On Error GoTo fD_로그인_Error:
    Dim strR_Base64변환 As String

    strR_Base64변환 = GetBase64String(StrConv(strR_아이디, vbFromUnicode))

    intD_SC결과 = SCDoHISPLogin(StrPtr(strR_Base64변환), StrPtr(strR_경로), lngG_권한)

    If intD_SC결과 = 1 Then

    Else
        MsgBox "로그인 인증 실패" & intD_SC결과 & vbCrLf & "[ ID:" & strR_아이디 & vbCrLf & "Path:" & strR_경로 & " ]", vbCritical, "소프트캠프 연동"
    End If

    Exit Function

fD_로그인_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_로그인 of Module mdl스마트의료보안"

End Function
'소프트캠프 로그아웃 인증
Public Function fD_로그아웃(strR_경로 As String) As Boolean
On Error GoTo fD_로그아웃_Error:

    intD_SC결과 = SCDoHISPLogout(StrPtr(strR_경로))

    Exit Function

fD_로그아웃_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_로그아웃 of Module mdl스마트의료보안"

End Function

'소프트캠프 DB암복호화
Public Function fD_DB암호화(strR_경로 As String, strR_구분 As String) As Boolean
On Error GoTo fD_DB암호화_Error:

    If strR_구분 = "암호화" Then
        intD_SC결과 = SCDoEncryptDBFile(StrPtr(strR_경로))
    ElseIf strR_구분 = "복호화" Then
        intD_SC결과 = SCDoDecryptDBFile(StrPtr(strR_경로))
    End If

    If Not intD_SC결과 = 1 Then
        MsgBox "DB암호화 실패" & "[ 경로" & strR_경로 & "]", vbCritical, "소프트캠프 연동"
    End If

    Exit Function
fD_DB암호화_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_보안경로 of Module mdl스마트의료보안"

End Function

'소프트캠프 권한정보 호출
Public Function fD_권한정보호출() As Boolean
On Error GoTo fD_권한정보호출_Error:

'    fD_권한정보호출 = SCUserInfoDlg()

fD_권한정보호출_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_권한정보호출 of Module mdl스마트의료보안"

End Function

'소프트캠프 반입내역 보기
Public Function fD_반입내역() As Boolean
On Error GoTo fD_반입내역_Error:

'    fD_반입내역 = SCVRoomLogView()

fD_반입내역_Error:
    Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure fD_반입내역 of Module mdl스마트의료보안"

End Function
'BASE64 인코딩
Private Function GetBase64String(ByRef pUnicodeByteArray() As Byte) As String

    Dim pDOMDocument   As MSXML2.DOMDocument
    Dim pXMLDOMElement As MSXML2.IXMLDOMElement

    Set pDOMDocument = New MSXML2.DOMDocument
    Set pXMLDOMElement = pDOMDocument.createElement("b64")

    pXMLDOMElement.dataType = "bin.base64"
    pXMLDOMElement.nodeTypedValue = pUnicodeByteArray

    GetBase64String = pXMLDOMElement.Text

    Set pXMLDOMElement = Nothing
    Set pDOMDocument = Nothing

End Function
'BASE64 디코딩
Private Function GetUnicodeString(ByVal strBase64 As String) As Byte()

    Dim pDOMDocument   As MSXML2.DOMDocument
    Dim pXMLDOMElement As MSXML2.IXMLDOMElement

    Set pDOMDocument = New MSXML2.DOMDocument
    Set pXMLDOMElement = pDOMDocument.createElement("b64")

    pXMLDOMElement.dataType = "bin.base64"
    pXMLDOMElement.Text = strBase64

    GetUnicodeString = pXMLDOMElement.nodeTypedValue

    Set pXMLDOMElement = Nothing
    Set pDOMDocument = Nothing

End Function
