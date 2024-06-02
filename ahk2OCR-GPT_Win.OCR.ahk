#Requires AutoHotkey v2

#Include OCR.ahk
#Include _JXON.ahk

/*
====================================================
Variables
====================================================
*/

API_Key := "YOUR_CHAT_GPT_API"

API_URL := "https://api.openai.com/v1/chat/completions"
Status_Message := ""
Response_Window_Status := "Closed"
Retry_Status := ""

; Select Screen Region with Mouse
^#LButton:: ; Control+Win+Left Mouse to Select
{
	Area := SelectScreenRegion("LButton")
	Result := OCR.FromRect(Area.X, Area.Y, Area.W, Area.H)
	
	ChatGPT_Prompt := "YOUR_PROMT :  " . Result.Text

	API_Model := "gpt-3.5-turbo"
    ProcessRequest(ChatGPT_Prompt, Status_Message, API_Model, Retry_Status, Result.Text)	
}

Esc:: ExitApp

SelectScreenRegion(Key, Color := "Lime", Transparent:= 80)
{
	CoordMode("Mouse", "Screen")
	MouseGetPos(&sX, &sY)
	ssrGui := Gui("+AlwaysOnTop -caption +Border +ToolWindow +LastFound -DPIScale")
	WinSetTransparent(Transparent)
	ssrGui.BackColor := Color
	Loop 
	{
		Sleep 10
		MouseGetPos(&eX, &eY)
		W := Abs(sX - eX), H := Abs(sY - eY)
		X := Min(sX, eX), Y := Min(sY, eY)
		ssrGui.Show("x" X " y" Y " w" W " h" H)
	} Until !GetKeyState(Key, "p")
	ssrGui.Destroy()
	Return { X: X, Y: Y, W: W, H: H, X2: X + W, Y2: Y + H }
}

ProcessRequest(ChatGPT_Prompt, Status_Message, API_Model, Retry_Status, OriginalMess) {
    if (Retry_Status != "Retry") {
        ChatGPT_Prompt := ChatGPT_Prompt
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, '(\\|")+', '\$1') ; Clean back spaces and quotes
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, "`n", "\n") ; Clean newlines
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, "`r", "") ; Remove carriage returns
        global Previous_ChatGPT_Prompt := ChatGPT_Prompt
        global Previous_Status_Message := Status_Message
        global Previous_API_Model := API_Model
        global Response_Window_Status
    }

    global HTTP_Request := ComObject("WinHttp.WinHttpRequest.5.1")
    HTTP_Request.open("POST", API_URL, true)
    HTTP_Request.SetRequestHeader("Content-Type", "application/json")
    HTTP_Request.SetRequestHeader("Authorization", "Bearer " API_Key)    
    Messages := '{ "role": "user", "content": "' ChatGPT_Prompt '" }'
    JSON_Request := '{ "model": "' API_Model '", "messages": [' Messages '] }'
    HTTP_Request.SetTimeouts(60000, 60000, 60000, 60000)
    HTTP_Request.Send(JSON_Request)

    HTTP_Request.WaitForResponse
    try {
        if (HTTP_Request.status == 200) {
            SafeArray := HTTP_Request.responseBody
	    pData := NumGet(ComObjValue(SafeArray) + 8 + A_PtrSize, 'Ptr')
	    length := SafeArray.MaxIndex() + 1
	    JSON_Response := StrGet(pData, length, 'UTF-8')
            var := Jxon_Load(&JSON_Response)
            JSON_Response := var.Get("choices")[1].Get("message").Get("content")
            Msgbox(OriginalMess . "`n------------`n" . JSON_Response)
        } else {
			if (HTTP_Request.status == 429) {                
				MsgBox(GetJSONResponse(HTTP_Request.responseBody))				
			}
        }
    }
}

GetJSONResponse(SafeArray) {
    pData := NumGet(ComObjValue(SafeArray) + 8 + A_PtrSize, 'Ptr')
    length := SafeArray.MaxIndex() + 1
    JSON_Response := StrGet(pData, length, 'UTF-8')
    return JSON_Response
}
