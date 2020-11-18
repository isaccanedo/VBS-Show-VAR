' Script de exemplo do Windows Script Host
'
' ------------------------------------------------------------------------                                    
' Autor:             Isac Canedo                                                 
' Data Criação:      02/09/2009  
' ------------------------------------------------------------------------

' Este exemplo relacionará todas as variáveis de ambiente definidas neste computador.

L_Welcome_MsgBox_Message_Text    = "Este script relacionará todas as variáveis de ambiente definidas neste computador."
L_Welcome_MsgBox_Title_Text      = "Exemplo do Windows Scripting Host"
Call Welcome()


' ********************************************************************************
' *
' * Exemplo de ambiente
' *
CRLF = Chr(13) & Chr(10)

Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")


Sub show_env(strText)
    MsgBox strText, vbInformation, L_Welcome_MsgBox_Title_Text
End Sub

intIndex = 0
strText = ""
intNumEnv = 0
MAX_ENV = 20
For Each strEnv In WshShell.Environment("PROCESS")
    intIndex = intIndex + 1
    strText = strText & CRLF & Right("    " & intIndex, 4) & " " & strEnv
    intNumEnv = intNumEnv + 1
    If intNumEnv >= MAX_ENV Then
        Call show_env(strText)
        strText = ""
        intNumEnv = 0
    End If
Next

If intNumEnv >= 1 Then Call show_env(strText)


' ********************************************************************************
' *
' * Bem-vindo
' *
Sub Welcome()
    Dim intDoIt

    intDoIt =  MsgBox(L_Welcome_MsgBox_Message_Text, _
                      vbOKCancel + vbInformation,    _
                      L_Welcome_MsgBox_Title_Text )
    If intDoIt = vbCancel Then
        WScript.Quit
    End If
End Sub
