Attribute VB_Name = "ModClipboardCopy"
Option Explicit

'ClipboardCopy・・・元場所：FukamiAddins3.ModClipboard

'------------------------------



'------------------------------


Public Sub ClipboardCopy(ByVal InputClipText, Optional MessageIrunaraTrue As Boolean = False)
'入力テキストをクリップボードに格納
'配列ならば列方向をTabわけ、行方向を改行する。
'20210719作成
    
    '入力した引数が配列か、配列の場合は1次元配列か、2次元配列か判定
    Dim HairetuHantei%
    Dim Jigen1%, Jigen2%
    If IsArray(InputClipText) = False Then
        '入力引数が配列でない
        HairetuHantei = 0
    Else
        On Error Resume Next
        Jigen2 = UBound(InputClipText, 2)
        On Error GoTo 0
        
        If Jigen2 = 0 Then
            HairetuHantei = 1
        Else
            HairetuHantei = 2
        End If
    End If
    
    'クリップボードに格納用のテキスト変数を作成
    Dim Output$
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    If HairetuHantei = 0 Then '配列でない場合
        Output = InputClipText
    ElseIf HairetuHantei = 1 Then '1次元配列の場合
    
        If LBound(InputClipText, 1) <> 1 Then '最初の要素番号が1出ない場合は最初の要素番号を1にする
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        
        Output = ""
        For I = 1 To N
            If I = 1 Then
                Output = InputClipText(I)
            Else
                Output = Output & vbLf & InputClipText(I)
            End If
            
        Next I
    ElseIf HairetuHantei = 2 Then '2次元配列の場合
        
        If LBound(InputClipText, 1) <> 1 Or LBound(InputClipText, 2) <> 1 Then
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        M = UBound(InputClipText, 2)
        
        Output = ""
        
        For I = 1 To N
            For J = 1 To M
                If J < M Then
                    Output = Output & InputClipText(I, J) & Chr(9)
                Else
                    Output = Output & InputClipText(I, J)
                End If
                
            Next J
            
            If I < N Then
                Output = Output & Chr(10)
            End If
        Next I
    End If
    
    
    'クリップボードに格納'参考 https://www.ka-net.org/blog/?p=7537
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = Output
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    '格納したテキスト変数をメッセージ表示
    If MessageIrunaraTrue Then
        MsgBox ("「" & Output & "」" & vbLf & _
                "をクリップボードにコピーしました。")
    End If
    
End Sub


