Sub Surf()
' WRV
Static flechr As Integer
Static rangoi As Range
Static rangoj As Range
Static nconti As Integer
Static nconta As Integer
    
    Application.ScreenUpdating = False
    
    Select Case flechr
            Case 0

'---------------------------------------------------------------------------------
'                   Primera ejecución de la macro
'---------------------------------------------------------------------------------
Prima:
                    If rangoi Is Nothing Then
                    Else
                        If rangoi <> ActiveCell And nconti >= 0 And nconta >= 0 Then
                            nconti = 0
                            nconta = 0
                        End If
                    End If
                    
                        Set rangoi = ActiveCell 'Graba como la celda activa
                        conti = 1                'Me sirve para contar los vinculos
                        conta = 1                'Me sirve para contar los vinculos

'---------------------------------------------------------------------------------
'Trata de identificar si se esta ejecutando por segunda continua la macro en la misma celda
'---------------------------------------------------------------------------------
            'If rangoi.Address(external:=True) = ActiveCell.Address(external:=True) Then
            '        GoTo Vinculo
            'Else
'---------------------------------------------------------------------------------
'                   'Hallar el número de vinculos precedentes
'---------------------------------------------------------------------------------
            If nconti = 0 And nconta = 0 Then
            Do
                Do
                ActiveCell.ShowPrecedents   'Muestra los vínculos precedentes
                On Error Resume Next
                ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=conta, LinkNumber:=conti
                If Err.Number > 0 Then
                    nconti = conti - 1
                    rangoi.Parent.ClearArrows
                    GoTo Vinculo
                End If
                On Error GoTo 0
                    If rangoi.Address(external:=True) = ActiveCell.Address(external:=True) Then
                        rangoi.Parent.ClearArrows
                        nconta = conta - 1
                        If nconta = 0 Then
                            Exit Sub
                        End If
                        Exit Do
                    End If
                Application.GoTo rangoi
                If conti > 1 Then Exit Do
                conta = conta + 1
                Loop
                conta = 1
                conti = conti + 1
            Loop
            End If
'---------------------------------------------------------------------------------
'                   Llevar a la primera ubicación
'---------------------------------------------------------------------------------
            Case Else
            If rangoj = ActiveCell Then
                Application.GoTo rangoi
                flechr = 0
                Exit Sub
            Else
                nconti = 0
                nconta = 0
                GoTo Prima
            End If
    End Select
    
Vinculo:

flechr = flechr + 1
            ActiveCell.ShowPrecedents
            If nconta = 0 Then
            Exit Sub
            ElseIf nconti = 1 Or nconta > 1 Then
                ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=nconta, LinkNumber:=1
                Set rangoj = ActiveCell
                nconta = nconta - 1
                rangoi.Parent.ClearArrows
                    If nconta = 0 And nconti = 1 Then
                        nconti = 0
                    End If
                Exit Sub
            End If
            
            If nconti = 0 Then
            Exit Sub
            ElseIf nconta = 1 Then
                ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=1, LinkNumber:=nconti
                Set rangoj = ActiveCell
                nconti = nconti - 1
                nconta = 1
                rangoi.Parent.ClearArrows
                    If nconti = 0 Then
                        nconta = 0
                    End If
            End If
    Application.ScreenUpdating = True
End Sub
