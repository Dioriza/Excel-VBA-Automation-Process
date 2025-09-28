Sub Check_Error_Level_Consignment_Part()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim LV1 As Variant, LV2 As Variant
    Dim J1 As String, J2 As String
    Dim FP1 As String, FP2 As String
    Dim P1 As String, P2 As String
    Dim answer As VbMsgBoxResult
    Dim found As Boolean
    
    ' LV1/LV2 = cek urutan sequence level di kolom D
    ' J1/J2 = pengecualian kondisi color part (K & C) pada kolom J
    ' FP1/FP2 = cek apakah kolom O memiliki kode yang sama (following process supplier code)
    ' P1/P2 = abaikan jika kosong (hanya membaca consignment part)
    ' answer = respon user di popup (output response)
    ' found = flag apakah ada error yang ditemukan (output response)
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    found = False
    
    For i = 2 To lastRow - 1
        LV1 = ws.Cells(i, "D").Value
        LV2 = ws.Cells(i + 1, "D").Value
        J1 = Trim(CStr(ws.Cells(i, "J").Value))
        J2 = Trim(CStr(ws.Cells(i + 1, "J").Value))
        FP1 = Trim(CStr(ws.Cells(i, "O").Value))
        FP2 = Trim(CStr(ws.Cells(i + 1, "O").Value))
        P1 = Trim(CStr(ws.Cells(i, "P").Value))
        P2 = Trim(CStr(ws.Cells(i + 1, "P").Value))
        
        ' Abaikan jika kolom O atau P kosong
        If Len(FP1) = 0 Or Len(FP2) = 0 Or Len(P1) = 0 Or Len(P2) = 0 Then GoTo NextLoop
        
        ' ==========================
        ' Rule 1: cek urutan sequence level numeric +1
        ' ==========================
        If IsNumeric(LV1) And IsNumeric(LV2) Then
            If LV2 = LV1 + 1 Then
                If FP1 = FP2 Then
                    ' Abaikan jika J1=K dan J2=C == color part
                    If Not (J1 = "K" And J2 = "C") Then
                        found = True
                        
                        ' Highlight baris error
                        ws.Rows(i).Interior.Color = vbYellow
                        ws.Rows(i + 1).Interior.Color = vbYellow
                        
                        answer = MsgBox("Error ditemukan di baris " & i & " dan " & (i + 1) & vbCrLf & _
                                        "LV sequence (" & LV1 & " -> " & LV2 & ") dengan supplier code sama (" & FP1 & ")" & vbCrLf & vbCrLf & _
                                        "Klik Yes untuk lompat ke baris " & i, _
                                        vbExclamation + vbYesNo, "Error Following Process Consignment part")
                        If answer = vbYes Then
                            ws.Activate
                            ws.Cells(i, "D").Select
                        End If
                    End If
                End If
            End If
        End If
        
        ' ==========================
        ' Rule 2: LV sama + FP sama + J berbeda (Color part vs Non-color)
        ' ==========================
        If LV1 = LV2 And FP1 = FP2 Then
            If (J1 = "C" And J2 <> "C") Or (J2 = "C" And J1 <> "C") Then
                found = True
                
                ' Highlight baris error
                ws.Rows(i).Interior.Color = vbYellow
                ws.Rows(i + 1).Interior.Color = vbYellow
                
                answer = MsgBox("Error ditemukan di baris " & i & " dan " & (i + 1) & vbCrLf & _
                                "Color part vs non-color (" & LV1 & "), Supplier code sama (" & FP1 & ")" & vbCrLf & vbCrLf & _
                                "Klik Yes untuk lompat ke baris " & i, _
                                vbExclamation + vbYesNo, "Error Following Process Consignment part")
                If answer = vbYes Then
                    ws.Activate
                    ws.Cells(i, "D").Select
                End If
            End If
        End If
        
NextLoop:
    Next i
    
    If Not found Then
        MsgBox "Tidak ditemukan error.", vbInformation
    Else
        MsgBox "Pengecekan selesai.", vbInformation
    End If
End Sub


