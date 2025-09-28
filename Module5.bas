Sub LoadCustRibbon()

    Dim hFile As Long
    Dim path As String, fileName As String, ribbonXML As String, user As String
    
    hFile = FreeFile
    user = Environ("Username")
    path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
    fileName = "Excel.officeUI"
    
    ribbonXML = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
    ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='Data Preparation' insertBeforeQ='mso:TabFormat'>" & vbNewLine
    
    ' Automate Process Group
    ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='Automate Process' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='PMCWeeklyKD' label='PMC Weekly for KD'" & vbNewLine
    ribbonXML = ribbonXML + "            imageMso='ShapeFillTextureGallery' onAction='PMCWeeklyKD' screentip='PMC Weekly Data Preparation for KD' supertip='This function is to wrap up data analysis steps into one click. How to operate it is Clear all filter, Move the cursor to column A1 then click function. &#10;" & _
                  "&#10;" & _
                  "Below are the master steps data of this function: &#10;" & _
                  "1. Filter Column M untick A, Column E untick L  -> Delete &#10;" & _
                  "2. Filter Column E tick B -> Delete &#10;" & _
                  "3. Filter Column J tick C -> Delete &#10;" & _
                  "4. Filter Column A tick AAA -> Delete &#10;" & _
                  "5. Filter Column A tick M69 -> Delete &#10;" & _
                  "6. Delete columns B, C, H, I, J, K &#10;" & _
                  "&#10;" & _
                  "FYI: After the function is executed, it cannot be undone (Ctrl + Z)" & _
                  "&#10;" & _
                  "&#10;" & _
                  "Not all columns already deleted, please check again' />" & vbNewLine
    
    ribbonXML = ribbonXML + "          <mso:button id='Automate_PMC_Weekly' label='PMC Weekly for CBU'" & vbNewLine
    ribbonXML = ribbonXML + "            imageMso='ShapeFillTextureGallery' onAction='Automate_PMC_Weekly' screentip='PMC Weekly Data Preparation for CBU' supertip='This function is to wrap up data analysis steps into one click. How to operate it is Clear all filter, Move the cursor to column A1 then click function. &#10;" & _
                  "&#10;" & _
                  "Below are the master steps data of this function: &#10;" & _
                  "1. Filter Column L untick 1, Column M untick A -> Delete &#10;" & _
                  "2. Filter Column J tick C -> Delete &#10;" & _
                  "3. Filter Column A tick AAA -> Delete &#10;" & _
                  "4. Filter Column A tick M69 -> Delete &#10;" & _
                  "5. Delete columns B, C, E, F, H, I, J, K, T, U &#10;" & _
                  "&#10;" & _
                  "FYI: After the function is executed, it cannot be undone (Ctrl + Z)' />" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='SumAllValueFields' label='Pivot Sum All Fields'" & vbNewLine
    ribbonXML = ribbonXML + "            imageMso='DatabasePartialReplica' onAction='SumAllValueFields' screentip='Pivot Sum All Fields' supertip='Functions to change all columns containing COUNT values in Pivot into SUM  &#10;" & _
                  "&#10;" & _
                  "' />" & vbNewLine
    ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine

    ' Shortcut Way Group
    ribbonXML = ribbonXML + "        <mso:group id='shortcutGroup' label='Format Number' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='FormatToTwoDigits' label='Convert To Two Digits'" & vbNewLine
    ribbonXML = ribbonXML + "            imageMso='ReviewCompareMajorVersion' onAction='FormatToTwoDigits' screentip='Convert Format Number' supertip='Functions to change digit format, exp: 1 to 01  &#10;" & _
                  "&#10;" & _
                  "' />" & vbNewLine
    ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
    
        ' New Formula Reference Group
    ribbonXML = ribbonXML & "      <mso:group id='newFormulaGroup' label='New Formula Reference' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='ExtractValue' label='Extract Value'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='PivotTableListFormulas' onAction='ShowFILLEDVALUE' screentip='FILLEDVALUE' supertip='Mengekstrak value pada range cell, hanya mengembalikan satu nilai. Formula: =FILLEDVALUE(RANGE)' />" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='MakeBuyCode' label='Generate MB Code'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='PivotTableListFormulas' onAction='ShowMBGENERATE' screentip='Make or Buy Generate' supertip='Mengekstrak Code M/B sesuai dengan Supplier yang ada pada Sourcing List. Formula: =MBGENERATE(Supplier_name_cell)' />" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='FollowingProcessCode' label='Generate FP Code'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='PivotTableListFormulas' onAction='ShowFPGENERATE' screentip='Following Process Generate' supertip='Mengekstrak Code FP sesuai dengan Supply Line yang ada pada Sourcing List.                       Formula: =FPGENERATE(Supply_line_cell)' />" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='ExtractAllRangeValue' label='Extract All Range Value'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='PivotTableListFormulas' onAction='ShowEXTRACTRANGE' screentip='EXTRACTRANGE' supertip='Mengekstrak value pada range cell, mengembalikan semua nilai unique. Formula: =EXRACTRANGE(RANGE)' />" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='ExtractAllRangeValu' label='Compare Multiple Range'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='PivotTableListFormulas' onAction='ShowRANGECOMPARE' screentip='RANGECOMPARE' supertip='Compare antar range, menghasilkan output TRUE/FALSE. Contoh penggunaan: Compare A10:C10 dengan F10:H10.      Formula: =RANGECOMPARE(RANGE1, RANGE2,.. Dst)' />" & vbNewLine
    ribbonXML = ribbonXML & "      </mso:group>" & vbNewLine
    
            ' New Formula Reference Group
    ribbonXML = ribbonXML & "      <mso:group id='newFunctionGroup' label='Data Checking' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML & "        <mso:button id='Check_Error_Level_Consignment_Part' label='Consignment Error Check'" & vbNewLine
    ribbonXML = ribbonXML & "          imageMso='TraceRemoveAllArrows' onAction='Check_Error_Level_Consignment_Part' screentip='LV-FP CONSIGNMENT' supertip='Mengecek kondisi error berdasarkan levelling dan following process di consignment part' />" & vbNewLine
    ribbonXML = ribbonXML & "      </mso:group>" & vbNewLine
    

    ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
    ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "</mso:customUI>"
    
    ribbonXML = Replace(ribbonXML, """", "")
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile
    
End Sub
