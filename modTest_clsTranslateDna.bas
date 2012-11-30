Attribute VB_Name = "modTest_clsTranslateDna"
Option Explicit
'To test code in class "clsTranslateDna".
Sub Run_Test_clsTranslateDna()
    Dim dnaSeq As String
    'See link for source sequence:
    'http://www.hiv.lanl.gov/components/sequence/HIV/asearch/query_one.comp?se_id=AB025113
    'Using HIV-1 protease sequence as test.
    dnaSeq = "cctcagatcactctttggcaacgacccctcgtcacaataaggataggggggcagctaaaggaagctctattagatacaggagcagatgatacagtattagaagaaatgaatttgccaggaagatggaaaccaaaaatgatagggggaattggaggttttatcaaagtaagacagtatgatcagatacccatagaaatcagtggaaagaaagctataggtacagtattaataggacctacacctgtcaacataattggaagaaatctgttgactcagattggctgcactttaaatttt"
    Call Test_clsTranslateDna(dnaSeq)
End Sub
'Output is compared to translation given on page referenced above.
'Debug.Assert will throw an error if the translation performed here does not equal the expected result.
Sub Test_clsTranslateDna(dnaSeq As String)
    If Len(dnaSeq) < 3 Then
        Debug.Print "Input sequence " & "'" & "'" & " < one codon in length. Exiting......"
        Exit Sub
    End If
    Dim translator As clsTranslateDna
    Set translator = New clsTranslateDna
    Dim arrAA() As String
    Dim i As Long
    Dim givenTranslation As String
    'Translation given on link above, take this as "correct"
    givenTranslation = "PQITLWQRPLVTIRIGGQLKEALLDTGADDTVLEEMNLPGRWKPKMIGGIGGFIKVRQYDQIPIEISGKKAIGTVLIGPTPVNIIGRNLLTQIGCTLNF"
    
    'Translate ignoring IUPAC ambiguity codes.
    arrAA = translator.AminoAcidsForDNA(dnaSeq, False)
    Debug.Assert (Join(arrAA, "") = givenTranslation)
    Debug.Print "Output WITHOUT IUPAC ambiguity Code Translation"
    For i = 0 To UBound(arrAA)
        Debug.Print i + 1, ": ", arrAA(i)
    Next i
    
    'Translate taking IUPAC ambiguity codes into account.
    Debug.Print "Output WITH IUPAC ambiguity Code Translation"
    arrAA = translator.AminoAcidsForDNA(dnaSeq, True)
    Debug.Assert (Join(arrAA, "") = givenTranslation)
    Debug.Print Join(arrAA, "")
    For i = 0 To UBound(arrAA)
        Debug.Print i + 1, ": ", arrAA(i)
    Next i
   
   'Precautionary cleanup
    Set translator = Nothing
    
End Sub

