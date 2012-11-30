VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTranslateDna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' SUMMARY:
'A class for translating DNA sequences into protein.
'It can translate amino acid mixtures that are specified in the DNA sequence as IUPAC ambiguity codes.
'This makes it suitable for translating highly polymorphic DNA, e.g. viral sequences such as those from Hepatitis C or HIV as well as mammal MHC sequences.
'Has just one public method "AminoAcidsForDNA" that returns a string array with elements containing one or more (if IUPAC ambiguity coding is used) amino
'acid single letter codes.
'Has one public property "version" that returns the code version.
'ASSUMES:
'The given DNA sequence is in-frame.
'CODE NOTES:
'Uses the class "clsKeyValueMap" as a simple dictionary.  This class was used because development was done on the Mac
'where Excel does not have the dictionary object.  Code for that class is given in this repository and described there.
'Translation code that ignores IUPAC ambiguity codes
'EXAMPLE USAGE:
'See test code in "modTest_clsTranslateDna.bas"
'TO DO:
'On Windows, the hand-rolled "clsKeyValueMap" could/should be replaced by the Microsoft Dictionary object.
'TO USE:
'Import this class and "clsKeyValueMap" into the VBA project and run the test code to ensure everything is ok.
'Nothing in this class is specific to any application that hosts VBA.  It was developed in Excel but
'it SHOULD run in MS Word, MS Access, etc.  But it has only been tested in Excel.
'AUTHOR:
'mick@javascript-spreadsheet-programming.com

Private dictCodonToAA  As clsKeyValueMap 'Maps three letter codons to single letter amino acid codes.
Private dictIupacAmbig As clsKeyValueMap 'Maps nucleotide ambiguity codes to the nucleotides the represent.
Private Const UNKNOWN_AMINOACID As String = "X" 'Character to be returned when a codon cannot be mapped to an amino acid.
Private Const CODEVERSION As String = "Test"
'Class constructor.
'Sets two instance variables of class "clsKeyValueMap" and calls two methods to populate these instances
'One instance maps codons to single-letter amino acid codes (the translation table).
'The second maps IUPAC ambiguity codes to the nucleotide mixtures that they specify.
Private Sub class_initialize()
    Set dictCodonToAA = New clsKeyValueMap
    Call PopulateDictCodonToAA
    Set dictIupacAmbig = New clsKeyValueMap
    Call PopulateDictIupacAmbig
End Sub
'A flag to providing the code version of the class
Public Property Get Version() As String
    Version = CODEVERSION
End Property
'Populate the translation table using values specified in an array.
'Called by constructor
'Each element of the local array contains a codon-single letter amino acid code separated by a comma.
'The code processes each element by using the comma to split them into a two element array.
'The first element is the codon, the second the single letter amino acid code.
'The paired values are used to populate the "clsKeyValueMap" instance.
Private Sub PopulateDictCodonToAA()
    Dim i As Long
    Dim arrPair() As String
    Dim codonsToAminoAcids As Variant
    
    codonsToAminoAcids = Array("TAA,*", "TAG,*", "TGA,*", _
                                                "GCA,A", "GCC,A", "GCG,A", "GCT,A", _
                                                "TGC,C", "TGT,C", _
                                                "GAC,D", "GAT,D", _
                                                "GAA,E", "GAG,E", _
                                                "TTC,F", "TTT,F", _
                                                "GGA,G", "GGC,G", "GGG,G", "GGT,G", _
                                                "CAC,H", "CAT,H", _
                                                "ATA,I", "ATC,I", "ATT,I", _
                                                "AAA,K", "AAG,K", _
                                                "CTA,L", "CTC,L", "CTG,L", "CTT,L", "TTA,L", "TTG,L", _
                                                "ATG,M", _
                                                "AAC,N", "AAT,N", _
                                                "CCA,P", "CCC,P", "CCG,P", "CCT,P", _
                                                "CAA,Q", "CAG,Q", _
                                                "AGA,R", "AGG,R", "CGA,R", "CGC,R", "CGG,R", "CGT,R", _
                                                "AGC,S", "AGT,S", "TCA,S", "TCC,S", "TCG,S", "TCT,S", _
                                                "ACA,T", "ACC,T", "ACG,T", "ACT,T", _
                                                "GTA,V", "GTC,V", "GTG,V", "GTT,V", _
                                                "TGG,W", _
                                                "TAC,Y", "TAT,Y")
    
    For i = 0 To UBound(codonsToAminoAcids)
        arrPair = Split(codonsToAminoAcids(i), ",")
        Call dictCodonToAA.addKeyValuePair(arrPair(0), arrPair(1))
    Next i
    
End Sub
'Populate the lookup table mapping IUPAC ambiguity codes to a string of the nucleotide mixtures they represent.
'Called by constructor.
'A local array stores the mappings as strings in the format "nucleotide-comma-ambiguity code"
'Each element is processed by splitting on the comma to produce a two-element string array.
'The first element is the key, the second the value of the lookup table (instance of of class "clsKeyValueMap")
'populated in the "for" loop.
'NOTE:Only IUPAC ambiguity codes specifying two nucleotides are included.  The others could be added but
'have been omitted here to ensure that the amino acid counts at ambiguous positions does not become over-large.
Private Sub PopulateDictIupacAmbig()
    Dim arrAmbigsNucs As Variant
    Dim i As Long
    Dim arrPair() As String
    
    arrAmbigsNucs = Array("R,AG", "Y,CT", "K,GT", "M,AC", "S,CG", "W,AT", "A,A", "C,C", "G,G", "T,T")
    
    For i = 0 To UBound(arrAmbigsNucs)
        arrPair = Split(arrAmbigsNucs(i), ",")
        Call dictIupacAmbig.addKeyValuePair(arrPair(0), arrPair(1))
    Next i
    
End Sub
'Only public method of class.
'Performs the translation of given DNA sequence by calling one of two other private methods.
'If its second argument is "True" it calls method "TranslateDnaAmbig" to do a translation that handles IUPAC ambiguity codes.
'Otherwise, it does a simple translation by calling "TranslateDNANonAmbig" that returns the value of the constant
'"UNKNOWN_AMINOACID" for any codon containing any characters other than A,C,G,T.
'Return a string array where each element is an amino acid, the value of "UNKNOWN_AMINOACID",
'or a mixture of amino acids (if second argument = "True")
'All input sequence is converted to upper case before processing.
Public Function AminoAcidsForDNA(dnaSeq As String, iupacAmbig As Boolean) As String()
    dnaSeq = UCase(dnaSeq)
    If iupacAmbig Then
        AminoAcidsForDNA = TranslateDnaAmbig(dnaSeq)
    Else
        AminoAcidsForDNA = TranslateDNANonAmbig(dnaSeq)
    End If
End Function

'Simple translation of DNA sequence.
'Called by method "AminoAcidsForDNA" when this method's second argument is "False".
'Process input sequence by retrieving consecutive codons and return a string array of the amino acid codes from the translation table.
'The translation table returns a VB null string if the codon is not found, this is converted to the value in constant "UNKNOWN_AMINOACID".
Private Function TranslateDNANonAmbig(dnaSeq As String) As String()
    Dim arrAminoAcids() As String
    Dim seqLen As Long: seqLen = Len(dnaSeq)
    Dim i As Long: i = 1
    Dim elementCounter As Long: elementCounter = 0
    Dim codon As String
    Dim aminoAcid As String
    
    For i = 1 To seqLen Step 3
        codon = Mid(dnaSeq, i, 3)
        aminoAcid = dictCodonToAA.ValueForKey(codon)
        If aminoAcid = vbNullString Then
            aminoAcid = UNKNOWN_AMINOACID
        End If
        ReDim Preserve arrAminoAcids(elementCounter)
        arrAminoAcids(elementCounter) = aminoAcid
        elementCounter = elementCounter + 1
    Next i
    
    TranslateDNANonAmbig = arrAminoAcids
    
End Function
'Translates DNA so that IUPAC nucleotide ambiguity codes are recognised.
'Called by method "AminoAcidsForDNA" when this method's second argument is "True".
'Process each codon in a loop:
'  If the codon is in translation table, retrieve the associated amino acid code.
'  Else pass the codon to method "GetMixedCodons" with the mixed codon as its argument.
'  Pass the returned array of codons to another method, "GetMixedAminoAcids" and join its returned array with an empty string.
'  Push the amino acid code(s) onto a string array and return this array when the loop has completed.
Private Function TranslateDnaAmbig(dnaSeq As String) As String()
    Dim arrAminoAcids() As String
    Dim seqLen As Long: seqLen = Len(dnaSeq)
    Dim i As Long: i = 1
    Dim j As Long: j = 0
    Dim elementCounter As Long: elementCounter = 0
    Dim codon As String
    Dim arrMixedCodons() As String
    Dim arrMixedAminoAcids() As String
    Dim aminoAcid As String
    
    For i = 1 To seqLen Step 3
        codon = Mid(dnaSeq, i, 3)
        aminoAcid = dictCodonToAA.ValueForKey(codon)
        If aminoAcid = vbNullString Then
            arrMixedCodons = GetMixedCodons(codon)
            arrMixedAminoAcids = GetMixedAminoAcids(arrMixedCodons)
            aminoAcid = Join(arrMixedAminoAcids, "")
        End If
        ReDim Preserve arrAminoAcids(elementCounter)
        arrAminoAcids(elementCounter) = aminoAcid
        elementCounter = elementCounter + 1
    Next i
    
    TranslateDnaAmbig = arrAminoAcids
    
End Function
'For any codon with nucleotide IUPAC ambiguity codes, return a string array containing all the potential codons present.
'If the given codon contains any character not specifed in the lookup table keys (instance variable "dictIupacAmbig")
'return a single element array containing the constant "UNKNOWN_AMINOACID" repeated three times (the "If .. Or..").
'This is a tricker method than the others so there is in-line comment.
Private Function GetMixedCodons(codon As String) As String()
    Dim arrMixedCodons() As String
    Dim codonCount As Long: codonCount = 0
    'Extract each letter from the given codon.
    Dim mix1 As String: mix1 = Mid(codon, 1, 1)
    Dim mix2 As String: mix2 = Mid(codon, 2, 1)
    Dim mix3 As String: mix3 = Mid(codon, 3, 1)
    Dim i As Long: i = 1
    Dim j As Long: j = 1
    Dim k As Long: k = 1
    Dim nuc1 As String
    Dim nuc2 As String
    Dim nuc3 As String
    
    'Check if there are any unrecognised letters, if there are, then exit the function
    'returning a "dummy" codon that wiil resolve to "UNKNOWN_AMINOACID" in the translation.
    If dictIupacAmbig.ValueForKey(mix1) = vbNullString Or _
            dictIupacAmbig.ValueForKey(mix2) = vbNullString Or _
            dictIupacAmbig.ValueForKey(mix3) = vbNullString Then
        ReDim Preserve arrMixedCodons(0)
        arrMixedCodons(0) = UNKNOWN_AMINOACID & UNKNOWN_AMINOACID & UNKNOWN_AMINOACID
        GetMixedCodons = arrMixedCodons
        Exit Function
    End If
    
    'Worst bit of code!
    'Process characters 1, 2, and 3 from the given codon by looping one character at a time over the string values they return from
    '"dictIupacAmbig".
    'Assemble codons that the mixed input codons represent in the inner-most "for" loop.
    'Push each codon onto an array and ,once the loops have completed, return the array.
    For i = 1 To Len(dictIupacAmbig.ValueForKey(mix1))
        nuc1 = Mid(dictIupacAmbig.ValueForKey(mix1), i, 1)
        For j = 1 To Len(dictIupacAmbig.ValueForKey(mix2))
            nuc2 = Mid(dictIupacAmbig.ValueForKey(mix2), j, 1)
            For k = 1 To Len(dictIupacAmbig.ValueForKey(mix3))
                nuc3 = Mid(dictIupacAmbig.ValueForKey(mix3), k, 1)
                codon = nuc1 & nuc2 & nuc3
                ReDim Preserve arrMixedCodons(codonCount)
                arrMixedCodons(codonCount) = codon
                codonCount = codonCount + 1
            Next k
        Next j
    Next i
    
    GetMixedCodons = arrMixedCodons
    
End Function
'Given an array of codons wheere all IUPAC ambiguities have been resolved,
'return a string array containing the codes for the amino acids that they encode.
'The amino acid translations are added as keys to an instance of "clsKeyValueMap".
'This is done to prevent duplicates appearing in the returned array.
'Duplicates arise due to the redundancy of the genetic code where multiple codons specify the same amino acid.
'Example "CAA" and "CAG" both specify "Q" so a "CAR" would resolve to "CAA" and "CAG" giving two "Q"'s
'in the output.  Making them lookup keys that map to a dummy value ("1" here") suppresses these duplicates.
Private Function GetMixedAminoAcids(arrMixedCodons() As String) As String()
    Dim i As Long: i = 0
    Dim dictAminoAcid As clsKeyValueMap
    Dim aminoAcid As String
    Dim arrMixedAminoAcids() As String
    
    
    Set dictAminoAcid = New clsKeyValueMap
    
    For i = 0 To UBound(arrMixedCodons)
        aminoAcid = dictCodonToAA.ValueForKey(arrMixedCodons(i))
        If aminoAcid = vbNullString Then
            aminoAcid = UNKNOWN_AMINOACID
        End If
        Call dictAminoAcid.addKeyValuePair(aminoAcid, "1")
    Next i
    
    GetMixedAminoAcids = dictAminoAcid.Keys

    Set dictAminoAcid = Nothing
    
End Function

