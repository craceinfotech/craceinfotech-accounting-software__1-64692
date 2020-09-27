VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Profit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profit and Loss Account"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameLedger 
      Height          =   6915
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   10515
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   3990
         TabIndex        =   2
         Top             =   6540
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5670
         TabIndex        =   1
         Top             =   6540
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6195
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   10927
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   16000
         TextRTF         =   $"frmProfit.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Profit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineCount As Integer
Dim HeaderLength As Integer
Dim FooterLength As Integer
Dim DetailLength As Integer
Dim PageLength As Integer
Dim PageWidth As Integer
Dim PageCount As Integer
Dim MasterRS As Recordset
Dim TransactRS As Recordset
Dim TrialRS As Recordset
Dim ProfitRS As Recordset
Dim NumberOfRecords As Long
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewPage As Boolean


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    PrepareProfit
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\Profit.txt", rtfText
    cmdPrint.Enabled = True
    RichTextBox1.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub ProfitHeader()
    'IMPORTANT:48+2+14 + 4 + 48+2+14=132
    '1,51,65,69,119,132
    
    Print #1,
    Print #1, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageCount, "@@@@")
    Print #1, PADC("PROFIT AND LOSS ACCOUNT", PageWidth)
    Print #1, PADC("(For the Period from " & Format(FromDate, "dd/mm/yyyy") & " to " & Format(ToDate, "dd/mm/yyyy") & ")", PageWidth)
    'Print #1, PADC("As on " & Format(ToDate, "dd/mm/yyyy"), PageWidth)
    Print #1, String(PageWidth, "-")
    Print #1, "PARTICULARS OF EXPENSES"; Tab(51); Spc(8); "AMOUNT"; Tab(65); Spc(4); "PARTICULARS OF INCOME"; Tab(119); Spc(8); "AMOUNT"
    Print #1, String(PageWidth, "-")
    
    If PageCount > 1 Then
        Print #1, Tab(45); "b/f"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Tab(119); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1,
    End If
    
    LineCount = 0
    NewPage = True

'HeaderLength = 9
End Sub
Private Sub ProfitFooter()
    Print #1,
    Print #1, String(PageWidth, "-")
    Print #1, Tab(45); "c/d"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
    Print #1, Tab(119); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
    Print #1,
    'Print #1, "Page Length : "; LineCount + HeaderLength + FooterLength
    Print #1, Chr(12)
    'FooterLength = 5
    'If LineCount + HeaderLength + FooterLength > 67 Then MsgBox PageCount
    'If LineCount + HeaderLength + FooterLength < 65 Then MsgBox PageCount
End Sub
Private Sub ProfitDetail()
'1,51,69,119
Print #1, Mid(ProfitRS!ACTITLE1, 1, 48);
Print #1, Tab(51); ZeroSup(ProfitRS!head_bal1);
DebitBalance = DebitBalance + (ProfitRS!head_bal1)
Print #1, Tab(69); Mid(ProfitRS!ACTITLE2, 1, 48);
Print #1, Tab(119); ZeroSup(ProfitRS!head_bal2)
CreditBalance = CreditBalance + ProfitRS!head_bal2

LineCount = LineCount + 1
End Sub
Private Sub ProfitSummary()
        If NumberOfRecords > 0 Then
            Print #1, Tab(51); String(14, "-");
            Print #1, Tab(119); String(14, "-")
            Print #1, Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
            Print #1, Tab(119); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
            Print #1, Tab(51); String(14, "-");
            Print #1, Tab(119); String(14, "-")
        End If
        
        Print #1,
        If DebitBalance <> CreditBalance Then
        
        Print #1, Tab(23); "Difference";
        If CreditBalance - DebitBalance > 0 Then
            Print #1, Tab(67); Format(Format(CreditBalance - DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(DebitBalance - CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        End If
        End If
        Print #1,
        Print #1, String(PageWidth, "-")
        Print #1, Chr(12)
        'summaryLength = 10

End Sub

Private Sub PrepareProfit()
    Dim i As Long
    Dim CashTrans As Currency
    PageLength = 66
    PageWidth = 132
    HeaderLength = 9
    FooterLength = 5
    DetailLength = PageLength - (HeaderLength + FooterLength + 1)
    PageCount = 0
    LineCount = DetailLength + 1
    cmdPrint.Enabled = False
    CashTrans = 0
    DebitBalance = 0
    CreditBalance = 0
    i = 0
        
    Set MasterRS = New Recordset
    MasterRS.Open "select acnumber,actitle,balancetyp,balance,fnlrptcode,fnlrptposi FROM " & MasterTable & " order by actitle", _
        db, adOpenStatic, adLockReadOnly, adCmdText
    
    Set TrialRS = New Recordset
    TrialRS.Fields.Append "ACNUMBER", adInteger, , adFldKeyColumn
    TrialRS.Fields.Append "ACTITLE", adChar, 50
    TrialRS.Fields.Append "BALANCETYP", adChar, 1
    TrialRS.Fields.Append "BALANCE", adCurrency
    TrialRS.Fields.Append "FNLRPTCODE", adInteger
    TrialRS.Fields.Append "FNLRPTPOSI", adInteger
    TrialRS.Fields.Append "TRANSACT", adCurrency
    TrialRS.Fields.Append "HEAD_BAL", adCurrency
    TrialRS.Fields.Append "NO_NEED", adInteger
    
    TrialRS.CursorLocation = adUseClient
    TrialRS.CursorType = adOpenStatic
    TrialRS.LockType = adLockOptimistic
    TrialRS.Open
    
    With MasterRS
        .MoveFirst
        Do While Not .EOF
            TrialRS.AddNew
            TrialRS!ACNUMBER = MasterRS("acnumber").Value
            TrialRS!ACTITLE = Trim(MasterRS("actitle").Value)
            TrialRS!balancetyp = MasterRS("balancetyp").Value
            TrialRS!Balance = IIf(MasterRS("balancetyp").Value = "C", MasterRS("balance").Value, -1 * MasterRS("balance").Value)
            TrialRS!fnlrptcode = MasterRS("fnlrptcode").Value
            TrialRS!fnlrptposi = MasterRS("fnlrptposi").Value
            TrialRS!TRANSACT = 0
            TrialRS!head_bal = 0
            TrialRS!no_need = 0
            TrialRS.Update
            .MoveNext
        Loop
    End With

    
    Set TransactRS = New Recordset
    TransactRS.Open "select ACNUMBER,SUM(CREDIT-DEBIT) as TRANSACT FROM " & TransactionTable & "  where acn_date<={" & Format(ToDate, "mm/dd/yyyy") & "} group by ACNUMBER", _
        db, adOpenStatic, adLockReadOnly, adCmdText
    
    With TransactRS
    If .BOF = False And .EOF = False Then
        .MoveFirst
        Do While Not .EOF
            TrialRS.Find "acnumber = " & TransactRS("acnumber").Value, , adSearchForward, adBookmarkFirst
            If Not TrialRS.EOF Then
                TrialRS!TRANSACT = TransactRS("transact").Value
                TrialRS.Update
                CashTrans = CashTrans + TransactRS("transact").Value
            End If
            .MoveNext
        Loop
    End If
    End With
    
    With TrialRS
    If .BOF = False And .EOF = False Then
        .MoveFirst
        .Find "actitle='CASH'", , adSearchForward
        If Not .EOF Then
            TrialRS!TRANSACT = -1 * CashTrans
            .Update
        End If
        .MoveFirst
        Do While Not .EOF
                TrialRS!head_bal = TrialRS("transact").Value + TrialRS("balance").Value
                .Update
            .MoveNext
        Loop
    End If
    End With

    TrialRS.MoveFirst

    TrialRS.Filter = "head_bal <> 0 and fnlrptcode=1"
    'TrialRS.MoveFirst
    If TrialRS.EOF = False And TrialRS.BOF = False Then
        TrialRS.MoveFirst
    Else
        'no trading account
        'MsgBox "Not Sufficient Entries"
        'Unload Me
        'Exit Sub
    End If


    CashTrans = 0
    
    With TrialRS
        Do While Not .EOF
            CashTrans = CashTrans + TrialRS("head_bal").Value
            .MoveNext
        Loop
    End With
    
    CashTrans = CashTrans + StockInHand
    
    TrialRS.Filter = "head_bal <> 0 and fnlrptcode=2"
    TrialRS.MoveFirst
    
    
    
    Set ProfitRS = New Recordset
    ProfitRS.Fields.Append "SNUMBER", adInteger, , adFldKeyColumn
    ProfitRS.Fields.Append "ACTITLE1", adChar, 50
    ProfitRS.Fields.Append "HEAD_BAL1", adCurrency
    ProfitRS.Fields.Append "ACTITLE2", adChar, 50
    ProfitRS.Fields.Append "HEAD_BAL2", adCurrency

    
    ProfitRS.CursorLocation = adUseClient
    ProfitRS.CursorType = adOpenStatic
    ProfitRS.LockType = adLockOptimistic
    ProfitRS.Open
    i = 0
    DebitBalance = 0
    CreditBalance = 0
'*************
    If CashTrans > 0 Then
            ProfitRS.AddNew
            i = i + 1
            ProfitRS!SNUMBER = i
            ProfitRS!ACTITLE1 = Space(50)
            ProfitRS!head_bal1 = 0
            ProfitRS!ACTITLE2 = "GROSS PROFIT"
            ProfitRS!head_bal2 = CashTrans
            ProfitRS.Update
            CreditBalance = CreditBalance + CashTrans
    ElseIf CashTrans < 0 Then
            ProfitRS.AddNew
            i = i + 1
            ProfitRS!SNUMBER = i
            ProfitRS!ACTITLE1 = "GROSS LOSS"
            ProfitRS!head_bal1 = -1 * CashTrans
            ProfitRS!ACTITLE2 = Space(50)
            ProfitRS!head_bal2 = 0
            ProfitRS.Update
            DebitBalance = DebitBalance + (-1 * CashTrans)
    End If
    

    
    With TrialRS
        .MoveFirst
        Do While Not .EOF
            If TrialRS("head_bal").Value < 0 Then
            ProfitRS.AddNew
            i = i + 1
            ProfitRS!SNUMBER = i
            ProfitRS!ACTITLE1 = Trim(TrialRS("actitle").Value)
            ProfitRS!head_bal1 = -1 * TrialRS("HEAD_BAL").Value
            ProfitRS.Update
            DebitBalance = DebitBalance + (-1 * TrialRS("HEAD_BAL").Value)
            End If
            .MoveNext
        Loop
    End With
    
    ProfitRS.MoveFirst
    If CashTrans <> 0 Then
        ProfitRS.MoveNext
    End If
    
    With TrialRS
        .MoveFirst
        Do While Not .EOF
            If TrialRS("head_bal").Value > 0 Then
                ProfitRS!ACTITLE2 = Trim(TrialRS("actitle").Value)
                ProfitRS!head_bal2 = TrialRS("HEAD_BAL").Value
                ProfitRS.Update
                CreditBalance = CreditBalance + (TrialRS("HEAD_BAL").Value)
                ProfitRS.MoveNext
                If ProfitRS.EOF And TrialRS.EOF = False Then
                    ProfitRS.AddNew
                    i = i + 1
                    ProfitRS!SNUMBER = i
                    ProfitRS!ACTITLE1 = Space(50)
                    ProfitRS!head_bal1 = 0
                End If
            End If
            .MoveNext
        Loop
    End With

    
    If CreditBalance > DebitBalance Then
            ProfitRS.AddNew
            i = i + 1
            ProfitRS!SNUMBER = i
            ProfitRS!ACTITLE1 = "NET PROFIT"
            ProfitRS!head_bal1 = CreditBalance - DebitBalance
            ProfitRS!ACTITLE2 = Space(50)
            ProfitRS!head_bal2 = 0
            ProfitRS.Update
    ElseIf CreditBalance < DebitBalance Then
            ProfitRS.AddNew
            i = i + 1
            ProfitRS!SNUMBER = i
            ProfitRS!ACTITLE2 = "NET LOSS"
            ProfitRS!head_bal2 = DebitBalance - CreditBalance
            ProfitRS!ACTITLE1 = Space(50)
            ProfitRS!head_bal1 = 0
            ProfitRS.Update
    End If
    

    
    
    NumberOfRecords = ProfitRS.RecordCount
    i = 0
    LineCount = DetailLength + 1
    DebitBalance = 0
    CreditBalance = 0
        
        If NumberOfRecords > 0 Then
            ProfitRS.MoveFirst
        End If
        

    
    Open "c:\vbprog\vba\rpt\Profit.txt" For Output As #1
   
    Print #1,
    'Chr(27) + Chr(15)
    Do While i < NumberOfRecords
        
        If LineCount > DetailLength Then
                If PageCount > 0 Then
                    ProfitFooter
                End If
                PageCount = PageCount + 1
                ProfitHeader
        End If

        If LineCount < DetailLength Then
            With ProfitRS
                If Not .EOF Then
                    ProfitDetail
                    .MoveNext
                    i = i + 1
                End If
            End With
        Else
            LineCount = DetailLength + 1
        End If

    Loop
    

        If DetailLength - LineCount > 10 Then
            ProfitSummary
        Else
            ProfitFooter
            ProfitHeader
            ProfitSummary
        End If
        
    Print #1,
    'Chr(27) + Chr(18)
    Close #1
    
End Sub

Private Sub cmdprint_Click()
    cmdPrint.Enabled = False
    PrintText (RichTextBox1.Text)
    ContinueProcess = True
End Sub

