Imports System.Data.SqlClient
Imports System.Data
Imports System
Imports System.IO
Imports AutSessTypeLibrary
Imports AutPSTypeLibrary
Imports AutOIATypeLibrary
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data.DataTable
Imports System.Collections.Generic
Imports CommonToolsv2


'Macro for ER 707 - Developed b Swati Kulkarni
'07/06/2016 - Macro looks for units greater than allowed units and contracted by % billed
'10/26/2016 - Adding CommonTools.dll for username/password and PComm Emulator
'05/14/2018 - Removed 00603 Review Code Exclusion (UBH Claims identified by Legal Entity U) - DD
'01/06/2020 - Transition on MARCO - RM
'04/01/2020 - Added 1033 review code to the 707 review codes that needs to be removed - SK
'04/17/2020 - US161247 - Added delegated logic - SK
'05/14/2020 - - US164257 - Updating scrape to remove COVID claims from normal edits - SK

Module Module1
    Dim ps As New AutPS
    Dim oia As New AutOIA
    Dim ComT As CT = New CT()
    Dim proc_list = New List(Of procs)()
    Dim DLMGroupList = New List(Of DelGroups)()
    Dim CovidClmsList = New List(Of CovidClms)()
    Dim CovidProcs = New List(Of String)()
    Dim CovidMods = New List(Of String)()

    Sub Main()

#If (DEBUG) Then
        'setting the connection string to the test database mirroring the SAM server
        Const connection As String = "Data Source=APSEP04828;Initial Catalog=Macro;Integrated Security=True"
        Dim DataTable As String = "SAM_MACRO_INTERFACE_COSMOS"
#Else
        'setting the connection string to the SAM server for production
        Const connection As String = "Data Source=DBSEP2093;Initial Catalog=COSMOS_GRADING;Integrated Security=True"
        Dim DataTable As String = "SAM_MACRO_INTERFACE"
#End If

        Dim conn As New SqlConnection(connection)
        Dim ERdataSet As New DataSet()
        Dim da As SqlDataAdapter
        Dim dRow As DataRow
        Dim cmdBuilder As SqlCommandBuilder
        Dim x = 0
        Dim notes2 As String = ""
        Dim emu As String = ComT.emuSelect()
        Dim TotalClaims = 0, cnt = 1

        'Connect to PCOMM emulator 
        ps.SetConnectionByName(emu)
        oia.SetConnectionByName(emu)
        ApiChk()
        ps.StopCommunication()
        ps.Wait(1000)
        ps.StartCommunication()
        ps.Wait(1000)

        'Set the connection string of the SqlConnection object to connect to the
        'SQL Server database in which you created the sample table.
        Try
            conn.Open()

            'Initialize the SqlDataAdapter object by specifying a Select command 
            'that retrieves data from the sample table.  sam_MACRO_Interface

            da = New SqlDataAdapter("Select * from " & DataTable & " " & _
                                    "Where EDIT_RULE_NO in (707, 1012, 1013)" & _
                                    "and CURRENT_STATUS = 0", conn)

            'Initialize the SqlCommandBuilder object to automatically generate and initialize
            'the UpdateCommand, InsertCommand and DeleteCommand properties of the SqlDataAdapter.
            GetProc_list()

            'Populate the list of Delegated groups for later use
            PopulateDelGroups()

            'Populate Covid Procs And Modifiers for later use
            PopulateCovidProcsMods()

            'Populate Covid Claims flagged by the exclusion criteria
            PopulateCovidClms()

            cmdBuilder = New SqlCommandBuilder(da)
            'Populate the dataset by running the Fill method of the SqlDataAdapter.
            da.Fill(ERdataSet, DataTable)
            'To track the scrape progress
            TotalClaims = ERdataSet.Tables(DataTable).Rows.Count

            'conn.Close()
            For Each dRow In ERdataSet.Tables(DataTable).Rows
                Dim icn As String
                Dim claim As String
                icn = Left(dRow.Item("CLAIM_NO"), 13)
                icn = RTrim(icn)
                Dim PreviousNotes As String
                PreviousNotes = dRow.Item("MACRO_NOTES").ToString
                'pulls on digits needed for Hosp (9) and Phys claims (10)
                If Right(icn, 1) = " " Then
                    icn = Left(dRow.Item("CLAIM_NO"), 12)
                    claim = Right(icn, 9)
                Else
                    claim = Right(icn, 9)
                End If


                Dim div As String
                'Pulling in only the div designation from the icn field
                div = Left(icn, 3)
                Dim curStatus As String
                Dim result As String = 0
                Dim notes As String = ""
                Dim dtvar As DateTime
                dtvar = Date.Now
                'To display the Scrape progress
                'Console.Write(cnt & " of " & TotalClaims)

                runER(result, div, claim, notes, notes2)

                curStatus = "1"
                If notes = "password expired" Then
                    Exit Sub
                End If

                If result <> "0" Then

                    'Modify the value of the following fields
                    ERdataSet.Tables(DataTable).Rows(x)("CURRENT_STATUS") = curStatus
                    ERdataSet.Tables(DataTable).Rows(x)("MACRO_RESULT") = result
                    ERdataSet.Tables(DataTable).Rows(x)("MACRO_NOTES") = notes2
                    ERdataSet.Tables(DataTable).Rows(x)("DATE_UPDATED") = dtvar
                    ERdataSet.Tables(DataTable).Rows(x)("ATTRIBUTE_5") = notes
                    ERdataSet.Tables(DataTable).Rows(x)("ATTRIBUTE_2") = "M&R"
                    ERdataSet.Tables(DataTable).Rows(x)("ATTRIBUTE_3") = "hospital"

                    x = x + 1
                    cnt += 1
                    'Post the data modification to the database.
                    da.Update(ERdataSet, DataTable)
                End If
            Next

            'Close the SQL connection
            conn.Close()

            If ps.GetText(2, 1, 8) <> "UHC0010:" Then
                SignOff()
            End If

            'Send 'completed' message to text file, or 'no claims in server' message if no claims present.
            If x <> 0 Then
                Dim myfile As String = "C:\Users\Public\MacroBotUpdates.txt"
                If IO.File.Exists(myfile) Then
                    Dim vt As String = vbCrLf & Date.Now & " - " & My.Application.Info.AssemblyName.ToString & "  - Cosmos ER 707 macrobot complete run"
                    My.Computer.FileSystem.WriteAllText(myfile, vt, True)
                End If
            Else
                Dim myfile As String = "C:\Users\Public\MacroBotUpdates.txt"
                If IO.File.Exists(myfile) Then
                    Dim vt As String = vbCrLf & Date.Now & " - " & My.Application.Info.AssemblyName.ToString & "  - Cosmos ER 707 macrobot no claims in server"
                    My.Computer.FileSystem.WriteAllText(myfile, vt, True)
                End If
            End If

        Catch ex As Exception
            Dim myfile As String = "C:\Users\Public\MacroBotUpdates.txt"
            If IO.File.Exists(myfile) Then
                Dim vt As String = vbCrLf & Date.Now & " - " & My.Application.Info.AssemblyName.ToString & "  - Cosmos ER 707 macrobot server error"
                My.Computer.FileSystem.WriteAllText(myfile, vt, True)
            End If
        End Try

    End Sub
    Sub runER(ByRef result As String, ByRef div As String, ByRef claim As String, ByRef notes As String, ByRef notes2 As String)
        Dim password As String = ""
        'Sign in to Cosmos
        If ps.GetText(2, 1, 8) = "UHC0010:" Then
            CosmosSignIn(password)
        End If

        If password = "expired" Then
            notes = "password expired"
            Exit Sub
        End If

        CosmosDivChange(div, claim)

        If result <> "1" Then
            HO400(result, claim, notes, div, notes2, DLMGroupList, CovidProcs, CovidMods, CovidClmsList)
        Else
            'notes2 = "check field not populated"
        End If

    End Sub
    Sub CosmosSignIn(ByRef password As String)
        Dim div As String = ""
        'sub to sign into Cosmos platform

        ApiChk()
        ps.SendKeys("COSMOSP", 3, 1)
        ApiChk()
        ps.SendKeys("[enter]")
        ApiChk()
        ps.Wait(200)

        ComT.SecurePass("COSMOS")

        ps.SendKeys(ComT.SecUser, 10, 26)
        ApiChk()
        ps.SendKeys(ComT.SecPass, 11, 26)
        ApiChk()
        ps.SendKeys("[enter]")
        ApiChk()
        ps.Wait(200)

        If ps.GetText(23, 30, 7) = "expired" Then
            Dim myfile As String = "C:\Users\Public\MacroBotUpdates.txt"
            If IO.File.Exists(myfile) Then
                Dim vt As String = vbCrLf & Date.Now & " - " & My.Application.Info.AssemblyName.ToString & "  - Cosmos password expired"
                My.Computer.FileSystem.WriteAllText(myfile, vt, True)
            End If
            password = "expired"
            Exit Sub
        End If

        ps.SendKeys("KOS")
        ApiChk()
        ps.Wait(200)

        ps.SendKeys("[enter]")
        ApiChk()
        ps.Wait(200)

    End Sub
    Sub CosmosDivChange(ByRef div As String, ByRef claim As String)
        'Goes to the PCOMM Cosmos screen to check what DIV Cosmos is currently in
        Dim DivScreenCheck As String
        Dim CosmosDiv As String, CosmosDiv2 As String
        DivScreenCheck = ps.GetText(4, 51, 4)

        If DivScreenCheck = "Site" Then
            ApiChk()
            ps.SendKeys(div, 5, 52)
            ApiChk()
            ps.SendKeys("[enter]")
            ApiChk()
            ps.Wait(200)
        End If

        CosmosDiv = ps.GetText(1, 64, 3)
        CosmosDiv2 = ps.GetText(1, 66, 3)
        'If DIV is not the same as the current claim then this will go to the DIV change screen "CS010" to change it
        If CosmosDiv <> div Then
            If CosmosDiv2 <> div Then
                ApiChk()
                ps.SendKeys("CS010", 1, 3)
                ApiChk()
                ps.SendKeys("[enter]")
                ApiChk()
                ps.Wait(200)
                ps.SendKeys(div, 8, 47)
                ApiChk()
                ps.SendKeys("[enter]")
                ApiChk()
                ps.Wait(200)
            End If
        End If

    End Sub
    Sub HO400(ByRef result As String, ByRef claim As String, ByRef notes As String, ByRef div As String, ByRef notes2 As String, DLMGroupList As List(Of DelGroups), CovidProcs As List(Of String), CovidMods As List(Of String), CovidClmsList As List(Of CovidClms))


        Dim CHK As String, password As String = "", HO400memberNumber As String = "", DLMFlag As String = ""
        Dim proccode As String, DOS As String, mod1 As String, mod2 As String, amountpaid As Decimal, unt As String
        Dim units As Integer
        Dim POS As String = "", Tele As String = "", icn As String = ""

        ApiChk()
        ps.SendKeys("HO400", 1, 3)
        ApiChk()
        ps.Wait(200)
        ps.SendKeys("[enter]")
        ps.Wait(200)

        ApiChk()
        ps.SendKeys(claim, 4, 6)
        ApiChk()
        ps.SendKeys("[enter]")
        ApiChk()

        If ps.GetText(23, 2, 5) = "DFHAC" Then       'procedure when encountering abend
            ApiChk()
            ps.SendKeys("[clear]")
            ps.Wait(2000)
            ApiChk()
            ps.SendKeys("[clear]")
            ps.Wait(2000)
            ApiChk()
            ps.SendKeys("OFF", 3, 1)
            ApiChk()
            ps.SendKeys("[enter]")
            ps.Wait(2000)
            ApiChk()
            notes = "abend"
            Exit Sub
        End If
        ApiChk()
        CHK = ps.GetText(7, 54, 8)

        'If CHK field = "        " that means that the check hasn't been released, defect
        If CHK = "        " Or CHK = "00000000" Then
            notes2 = "check field not populated"
        Else
            result = "1"
            notes2 = "check field is populated"   'Check has been released, non-defect
            Exit Sub
        End If

        HO400memberNumber = ps.GetText(5, 11, 5)

        'Logic to drop the claim if delegated group found
        For Each ln In DLMGroupList
            If ln.DIV = div And ln.MbrGroup = HO400memberNumber Then
                result = "1"
                notes = "Delegated Claim"
                notes2 = "Delegated group: " & HO400memberNumber
                Exit Sub
            End If
        Next

        DLMFlag = ps.GetText(22, 73, 1)
        If DLMFlag = "Y" Then
            result = "1"
            notes = "Delegated Claim"
            notes2 = "Delegated group: " & HO400memberNumber
            Exit Sub
        End If

        'Logic to remove claim if selected for non-Covid ER and has condition = DIAG/PROC
        icn = div & claim
        For Each clmLn In CovidClmsList
            If clmLn.ClmNbr = icn And (clmLn.Condition = "DIAG" Or clmLn.Condition = "PROC") Then
                result = "1"
                notes = "COVID Criteria met"
                notes2 = "Condition = " & clmLn.Condition
                Exit Sub
            End If

            If clmLn.ClmNbr = icn And clmLn.Condition = "TELE" Then
                Tele = "Yes"
            End If
        Next

        If result <> "1" Then
            'procedure to check to see if the claim is URN or UHB related
            HO409(result, claim, notes2, div)
        End If

        If result = "1" Then
            Exit Sub
        End If

        'Logic for removing the claim if selected for non-covid ER and has condition = TELE
        POS = ps.GetText(8, 77, 4).Trim()

        If POS = "02" And Tele = "Yes" Then
            result = "1"
            notes = "POS=02 found for Cond = TELE"
            Exit Sub
        End If

        Do
            Dim x As Integer = 11
            ApiChk()
            Do
                If result <> "2" Then
                    ApiChk()
                    If (ps.GetText(x, 6, 4)) <> "0001" Then
                        ApiChk()
                        proccode = Trim(ps.GetText(x, 39, 5))
                        ApiChk()
                        If proccode <> "" Then
                            ApiChk()
                            If (Trim((ps.GetText(x, 12, 3))) = "") Then
                                units = 0
                            Else
                                units = CInt(ps.GetText(x, 12, 3))
                            End If
                            DOS = Replace((ps.GetText(x, 31, 6)), (ps.GetText(x, 31, 6)), "20" & Right((ps.GetText(x, 31, 6)), 2) & Left((ps.GetText(x, 31, 6)), 2) & Mid((ps.GetText(x, 31, 6)), 3, 2))
                            mod1 = ps.GetText(x, 44, 2)
                            mod2 = ps.GetText(x, 46, 2)
                            amountpaid = CDec(ps.GetText(x + 5, 72, 7))

                            If CovidProcs.Contains(proccode) And Tele = "Yes" Then

                                If CovidMods.Contains(mod1) Or CovidMods.Contains(mod2) Then
                                    result = "1"
                                    notes = "Modifier found for Cond = TELE"
                                    Exit Sub
                                End If
                            End If

                            compare(result, notes2, claim, proccode, units, DOS, mod1, mod2, amountpaid, proc_list)
                            x = x + 1
                            ApiChk()
                        Else
                            x = x + 1
                        End If
                    Else
                        ApiChk()
                        Exit Do
                    End If
                Else
                    ApiChk()
                    Exit Do
                End If
                ApiChk()
            Loop Until x >= 15
            If ps.GetText(23, 3, 8) = "ORIGINAL" Then
                ps.SendKeys("HO400", 1, 3)
                ApiChk()
                ps.Wait(200)
                ps.SendKeys("[enter]")
                ps.Wait(200)
            Else
                Exit Do
            End If
        Loop
    End Sub
    Sub HO409(ByRef result As String, ByRef claim As String, ByRef notes2 As String, ByRef div As String)
        'code to go into PCOMM Cosmos screen "HO400" to verify if review code present
        Dim ReviewCode As String, ReviewCode2 As String, password As String = ""
        Dim x As Integer
        Dim Reviews As String() = {"00358", "00372", "00383", "00392", "00393", "00433", "00436", "00473", "01033"}

        ps.Wait(200)
        ApiChk()
        ps.SendKeys("HO409", 1, 3)
        ps.Wait(200)
        ApiChk()
        ps.SendKeys("[enter]")

        ApiChk()
        ps.SendKeys(claim, 4, 6)
        ApiChk()
        ps.SendKeys("[enter]")
        ApiChk()

        If ps.GetText(23, 2, 10) = "NO REVIEWS" Then
            'Review code not present, non-defect
            Exit Sub
        End If
        'code to check to see if PCOMM system error "DFHAC"
        If ps.GetText(23, 2, 5) = "DFHAC" Then
            ApiChk()
            ps.SendKeys("[clear]")
            ApiChk()
            ps.SendKeys("[clear]")
            ps.Wait(750)
            ApiChk()
            ps.SendKeys("OFF", 3, 1)
            ApiChk()
            ps.SendKeys("[enter]")
            ps.Wait(1500)
            ApiChk()
            Exit Sub
        End If

        x = 8
        Do
            ReviewCode = ps.GetText(x, 7, 5)
            ReviewCode2 = ps.GetText(x, 47, 5)
            'If ReviewCode = 00603 or 00269 then claim is a non defect
            If ReviewCode = "00269" Or ReviewCode2 = "00269" Then
                result = "1"
                notes2 = "URN/UBH review code found"
                Exit Sub
            End If

            'Review Codes 382, 2099 & 2186 removed per PAWS project work item 5749
            If ReviewCode = "00382" Or ReviewCode = "02099" Or ReviewCode = "02186" Or ReviewCode2 = "00382" Or ReviewCode2 = "02099" Or ReviewCode2 = "02186" Then
                result = "1"
                notes2 = "Pay Subscriber review code found"
                Exit Sub
            End If

            'Review 479 per PAWS project workflow work item 5869
            If ReviewCode = "00479" Or ReviewCode2 = "00479" Then
                result = "1"
                notes2 = "Medicaid Reclamation review code found"
                Exit Sub
            End If

            If (Reviews.Contains(ReviewCode)) Or (Reviews.Contains(ReviewCode2)) Then
                result = 1
                notes2 = "707 Review found " & ReviewCode & ReviewCode2
                Exit Sub
            Else
                x = x + 1
                If ReviewCode = "     " Then
                    'HO410(result, claim, notes2)
                    ps.Wait(200)
                    ApiChk()
                    ps.SendKeys("HO400", 1, 3)
                    ps.Wait(200)
                    ApiChk()
                    ps.SendKeys("[enter]")
                    Exit Sub
                End If
            End If
        Loop Until x >= 23

    End Sub
    Sub HO410(ByRef result As String, ByRef claim As String, ByRef notes2 As String)
        'code to go into PCOMM Cosmos screen "HO410" to verify if LEGAL ENT = U, if yes then it is a no-defect

        Dim catcode As String, subcatcode As String
        Dim allowcat As String() = {"099", "399", "679"}

        ps.Wait(200)
        ps.SendKeys("HO410", 1, 3)
        ps.Wait(200)
        ps.SendKeys("[enter]")

        ps.SendKeys(claim, 4, 6)
        ApiChk()
        ps.SendKeys("[enter]")
        'Check to see if claim is CRT related, if so it is a non-defect
        If ps.GetText(6, 72, 3) = "CRT" Then
            result = "1"
            notes2 = "CRT Claim"
            Exit Sub
        End If

        If ps.GetText(10, 78, 1) = "U" Then
            'LEGAL ENT = U, non-defect
            result = "1"
            notes2 = "LEGAL ENT = U"
            Exit Sub
        End If

        Dim cat As Integer = 15

        Do
            catcode = ps.GetText(7, cat, 3)
            ApiChk()
            subcatcode = ps.GetText(8, cat, 3)
            ApiChk()
            If (allowcat.Contains(catcode)) And subcatcode = "001" Then
                result = "2"
                notes2 = "contracted by % billed"
                cat = cat + 4
                Exit Sub
            Else
                result = 1
                notes2 = " "
                cat = cat + 4
            End If
        Loop Until cat >= 54
    End Sub

    '=========================================
    Sub ApiChk()
        Dim x As Integer
        'Code to slow down system and wait for entry
        x = 1
        Do Until oia.InputInhibited = 0
            ps.Wait(600)
            x = x + 1
            If x = 50 Then
                ps.SendKeys("[reset]")
                ps.Wait(1000)
                x = 1
            End If
        Loop

    End Sub
    'Sub compare(ByRef result As String, ByRef notes2 As String, ByRef claim As String, procedurelist As List(Of procedure), proc_list As List(Of procs))
    Sub compare(ByRef result As String, ByRef notes2 As String, ByRef claim As String, ByRef proccode As String, ByRef units As Integer, ByRef DOS As String, ByRef mod1 As String, ByRef mod2 As String, ByRef amountpaid As Decimal, proc_list As List(Of procs))

        '*************************************code for 707************************************

        Dim modifier As String() = {"59", "76", "77", "91"}
        Dim pcd As String = "", startdt As String = "", enddt As String = ""
        Dim unit As Integer

        For Each sourceproc In proc_list
            If sourceproc.procID = proccode And (DOS >= sourceproc.startdte And DOS <= sourceproc.eddt) Then
                pcd = sourceproc.procID
                startdt = sourceproc.startdte
                enddt = sourceproc.eddt
                unit = sourceproc.units_alw
                Exit For
            End If
        Next
        '**********************************Code for comparing the claim details with the details from the table
        If (proccode = pcd) And (units > unit) And (modifier.Contains(mod1)) Or (modifier.Contains(mod2)) Then
            result = "1"
            notes2 = "modifier applied"
            Exit Sub
        Else
            If (proccode = pcd) And (DOS >= startdt) And (DOS <= enddt) And (units > unit) And (amountpaid > 0) Then
                ApiChk()
                HO410(result, claim, notes2)
                If result <> "2" Then
                    result = "2"
                    notes2 = "Units billed for code " & proccode & " is > allowed units"
                ElseIf result = "2" Then
                    Exit Sub
                End If
            Else
                result = "1"
                notes2 = "Not a defective claim"
            End If
        End If

    End Sub
    Sub SignOff()
        Dim div As String = "", claim As String = "", ClaimTypeBill As String = ""
        ps.SendKeys("BYE  ", 1, 3)
        ps.SendKeys("[enter]")
        ps.Wait(750)
        ApiChk()
        ps.SendKeys("[clear]")
        ps.Wait(750)
        ApiChk()
        ps.SendKeys("[clear]")
        ps.Wait(750)
        ApiChk()
        ps.SendKeys("OFF", 3, 1)
        ps.SendKeys("[enter]")
        ps.Wait(750)
        ApiChk()

    End Sub
    Sub GetProc_list()

        'Dim proc_list = New List(Of procs)()

        'setting the connection string to the test database mirroring the SAM server
        Const connection2 = "Data Source=DBSEP2093;Initial Catalog=COSMOS_ER;Integrated Security=True"
        Dim proc_cde_707 As New SqlConnection(connection2)
        Dim DXdataset As New DataSet()
        Dim da1 As SqlDataAdapter
        Dim dRow As DataRow
        Dim cmdBuilder As SqlCommandBuilder

        Try
            proc_cde_707.Open()
            'Initialize the SqlDataAdapter object by specifying a Select command 

            da1 = New SqlDataAdapter("Select * from procedure_cde_1013", proc_cde_707)

            'Initialize the SqlCommandBuilder object to automatically generate and initialize
            'the UpdateCommand, InsertCommand and DeleteCommand properties of the SqlDataAdapter.

            cmdBuilder = New SqlCommandBuilder(da1)
            'Populate the dataset by running the Fill method of the SqlDataAdapter.
            da1.Fill(DXdataset, "sourceProcedure")

            For Each dRow In DXdataset.Tables("sourceProcedure").Rows
                Dim listproc = New procs()
                listproc.Proc_cde = Trim(dRow.Item("procedure_cde"))
                listproc.startdte = Trim(dRow.Item("start_dt"))
                listproc.enddte = Trim(dRow.Item("stop_dt"))
                listproc.units_alw = CInt(dRow.Item("unit_svc"))
                proc_list.Add(listproc)
            Next

            ' proc_cde_707.Close()

            ' For Each proc In proc_list
            'Console.WriteLine("Procedure: {0},{1},{2},{3}", proc.Proc_cde, proc.startdte, proc.enddte, proc.units_alw)
            'Next
            'Console.ReadLine()

        Catch ex As Exception
            'send server error note to text file in my documents folder
            Dim myfile As String = "C:\Users\Public\MacrobotUpdates.txt"
            If IO.File.Exists(myfile) Then
                Dim vt As String = vbCrLf & Date.Now & " - " & My.Application.Info.AssemblyName.ToString & "  - Cosmos ER 707 Procedure list server error"
                My.Computer.FileSystem.WriteAllText(myfile, vt, True)
            End If
        End Try
    End Sub

    Sub PopulateDelGroups()
        Dim connection As String = "Data Source=DBSEP2093;Initial Catalog=COSMOS_ER;Integrated Security=True"

        Dim SAMDevConn As SqlConnection = New SqlConnection(connection)
        Dim ERdataSet As DataSet = New DataSet()
        Dim da As SqlDataAdapter
        Dim cmdBuilder As SqlCommandBuilder

        Try
            SAMDevConn.Open()

            da = New SqlDataAdapter("Select * from WS_DIV_DelegatedGroups", SAMDevConn)

            cmdBuilder = New SqlCommandBuilder(da)

            da.Fill(ERdataSet, "DelGroup")

            For Each row As DataRow In ERdataSet.Tables("DelGroup").Rows
                Dim listDelGroup = New DelGroups()
                listDelGroup.DIV = row.ItemArray(0).ToString().Trim()
                listDelGroup.MbrGroup = row.ItemArray(1).ToString().Trim()

                DLMGroupList.add(listDelGroup)
            Next
            SAMDevConn.Close()
            da.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Sub PopulateCovidProcsMods()
        Dim connection As String = "Data Source=DBSEP2093;Initial Catalog=COSMOS_ER;Integrated Security=True"

        Dim SAMDevConn As SqlConnection = New SqlConnection(connection)
        Dim ERdataSet As DataSet = New DataSet()
        Dim da As SqlDataAdapter
        Dim cmdBuilder As SqlCommandBuilder

        Try
            SAMDevConn.Open()

            da = New SqlDataAdapter("Select * from COVID_TELE_EXCLUSIONS_COSMOS", SAMDevConn)

            cmdBuilder = New SqlCommandBuilder(da)

            da.Fill(ERdataSet, "Covidgroup")

            For Each row As DataRow In ERdataSet.Tables("Covidgroup").Rows
                Dim CovidProc_CD As String
                Dim CovidMdfr As String

                CovidProc_CD = row.ItemArray(0).ToString().Trim()
                CovidMdfr = row.ItemArray(1).ToString().Trim()

                If CovidProc_CD <> "" Then
                    CovidProcs.add(CovidProc_CD)
                End If

                If CovidMdfr <> "" Then
                    CovidMods.add(CovidMdfr)
                End If
            Next

            SAMDevConn.Close()
            da.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Sub PopulateCovidClms()
        Dim connection As String = "Data Source=DBSEP2093;Initial Catalog=COSMOS_ER;Integrated Security=True"

        Dim SAMDevConn As SqlConnection = New SqlConnection(connection)
        Dim ERdataSet As DataSet = New DataSet()
        Dim da As SqlDataAdapter
        Dim cmdBuilder As SqlCommandBuilder

        Try
            SAMDevConn.Open()

            da = New SqlDataAdapter("Select DISTINCT INV_CTL_NBR, CONDITION from covid_criteria_Claims_H", SAMDevConn)

            cmdBuilder = New SqlCommandBuilder(da)

            da.Fill(ERdataSet, "CovidClmgroup")

            For Each row As DataRow In ERdataSet.Tables("CovidClmgroup").Rows
                Dim listCovidClms = New CovidClms()

                listCovidClms.ClmNbr = row.ItemArray(0).ToString().Trim()
                listCovidClms.Condition = row.ItemArray(1).ToString().Trim()

                CovidClmsList.add(listCovidClms)
            Next

            SAMDevConn.Close()
            da.Dispose()
        Catch ex As Exception

        End Try
    End Sub
End Module