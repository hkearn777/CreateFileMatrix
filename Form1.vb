Imports System.IO
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Windows.Win32.System
Public Class Form1
  ' CreateFileMatrix
  '
  Dim ProgramVersion As String = "0.1"
  Const Delimiter As String = "|"
  Const Quote As String = Chr(34)
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim SummaryWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FilesWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim ProgramsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim RecordsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FileToAppDispWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FileToAppsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim AppMatrixWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim testWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim rngSummary As Microsoft.Office.Interop.Excel.Range
  Dim rngFileToAppDisp As Microsoft.Office.Interop.Excel.Range
  Dim rngFileToApps As Microsoft.Office.Interop.Excel.Range
  Dim rngAppMatrix As Microsoft.Office.Interop.Excel.Range

  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
  Dim SetAsReadOnly = Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly


  Dim ListOfApps As New List(Of String)           'array which holds the Application name, start and end indexes
  Dim ListOfPrograms As New List(Of String)       'array which holds Program Name and Source Type
  Dim ListOfFileAppDisp As New List(Of String)    'Array which holds every file | app | DISP
  Dim ListOfFileApps As New List(Of String)       'Array which hold Unique file | app1 | app2 ... | appN
  Dim ListOfAppsColumns As New List(Of String)    'Array to hold the Apps Names found
  Dim ListOfDroppedFiles As New List(Of String)   'Array to hold the dropped datasets

  Dim FileMatrix As New List(Of String)           'Array which holds the Files and the apps [dns|app|disp|app1|disp1|...|appN|dispN]
  Dim AppsMatrix As New List(Of String)           'Array which holds the Apps to Apps
  Dim DDsToSkip As New List(Of String)            'Array which holds the list of DDNames to skip
  Dim DSNsToSkip As New List(Of String)           'Array which holds the list of Dataset names to skip

  Dim MostApps As Integer = 0


  Private Sub btnCreateFileMatrix_Click(sender As Object, e As EventArgs) Handles btnCreateFileMatrix.Click
    Me.Cursor = Cursors.WaitCursor
    Call LoadModelsToArrays()
    'Call BuildFileMatrix()
    'Call BuildAppsMatrix()

    Dim FileMatrixFileName As String = txtSandboxRoot.Text & "CreateFileMatrix.xlsx"
    Call RemovePreviousExcelFile(FileMatrixFileName)
    Call CreateSummaryTab()
    Call CreateFilesToAppsTab()
    'Call CreateFilesToAppsDispTab()
    Call CreateAppsToAppsCountsTab()
    Call SaveCloseAndQuitExcel(FileMatrixFileName)

    System.Media.SystemSounds.Beep.Play()
    Me.Cursor = Cursors.Default
  End Sub

  Sub BuildFileMatrix()
    ' Now go through the ListOfFileAppDisp array [dsn|app|disp] one entry at a time
    ' looking for other matching dataset names
    ' to append the [app|disp] value
    ' Output FileMatrix: [dns|app|disp|app1|disp1|...|appN|dispN]

    ' obsolete, just do this at the write of the tab instead.

    lblFileMatrix.Text = "File Matrix: Analyzing..."
    MostApps = 0
    For Each fileEntry In ListOfFileAppDisp
      Dim fileMatrixList As String = fileEntry    'load with first file|app|disp
      Dim fileEntries As String() = fileEntry.Split(Delimiter)
      Dim NumAppsForFile As Integer = 0
      For Each fileusage In ListOfFileAppDisp
        If fileEntry = fileusage Then
          Continue For
        End If
        Dim fileUsages As String() = fileusage.Split(Delimiter)
        If fileUsages(0) <> fileEntries(0) Then
          Continue For
        End If
        NumAppsForFile += 1
        If NumAppsForFile > MostApps Then
          MostApps = NumAppsForFile
        End If
        fileMatrixList &= Delimiter & fileUsages(1)
      Next
      FileMatrix.Add(fileMatrixList)
    Next
  End Sub

  Sub BuildAppsMatrix()
    ' create the AppsMatrix: app|app1|app2...|appN

    ' at moment ***obsolete***

    For Each App In ListOfApps
      Dim HoldList As New List(Of String)
      Dim tempList As String() = App.Split(Delimiter)
      Dim AppList As String = tempList(0)
      Dim StartIndex As Integer = Val(tempList(1)) - 1
      Dim EndIndex As Integer = Val(tempList(2)) - 1
      For Index As Integer = StartIndex To EndIndex
        Dim FileEntries As String() = FileMatrix(Index).Split(Delimiter)
        For x As Integer = 3 To FileEntries.Count - 1 Step 2
          If FileEntries(x) = AppList Then 'dont store same app
            Continue For
          End If
          If HoldList.IndexOf(FileEntries(x)) = -1 Then
            HoldList.Add(FileEntries(x))
          End If
        Next
      Next
      HoldList.Sort()
      For Each ReferencedApp In HoldList
        AppList &= Delimiter & ReferencedApp
      Next
      ' create the list of apps to apps (apps matrix)
      AppsMatrix.Add(AppList)
      'HoldList.Clear()
    Next

  End Sub
  Sub LoadModelsToArrays()
    ' Process through each Model xlsx file
    ' This will populate the arrays: ListOfFiles and ListOfApps
    Dim myFile As String = ""
    Dim myAppl As String = ""
    Dim myDisp As String = ""
    Dim modelsProcessed As Integer = 0
    Dim AppStartIndex As Integer = 0
    For Each model In lbModels.Items
      ' Open the Model xlsx
      Dim modelName As String = model.ToString
      Dim folders As String() = modelName.Split("\")
      myAppl = folders(7).Replace(" ", "_")
      workbook = objExcel.Workbooks.Open(modelName, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)

      ' need to verify the worksheets exists
      Dim FilesWorksheetFound As Boolean = False
      Dim ProgramsWorksheetFound As Boolean = False
      Dim RecordsWorksheetFound As Boolean = False
      Dim WorksheetFoundCnt As Integer = 0
      If RequiredWorksheetsFound() <> 3 Then
        lbMessages.Items.Add("Files, Programs, or Records Worksheet not found:" & modelName)
        workbook.Close()
        Continue For
      End If

      modelsProcessed += 1
      AppStartIndex += 1
      ' Point to the worksheets
      ProgramsWorksheet = CType(workbook.Worksheets("Programs"), Microsoft.Office.Interop.Excel.Worksheet)
      RecordsWorksheet = CType(workbook.Worksheets("Records"), Microsoft.Office.Interop.Excel.Worksheet)
      FilesWorksheet = CType(workbook.Worksheets("Files"), Microsoft.Office.Interop.Excel.Worksheet)

      ' Load the Programs used with their SourceType 
      Call LoadProgramsAndSourceType()

      ' Analyze the Files worksheet
      FilesWorksheet.Activate()

      'todo: need to validate columns are right in case format has changed
      Dim LastRow As Integer = FilesWorksheet.UsedRange.Rows.Count
      Dim AddedToMatrix As Integer = 0
      For x As Integer = 2 To LastRow
        Dim sRow As String = LTrim(Str(x))
        ' Get the DDNAme and skip SYSTEM LEVEL DD's ie. STEPLIB
        Dim DDName As String = FilesWorksheet.Range("G" & sRow).Value
        If DDsToSkip.IndexOf(DDName) > -1 Then
          Continue For
        End If
        ' Get the Dataset / filename and skip some datasets ie., SYSOUT, WORKSPACE, etc.
        myFile = FilesWorksheet.Range("J" & sRow).Value
        If myFile Is Nothing Then
          Continue For
        End If
        If myFile.StartsWith("SYSOUT=") Or
            myFile.StartsWith("WORKSPACE") Or
            myFile.StartsWith("&&") Or
            myFile.Trim.Length = 0 Then
          Continue For
        End If
        ' clean up poorly formated file names
        myFile = myFile.Replace("*", " ").Replace(Quote, "").Trim
        Dim myFileWords As String() = myFile.Split(" ")
        If myFileWords(0).Length = 0 Then
          Continue For
        End If
        myFile = myFileWords(0)
        If GetDSNType(myFile) = "PDSMember" Then
          Continue For
        End If
        If DSNsToSkip.IndexOf(myFile) > -1 Then
          Continue For
        End If

        Dim PgmName = FilesWorksheet.Range("F" & sRow).Value
        If PgmName Is Nothing Then
          PgmName = ""
        End If

        ' get disposition/Access.
        ' TODO: If Utility use the DISP else if code is found look into the code open mode
        myDisp = FilesWorksheet.Range("K" & sRow).Value
        If myDisp Is Nothing Then
          myDisp = ""
        End If
        myDisp = myDisp.Remove(1)
        '
        Dim fileEntry As String = myFile & Delimiter & myAppl & Delimiter & myDisp
        If ListOfFileAppDisp.IndexOf(fileEntry) = -1 Then
          ListOfFileAppDisp.Add(fileEntry)
          AddedToMatrix += 1
        End If
      Next


      Dim AppEndIndex As String = LTrim(Str(AppStartIndex + AddedToMatrix - 1))
      Dim AppsEntry As String = myAppl & Delimiter & LTrim(Str(AppStartIndex)) & Delimiter & AppEndIndex
      ListOfApps.Add(AppsEntry)
      AppStartIndex = Val(AppEndIndex)
      workbook.Close()
      lblModelsProcessed.Text = "Models Processed:" & LTrim(Str(modelsProcessed))

    Next
    ListOfFileAppDisp.Sort()


    ' build a list of columns of apps
    For Each App In ListOfApps
      Dim AppEntries As String() = App.Split(Delimiter)
      ListOfAppsColumns.Add(AppEntries(0))
    Next

    lblModelsProcessed.Text = "Models Loaded: complete."

  End Sub
  Function RequiredWorksheetsFound() As Integer
    Dim worksheetFoundCnt As Integer = 0
    For Each objWorksheet In workbook.Worksheets
      Select Case objWorksheet.name.ToString.ToUpper
        Case "FILES" : worksheetFoundCnt += 1
        Case "PROGRAMS" : worksheetFoundCnt += 1
        Case "RECORDS" : worksheetFoundCnt += 1
      End Select
    Next
    Return worksheetFoundCnt
  End Function
  Sub RemovePreviousExcelFile(ByRef FileMatrixFileName As String)
    ' remove previous excel file
    If File.Exists(FileMatrixFileName) Then
      Try
        File.Delete(FileMatrixFileName)
      Catch ex As Exception
        MessageBox.Show("Error deleting File Matrix spreadsheet:" & ex.Message)
        objExcel.Quit()
        Me.Cursor = Cursors.Default
      End Try
    End If

  End Sub
  Sub CreateSummaryTab()
    ' this will create the first tab with details about this run
    workbook = objExcel.Workbooks.Add
    SummaryWorksheet = workbook.Sheets.Item(1)
    SummaryWorksheet.Name = "Summary"
    SummaryWorksheet.Range("A1").Value = "Mainframe Documentation Project"
    SummaryWorksheet.Range("A2").Value = "Interface Model"
    SummaryWorksheet.Range("A3").Value = "Created:" & Date.Now
    SummaryWorksheet.Range("A4").Value = "CreateFileMatrix, Version:" & ProgramVersion

    rngSummary = SummaryWorksheet.Range("A1:A4")
    rngSummary.Font.Bold = True
    rngSummary.Font.Size = 16
    rngSummary.Columns.AutoFit()
    rngSummary.Rows.AutoFit()
    rngSummary.VerticalAlignment = XlVAlign.xlVAlignTop

  End Sub

  Sub CreateFilesToAppsTab()
    ' now create the file matrix Excel document (Files-to-Apps)
    ' this will be Unique Dataset names down the Rows and columns with the App names. 
    '   the cells will be filled with the DISP values 
    ' Input: ListOfFileAppDisp array


    lblFileMatrix.Text = "File Matrix: Writing to Excel tab: files-to-Apps..."

    ' set up the column headings
    Dim row As Integer = 1
    Dim lastCol As Integer = 0
    FileToAppsWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
    FileToAppsWorksheet.Name = "Files-to-Apps"
    FileToAppsWorksheet.Range("A1").Value = "Dataset Files"
    FileToAppsWorksheet.Range("B1").Value = "Count"
    For col As Integer = 0 To ListOfAppsColumns.Count - 1
      FileToAppsWorksheet.Cells(row, col + 3).value = ListOfAppsColumns(col)
    Next
    lastCol = ListOfApps.Count + 2
    FileToAppsWorksheet.Application.ActiveWindow.SplitRow = 1
    FileToAppsWorksheet.Application.ActiveWindow.SplitColumn = 1
    FileToAppsWorksheet.Application.ActiveWindow.FreezePanes = True


    ' for each entry:
    '  compute row: if dsn break row += 1;
    '  compute col: base on App name
    '  compute cell: Is there an any value:
    '                no: assign "" 
    '                yes: is it already an I or an O
    For Each Entry In ListOfFileAppDisp
      Dim subitems As String() = Entry.Split(Delimiter)
      Dim dsn As String = subitems(0)
      Dim app As String = subitems(1)
      Dim disp As String = subitems(2)
      'compute the row: if dsn break row += 1
      If FileToAppsWorksheet.Cells(row, 1).value <> dsn Then
        row += 1
        FileToAppsWorksheet.Cells(row, 1).value = dsn
        FileToAppsWorksheet.Cells(row, 2).value = "0"
      End If
      ' compute the col: based on App name
      Dim col As Integer = ListOfAppsColumns.IndexOf(app) + 3
      ' compute the cell content
      If FileToAppsWorksheet.Cells(row, col).value Is Nothing Then
        FileToAppsWorksheet.Cells(row, col).value = disp
        FileToAppsWorksheet.Cells(row, 2).value = LTrim(Str(Val(FileToAppsWorksheet.Cells(row, 2).value) + 1))
      End If
      Dim currdisp As String = FileToAppsWorksheet.Cells(row, col).value
      If currdisp.IndexOf(disp) = -1 Then
        FileToAppsWorksheet.Cells(row, col).value &= disp
      End If
    Next

    ''obsolete below....
    '' put the matrix details to worksheet
    '' find which column the app belongs to
    'For Each fileentry In FileMatrix
    '  row += 1
    '  Dim col As Integer = 1
    '  Dim FileEntries As String() = fileentry.Split(Delimiter)
    '  FileToAppsWorksheet.Cells(row, col).Value = FileEntries(0)              'dataset name
    '  ' count unique apps in the fileentries
    '  Dim countOfapps As Integer = CountUniqueApps(FileEntries)
    '  FileToAppsWorksheet.Cells(row, col + 1).Value = LTrim(Str(countOfapps))
    '  For index As Integer = 1 To FileEntries.Length - 1
    '    ' find column to place the X value 
    '    col = ListOfAppsColumns.IndexOf(FileEntries(index)) + 3
    '    FileToAppsWorksheet.Cells(row, col).Value &= "x"
    '  Next
    'Next

    ' Finishing presentation touches
    Dim myRange As String = getRangeInA1Style(1, 1, 1, lastCol)
    rngFileToApps = FileToAppsWorksheet.Range(myRange)
    rngFileToApps.Font.Bold = True
    workbook.Worksheets("Files-to-Apps").Range("A1").AutoFilter
    myRange = getRangeInA1Style(1, 1, row, lastCol)
    rngFileToApps = FileToAppsWorksheet.Range(myRange)
    rngFileToApps.VerticalAlignment = XlVAlign.xlVAlignTop

    myRange = getRangeInA1Style(1, 3, row, lastCol)
    FileToAppsWorksheet.Range(myRange).Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter

    rngFileToApps.Columns.AutoFit()
    rngFileToApps.Rows.AutoFit()

    lblFileMatrix.Text = "File Matrix: Writing to Excel tab: files-to-Apps complete."

  End Sub
  Function CountUniqueApps(ByRef FileEntries As String()) As Integer
    Dim DistinctValues As New List(Of String)
    For x = 1 To FileEntries.Length - 1
      If DistinctValues.IndexOf(FileEntries(x)) = -1 Then
        DistinctValues.Add(FileEntries(x))
      End If
    Next
    Return DistinctValues.Count
  End Function
  Function GetUniqueApps(ByRef FileEntries As String()) As List(Of String)
    Dim DistinctValues As New List(Of String)
    For x = 1 To FileEntries.Length - 1
      If DistinctValues.IndexOf(FileEntries(x)) = -1 Then
        DistinctValues.Add(FileEntries(x))
      End If
    Next
    Return DistinctValues
  End Function
  Sub LoadProgramsAndSourceType()
    ProgramsWorksheet.Activate()
    Dim LastRow As Integer = ProgramsWorksheet.UsedRange.Rows.Count
    For x = 2 To LastRow
      Dim sRow As String = LTrim(Str(x))
      ' Get the DDNAme and skip SYSTEM LEVEL DD's ie. STEPLIB
      Dim PgmName As String = ProgramsWorksheet.Range("F" & sRow).Value
      If PgmName Is Nothing Then
        Continue For
      End If
      If PgmName.Trim.Length = 0 Then
        Continue For
      End If
      Dim SourceType As String = ProgramsWorksheet.Range("G" & sRow).Value
      Dim ProgramEntry As String = PgmName & Delimiter & SourceType
      If ListOfPrograms.IndexOf(ProgramEntry) = -1 Then
        ListOfPrograms.Add(ProgramEntry)
      End If
    Next
  End Sub
  Sub CreateAppsToAppsCountsTab()
    ' now create the Apps-to-Apps excel tab
    ' this will be based on the ListOfFileAppDisp: dsn,app,disp (sorted)
    ' The Rows first Columns will be the App name
    ' The First Row's columns will be the App name

    lblFileMatrix.Text = "File Matrix: Writing to Excel tab: Apps-to-Apps..."

    Dim row As Integer = 1
    Dim lastCol As Integer = 0
    AppMatrixWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
    AppMatrixWorksheet.Name = "Apps-to-Apps"
    ' set up the column headings with the App Names
    AppMatrixWorksheet.Range("A1").Value = "App"
    For col As Integer = 0 To ListOfAppsColumns.Count - 1
      AppMatrixWorksheet.Cells(1, col + 2).value = ListOfAppsColumns(col)
    Next
    ' set up the row headings in column 1 with the App Names
    For row = 0 To ListOfAppsColumns.Count - 1
      AppMatrixWorksheet.Cells(row + 2, 1).value = ListOfAppsColumns(row)
    Next

    lastCol = ListOfApps.Count + 1

    AppMatrixWorksheet.Application.ActiveWindow.SplitRow = 1
    AppMatrixWorksheet.Application.ActiveWindow.SplitColumn = 1
    AppMatrixWorksheet.Application.ActiveWindow.FreezePanes = True

    'Compute row / col
    '  if dsn = prev dsn; set col = app + 2
    '  else set row = app and col = app + 2

    Dim PrevDSN As String = ""
    For Each Entry In ListOfFileAppDisp
      Dim subitems As String() = Entry.Split(Delimiter)
      Dim dsn As String = subitems(0)
      Dim app As String = subitems(1)
      Dim disp As String = subitems(2)
      Dim col As Integer = ListOfAppsColumns.IndexOf(app) + 2
      If dsn <> PrevDSN Then
        row = col
      End If
      ' compute the cell content
      If AppMatrixWorksheet.Cells(row, col).value Is Nothing Then
        AppMatrixWorksheet.Cells(row, col).value = "0"
      End If
      AppMatrixWorksheet.Cells(row, col).value = LTrim(Str(Val(AppMatrixWorksheet.Cells(row, col).value) + 1))

      PrevDSN = dsn
    Next

    '' put the File matrix details to Apps-to-Apps worksheet
    '' go through each file and put a count in each cell
    '' first should be Row and others should be column
    'For Each Filentry In FileMatrix
    '  Dim FileEntries As String() = Filentry.Split(Delimiter)
    '  ' skip if dataset is a PDS member 
    '  If GetDSNType(FileEntries(0)) = "PDSMember" Then
    '    Continue For
    '  End If
    '  Dim ListOfUniqueApps As New List(Of String)
    '  ListOfUniqueApps = GetUniqueApps(FileEntries)
    '  row = ListOfAppsColumns.IndexOf(ListOfUniqueApps(0)) + 2
    '  For Each app In ListOfUniqueApps
    '    Dim col As Integer = ListOfAppsColumns.IndexOf(app) + 2
    '    ' initialize to zero if new cell
    '    If AppMatrixWorksheet.Cells(row, col).Value Is Nothing Then
    '      AppMatrixWorksheet.Cells(row, col).Value = "0"
    '    End If
    '    ' add to counter
    '    Dim newValue As Integer = Val(AppMatrixWorksheet.Cells(row, col).Value) + 1
    '    AppMatrixWorksheet.Cells(row, col).Value = newValue
    '  Next
    'Next

    row = lastCol

    ' Finishing presentation touches on Apps-to-Apps
    Dim myRange As String = getRangeInA1Style(1, 1, 1, lastCol)
    rngAppMatrix = AppMatrixWorksheet.Range(myRange)
    rngAppMatrix.Font.Bold = True
    workbook.Worksheets("Apps-to-Apps").Range("A1").AutoFilter
    myRange = getRangeInA1Style(1, 1, row, lastCol)
    rngAppMatrix = AppMatrixWorksheet.Range(myRange)
    rngAppMatrix.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    rngAppMatrix.Columns.AutoFit()
    rngAppMatrix.Rows.AutoFit()

    lblFileMatrix.Text = "File Matrix: Writing to Excel tab: file-to-file complete."

  End Sub
  Function GetDSNType(ByRef dsn As String) As String
    If dsn.IndexOf("()") > -1 Then
      Return "GDG"
    End If
    If dsn.IndexOf("(") > -1 Then
      Return "PDSMember"
    End If
    Return "Normal"
  End Function
  Function getRangeInA1Style(srow As Integer, scol As Integer, erow As Integer, ecol As Integer) As String
    ' must be 1 based not 0 based
    Dim inputFormula As String = "=R" & LTrim(Str(srow)) & "C" & LTrim(Str(scol)) & ":" &
                        "R" & LTrim(Str(erow)) & "C" & LTrim(Str(ecol))
    Dim outputFormula As String = (objExcel.ConvertFormula(Formula:=inputFormula,
                                            FromReferenceStyle:=XlReferenceStyle.xlR1C1,
                                            ToReferenceStyle:=XlReferenceStyle.xlA1))
    Return outputFormula.Replace("=", "").Replace("$", "")
  End Function
  Sub SaveCloseAndQuitExcel(ByRef FileMatrixFileName As String)
    ' save, close the workbook, and quit
    lblFileMatrix.Text = "File Matrix: Saving Excel..."
    workbook.SaveAs(FileMatrixFileName, DefaultFormat,,, SetAsReadOnly)
    workbook.Close()
    objExcel.Quit()
    lblFileMatrix.Text = "File Matrix: Complete."
  End Sub
  Private Sub btnFindModels_Click(sender As Object, e As EventArgs) Handles btnFindModels.Click
    If Not Directory.Exists(txtSandboxRoot.Text) Then
      MessageBox.Show("Sandbox directory not found:" & vbLf & txtSandboxRoot.Text)
      Exit Sub
    End If
    Dim cntModelsFound As Integer = 0

    For Each foundfile As String In My.Computer.FileSystem.GetFiles(txtSandboxRoot.Text,
                                     FileIO.SearchOption.SearchAllSubDirectories, "*model.xlsx")
      Dim SandboxFound As Boolean = False
      Dim OutputFound As Boolean = False
      Dim folders As String() = foundfile.Split("\")
      ' validate file is a Application model.xlsx file
      For x = 0 To folders.Length - 1
        If folders(x) = "Sandbox" Then
          SandboxFound = True
        End If
        If folders(x) = "OUTPUT" Then
          OutputFound = True
        End If
      Next
      If SandboxFound = True And OutputFound = True Then
        lbModels.Items.Add(foundfile)
      End If
    Next
    lblModelsFound.Text = "Models found:" & lbModels.Items.Count
  End Sub
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = "CreateFileMatrix " & ProgramVersion
    txtSandboxRoot.Text = "C:\Users\906074897\Documents\All Projects\State of Illinois\Sandbox\"
    objExcel.Visible = False
    objExcel.ErrorCheckingOptions.NumberAsText = False

    DDsToSkip.Add("STEPLIB")
    DDsToSkip.Add("PROCLIB")
    DDsToSkip.Add("DFSRESLB")
    DDsToSkip.Add("IMS")
    DDsToSkip.Add("SORTLIB")
    DDsToSkip.Add("SYSUDUMP")
    DDsToSkip.Add("IEFRDER")
    DDsToSkip.Add("CEEDUMP")
    DDsToSkip.Add("IMSLOGR")
    DDsToSkip.Add("DFSVSAMP")
    DDsToSkip.Add("LEMSGS")
    DDsToSkip.Add("IMSERROR")
    DDsToSkip.Add("ABNLHELP")
    DDsToSkip.Add("ABNLSPRT")
    DDsToSkip.Add("ABNLDUMP")
    DDsToSkip.Add("EZTVFM")
    DDsToSkip.Add("MACROS")
    DDsToSkip.Add("DMNETMAP")
    DDsToSkip.Add("DMMSGFIL")
    DDsToSkip.Add("PANDD1")
    DDsToSkip.Add("PANDD2")

    DSNsToSkip.Add("IMSVSS.FINALIST.CTYSTATE")
    DSNsToSkip.Add("IMSVSS.FINALIST.DATAFILE")
    DSNsToSkip.Add("IMSVSS.FINALIST.DPVDB")
    DSNsToSkip.Add("IMSVSS.FINALIST.DPVSUD")
    DSNsToSkip.Add("IMSVSS.FINALIST.EWSFILE")
    DSNsToSkip.Add("IMSVSS.FINALIST.LLKDB")
    DSNsToSkip.Add("IMSVSS.FINALIST.LLKSUD")
    DSNsToSkip.Add("IMSVSS.FINALIST.SLKDB")
    DSNsToSkip.Add("MHEALTH.PANLIB")



  End Sub

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub


End Class
