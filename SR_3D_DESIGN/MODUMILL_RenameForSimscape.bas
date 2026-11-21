Attribute VB_Name = "modRenameForSimscape"
Option Explicit
' ============================================================================
'  MODUMILL  ->  MATLAB / Simscape Multibody preparation macro
'  Generated: 2026-11-02   Plan: rename_and_material_map.json
'  Source audit: MODUMILL_DeepAudit.json  (SOLIDWORKS 34.1.1 / 2026)
'  Macro version: 3  -  LOW-MEMORY "one file at a time" edition
'
'  WHY v3:  v2 kept the whole machine open at once (71 documents / 339
'  resolved components) -> SOLIDWORKS ran out of memory and crashed before
'  anything was saved.  v3 NEVER holds more than one small document (or one
'  LIGHTWEIGHT assembly "window") in memory, saves+closes after every file,
'  and pauses briefly (PAUSE_MS) so the machine can keep up.  If the run is
'  interrupted, just run the macro again - finished work is detected and
'  skipped (resume mode).
'
'  EXECUTION MODEL:
'    Phase 1  For every project file (children before parents):
'             OPEN it alone -> rename its configurations, set material and
'             custom properties, build this file's part of the "for_matlab"
'             chain (config creation + fastener suppression + child rewiring)
'             -> SAVE -> CLOSE.   One document in memory at a time.
'    Phase 2  For every assembly (deepest first) a small "window":
'             OPEN the assembly LIGHTWEIGHT -> repair child configuration
'             references -> RENAME its direct child files (official
'             RenameDocument flow; files shared by two parents are handled
'             with temporary copies) -> SAVE (SaveReferenced) -> CLOSE ALL.
'             Afterwards leftover old-name copies are deleted.
'    Phase 3  Verification: re-open every assembly lightweight, check all
'             references / configurations / suppressions, write CSV report.
'
'  HOW TO RUN:
'    1) BACK UP the whole folder C:\Users\User\Desktop\Project\SOLIDWORKS !
'    2) Close ALL documents in SOLIDWORKS (the macro refuses to start while
'       documents are open, so none of your work can be lost).
'    3) Tools > Macro > New...  -> save MODUMILL_Rename.swp
'       VBA editor: File > Import File... -> MODUMILL_RenameForSimscape.bas
'       (remove the OLD module first if you imported v1/v2 before!)
'    4) Run MainForSimscape with DRY_RUN = True  -> read the log; nothing
'       is changed.
'    5) Set DRY_RUN = False and run again.  Expect roughly 10-30 minutes.
'       Do NOT click inside SOLIDWORKS while it runs.  Progress lines
'       [n/65] and [win n/11] appear in the log file.
'
'  NOTES:
'   * Toolbox files (C:\SolidWorks Data\...) are NEVER renamed or edited.
'   * If the run is interrupted (crash / power loss), simply run the macro
'     again - it resumes and retries only what is missing.
'   * .SLDDRW drawings are NOT updated automatically - if you have drawings
'     of these parts, fix their references afterwards (Design Manager).
'   * All planned data lives in LoadData() at the bottom - easy to edit.
' ============================================================================

#If VBA7 Then
Private Declare PtrSafe Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' ------------------------------ USER SETTINGS ------------------------------
Public Const DRY_RUN As Boolean = True          ' <<< True = report only. Set False to execute.
Public Const TOP_ASM As String = "C:\Users\User\Desktop\Project\SOLIDWORKS\MODUMILL.SLDASM"
Public Const TOP_CFG As String = "for_matlab"   ' MATLAB/Simscape configuration name
Public Const MAT_LIB As String = "solidworks materials.sldmat"
Public Const ALLOW_FALLBACK_SUPPRESS As Boolean = True  ' if for_matlab cannot be added
                                                ' (design table), suppress fasteners in
                                                ' the base config instead (modifies it!)
Public Const OPEN_LIGHTWEIGHT As Boolean = True ' assemblies open lightweight (low RAM).
                                                ' Set False if you get odd failures.
Public Const PAUSE_MS As Long = 250             ' pacing pause (ms) after every file/window
Public Const CLEAN_LEFTOVER_OLD As Boolean = True ' delete leftover old-name files at the end
Public Const LOG_PATH As String = "C:\Users\User\Desktop\Project\SOLIDWORKS\_SW_Rename_ForSimscape_Log.txt"
Public Const REPORT_CSV As String = "C:\Users\User\Desktop\Project\SOLIDWORKS\_Simscape_Component_Report.csv"

' ---------------------------- EXPECTED COUNTS ------------------------------
Private Const EXP_INSTANCES As Long = 339
Private Const EXP_RENAMES As Long = 64
Private Const EXP_CFG_PAIRS As Long = 172
Private Const EXP_MAT_SET As Long = 23
Private Const EXP_FASTENERS As Long = 162
Private Const EXP_NEEDS_REVIEW As Long = 27

' ------------------------------ TYPES / STATE ------------------------------
Private Type FileRec
    Id As Long
    Action As String          ' RENAME / KEEP / SKIP_TOOLBOX
    DocType As String         ' Part / Assembly
    Depth As Long
    OldPath As String
    NewPath As String
    MatAction As String       ' SET / KEEP / NONE / SKIP_TOOLBOX
    MatTarget As String
    CfgPolicy As String       ' RENAME / NONE / SKIP_DESIGN_TABLE_SUSPECT / SKIP_TOOLBOX
    NeedsReview As Boolean
    IsToolbox As Boolean
    IsLayout As Boolean
    PropNames(1 To 6) As String
    PropValues(1 To 6) As String
    Renamed As Boolean
    UsedFallback As Boolean
End Type

Private Type CfgRec
    FileId As Long
    OldCfg As String
    NewCfg As String
    Done As Boolean
    Skipped As Boolean
End Type

Private Type FixRec
    AssmStem As String
    CompStem As String
    FromCfg As String
    ToCfg As String
    FromCfgNew As String
    ToCfgNew As String
End Type

Private Files() As FileRec
Private Cfgs() As CfgRec
Private Fixes() As FixRec
Private OrderIds() As Long        ' phase 1 order: children before parents
Private Wins() As Long            ' phase 2 window order: deepest assembly first
Private PairPar() As Long         ' parent/child pairs (file ids)
Private PairChild() As Long
Private gNFiles As Long, gNCfgs As Long, gNFixes As Long
Private gNOrder As Long, gNWins As Long, gNPairs As Long
Private Fasteners As Collection   ' lowercase model-name stems to suppress
Private ChainAsmys As Collection  ' lowercase stems of assemblies on the for_matlab chain

Private swApp As SldWorks.SldWorks
Private mIdByOld As Object        ' lcase(oldPath) -> fileId
Private mIdByNew As Object        ' lcase(newPath) -> fileId
Private mCfgMap As Object         ' fileId -> Dictionary(lcase(oldCfg) -> newCfg)
Private mToolbox As Object        ' lcase(toolbox path) -> True
Private gChainOk As Object        ' lcase(oldStem) -> True when TOP_CFG is present
Private gWinOk As Object          ' window fileId -> True when its window finished cleanly
Private gRenamedInWindow As Collection
Private gTmpKills As Collection

Private gLogNum As Integer
Private gTopFid As Long
Private gVerifyCount As Long

Private nCfgRenamed As Long, nCfgSkipped As Long
Private nFixApplied As Long
Private nMatSet As Long, nMatFail As Long
Private nPropSet As Long
Private nCfgCreated As Long, nRefRewired As Long, nRefRepaired As Long
Private nSuppressed As Long, nSuppFail As Long
Private nRenamed As Long, nRenameFail As Long
Private nMismatch As Long, nSuppVerified As Long, nOldLeft As Long, nStubsCleaned As Long
Private nDocsEdited As Long, nWindowsDone As Long
Private nWarnings As Long

' --------------------------------- MAIN ------------------------------------
Public Sub MainForSimscape()
    On Error GoTo MainErr

    ResetCounters
    OpenLog
    LogLine "=============================================================="
    LogLine "MODUMILL rename & MATLAB/Simscape preparation (macro v3 - low memory, file by file)"
    LogLine "Started: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    If DRY_RUN Then
        LogLine "*** DRY RUN - NOTHING will be modified ***"
    Else
        LogLine "*** LIVE RUN - documents WILL be modified, renamed and SAVED one at a time ***"
    End If

    LoadData
    LogLine "Plan: " & EXP_RENAMES & " file renames, " & EXP_CFG_PAIRS & " config renames, " & _
            EXP_MAT_SET & " files to receive a material, " & EXP_FASTENERS & " fastener instances to suppress."

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        LogLine "FATAL: cannot get the SOLIDWORKS application object."
        GoTo MainEnd
    End If

    ' refuse to run with documents open - the macro opens/closes files itself
    Dim vD As Variant, nOpen As Long
    On Error Resume Next
    vD = swApp.GetDocuments
    On Error GoTo 0
    nOpen = 0
    If IsArray(vD) Then nOpen = UBound(vD) - LBound(vD) + 1
    If nOpen > 0 Then
        LogWarn nOpen & " document(s) are open - macro aborted. Close them all and run again."
        MsgBox "Please CLOSE ALL open documents in SOLIDWORKS first, then run the macro again." & vbCrLf & vbCrLf & _
               "The macro opens, edits, saves and closes every file by itself (low-memory mode).", _
               vbExclamation, "MODUMILL Simscape prep v3"
        GoTo MainEnd
    End If

    If Not DRY_RUN Then
        Dim ans As Long
        ans = MsgBox("LIVE RUN: every project file will be edited, renamed and SAVED (one at a time)." & vbCrLf & _
                     "This takes roughly 10-30 minutes - do not close SOLIDWORKS while it runs." & vbCrLf & vbCrLf & _
                     "Continue?", vbYesNo + vbQuestion, "MODUMILL Simscape prep v3")
        If ans <> vbYes Then
            LogLine "User cancelled the live run."
            GoTo MainEnd
        End If
    End If

    InitLookups
    LogLine "Execution: " & gNOrder & " documents one-by-one (phase 1), " & gNWins & _
            " rename windows (phase 2), pause " & PAUSE_MS & " ms between steps, lightweight open=" & OPEN_LIGHTWEIGHT

    LogLine "Names flagged NEEDS REVIEW in the plan (inferred - verify!): " & EXP_NEEDS_REVIEW
    Dim i As Long
    For i = 1 To gNFiles
        If Files(i).NeedsReview Then
            LogLine "  REVIEW: " & Files(i).OldPath & "  ->  " & Files(i).NewPath
        End If
    Next i

    Phase1_EditEachDocument
    Phase2_RenameWindows
    Phase3_VerifyAndReport
    WriteSummary
    GoTo MainEnd

MainErr:
    LogLine "FATAL ERROR " & Err.Number & ": " & Err.Description
    MsgBox "Macro stopped with error " & Err.Number & ":" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Everything finished so far is SAVED on disk." & vbCrLf & _
           "You can simply run the macro again - it resumes where it stopped." & vbCrLf & _
           "Log: " & LOG_PATH, vbCritical, "MODUMILL Simscape prep v3"
MainEnd:
    CloseAllDocs
    LogLine "Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    CloseLog
End Sub

Private Sub ResetCounters()
    gVerifyCount = 0
    nCfgRenamed = 0: nCfgSkipped = 0: nFixApplied = 0
    nMatSet = 0: nMatFail = 0: nPropSet = 0
    nCfgCreated = 0: nRefRewired = 0: nRefRepaired = 0
    nSuppressed = 0: nSuppFail = 0
    nRenamed = 0: nRenameFail = 0
    nMismatch = 0: nSuppVerified = 0: nOldLeft = 0: nStubsCleaned = 0
    nDocsEdited = 0: nWindowsDone = 0: nWarnings = 0
End Sub

Private Sub InitLookups()
    Dim i As Long, j As Long, d As Object
    Set mIdByOld = CreateObject("Scripting.Dictionary"): mIdByOld.CompareMode = 1
    Set mIdByNew = CreateObject("Scripting.Dictionary"): mIdByNew.CompareMode = 1
    Set mCfgMap = CreateObject("Scripting.Dictionary"): mCfgMap.CompareMode = 1
    Set mToolbox = CreateObject("Scripting.Dictionary"): mToolbox.CompareMode = 1
    Set gChainOk = CreateObject("Scripting.Dictionary"): gChainOk.CompareMode = 1
    Set gWinOk = CreateObject("Scripting.Dictionary"): gWinOk.CompareMode = 1
    For i = 1 To gNFiles
        mIdByOld.Item(LCase(Files(i).OldPath)) = i
        If LCase(Files(i).NewPath) <> LCase(Files(i).OldPath) Then
            mIdByNew.Item(LCase(Files(i).NewPath)) = i
        End If
        If Files(i).Action = "SKIP_TOOLBOX" Then
            mToolbox.Item(LCase(Files(i).OldPath)) = True
        End If
    Next i
    For j = 1 To gNCfgs
        i = Cfgs(j).FileId
        If Not mCfgMap.Exists(i) Then
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            mCfgMap.Add i, d
        End If
        mCfgMap.Item(i).Item(LCase(Cfgs(j).OldCfg)) = Cfgs(j).NewCfg
    Next j
    gTopFid = 0
    If mIdByOld.Exists(LCase(TOP_ASM)) Then gTopFid = CLng(mIdByOld.Item(LCase(TOP_ASM)))
End Sub

' ------------------------------- PHASE 1 ------------------------------------
' Every project file is opened ALONE, edited, saved and closed again.
Private Sub Phase1_EditEachDocument()
    LogLine "--- Phase 1: editing every file one at a time (open -> edit -> save -> close) ---"
    Dim k As Long
    For k = 1 To gNOrder
        EditOneDocument OrderIds(k), k
        CloseAllDocs
        Pace
    Next k
    LogLine "Phase 1 done: documents processed=" & nDocsEdited & " / " & gNOrder & _
            ", cfg renames=" & nCfgRenamed & " (skipped " & nCfgSkipped & ")" & _
            ", materials=" & nMatSet & " (fail " & nMatFail & "), props=" & nPropSet & _
            ", '" & TOP_CFG & "' created=" & nCfgCreated & ", chain refs=" & nRefRewired & _
            ", fasteners suppressed=" & nSuppressed & " (fail " & nSuppFail & ")" & DryTag
End Sub

Private Sub EditOneDocument(ByVal fid As Long, ByVal k As Long)
    Dim eff As String, doc As SldWorks.ModelDoc2, tag As String
    On Error GoTo EH
    tag = "[" & k & "/" & gNOrder & "] "
    If Files(fid).Action = "SKIP_TOOLBOX" Then Exit Sub
    eff = EffectivePath(Files(fid).OldPath, Files(fid).NewPath)
    If Len(eff) = 0 Then
        LogWarn tag & "file not found on disk (old and new missing) - skipped: " & Files(fid).OldPath
        Exit Sub
    End If
    If LCase(eff) <> LCase(Files(fid).OldPath) Then
        LogLine tag & "resume: already renamed on disk - editing " & eff
    End If
    Set doc = OpenSmart(eff)
    If doc Is Nothing Then
        LogWarn tag & "cannot open - skipped: " & eff
        Exit Sub
    End If
    ActivateDoc doc
    LogLine tag & "edit: " & eff & "  (" & Files(fid).DocType & ", depth " & Files(fid).Depth & ")  " & Format(Now, "hh:nn:ss")

    ' 1) configuration renames (localized -> ASCII)
    ApplyCfgRenamesFor fid, doc
    ' 2) for_matlab chain work for this assembly (children were processed before)
    If doc.GetType = swDocASSEMBLY Then ChainWorkForDoc fid, doc, eff
    ' 3) component configuration fixes (after chain work so for_matlab is included)
    ApplyFixesForDoc fid, doc
    ' 4) material + custom properties
    ApplyMaterial Files(fid), doc
    ApplyProps Files(fid), doc
    ' 5) save + (caller closes)
    If DRY_RUN Then
        LogLine tag & "DRY: closing without saving (no changes written)"
    Else
        SaveDoc doc, tag
    End If
    nDocsEdited = nDocsEdited + 1
    Exit Sub
EH:
    LogWarn tag & "error " & Err.Number & ": " & Err.Description & " - continuing with next file"
    Err.Clear
End Sub

Private Sub ApplyCfgRenamesFor(ByVal fid As Long, ByVal doc As SldWorks.ModelDoc2)
    Dim i As Long, dtChecked As Boolean, hasDT As Boolean
    For i = 1 To gNCfgs
        If Cfgs(i).FileId = fid Then
            If Not dtChecked Then
                dtChecked = True
                hasDT = DocHasDesignTable(doc)
                If hasDT Then
                    LogLine "CFG: design table detected - skipping " & CountCfgsFor(fid) & _
                            " config rename(s) for " & StemOf(doc.GetPathName) & _
                            " (edit the design table manually if ASCII names are required)"
                ElseIf Files(fid).CfgPolicy = "SKIP_DESIGN_TABLE_SUSPECT" Then
                    LogLine "CFG: no design table at runtime - applying renames for " & StemOf(doc.GetPathName)
                End If
            End If
            If hasDT Then
                Cfgs(i).Skipped = True
                nCfgSkipped = nCfgSkipped + 1
            Else
                DoCfgRename i, doc
            End If
        End If
    Next i
End Sub

Private Sub DoCfgRename(ByVal i As Long, ByVal doc As SldWorks.ModelDoc2)
    Dim swCfg As SldWorks.Configuration
    Dim oldN As String, newN As String, docStem As String
    oldN = Cfgs(i).OldCfg
    newN = Cfgs(i).NewCfg
    docStem = StemOf(doc.GetPathName)
    On Error GoTo EH
    Set swCfg = doc.GetConfigurationByName(oldN)
    If swCfg Is Nothing Then
        If ConfigExists(doc, newN) Then
            Cfgs(i).Done = True
            LogLine "CFG: already renamed: '" & oldN & "' -> '" & newN & "' (" & docStem & ")"
        Else
            Cfgs(i).Skipped = True
            LogWarn "CFG: configuration not found: '" & oldN & "' in " & docStem
        End If
        Exit Sub
    End If
    If DRY_RUN Then
        LogLine "CFG [DRY]: would rename '" & oldN & "' -> '" & newN & "' in " & docStem
        Cfgs(i).Done = True
        nCfgRenamed = nCfgRenamed + 1
        Exit Sub
    End If
    If ConfigExists(doc, newN) Then
        Cfgs(i).Skipped = True
        LogWarn "CFG: target name already exists - skipped: '" & newN & "' in " & docStem
        Exit Sub
    End If
    swCfg.Name = newN
    If swCfg.Name = newN Then
        Cfgs(i).Done = True
        nCfgRenamed = nCfgRenamed + 1
        LogLine "CFG: renamed '" & oldN & "' -> '" & newN & "' in " & docStem
    Else
        Cfgs(i).Skipped = True
        LogWarn "CFG: rename did not take effect: '" & oldN & "' in " & docStem
    End If
    Exit Sub
EH:
    Cfgs(i).Skipped = True
    LogWarn "CFG: rename failed '" & oldN & "' -> '" & newN & "' in " & docStem & ": " & Err.Description
End Sub

Private Function CountCfgsFor(ByVal fileId As Long) As Long
    Dim i As Long
    For i = 1 To gNCfgs
        If Cfgs(i).FileId = fileId Then CountCfgsFor = CountCfgsFor + 1
    Next i
End Function

' ---- for_matlab chain work for a single assembly document ------------------
Private Sub ChainWorkForDoc(ByVal fid As Long, ByVal doc As SldWorks.ModelDoc2, ByVal eff As String)
    Dim isTop As Boolean, isChain As Boolean, oldStem As String, origCfg As String
    Dim derive As String, swC As SldWorks.Configuration
    oldStem = LCase(StemOf(Files(fid).OldPath))
    isTop = (fid = gTopFid)
    isChain = IsChainAssembly(oldStem)
    If Not (isTop Or isChain) Then Exit Sub   ' non-chain assembly: nothing to do here
    On Error Resume Next
    origCfg = doc.ConfigurationManager.ActiveConfiguration.Name
    On Error GoTo 0
    derive = DeriveBaseConfig(doc, origCfg)

    If ConfigExists(doc, TOP_CFG) Then
        gChainOk.Item(oldStem) = True
        LogLine "CHAIN: '" & TOP_CFG & "' already exists on " & StemOf(eff)
    ElseIf DRY_RUN Then
        gChainOk.Item(oldStem) = True
        LogLine "CHAIN [DRY]: would add '" & TOP_CFG & "' to " & StemOf(eff) & " (derived from '" & derive & "')"
    Else
        If DocHasDesignTable(doc) Then
            LogWarn "CHAIN: " & StemOf(eff) & " is design-table driven - AddConfiguration may be rejected"
        End If
        On Error Resume Next
        Set swC = doc.ConfigurationManager.AddConfiguration(TOP_CFG, _
                "MATLAB/Simscape export configuration", "", 0, derive, "")
        If Err.Number <> 0 Then
            LogWarn "CHAIN: AddConfiguration raised " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        If swC Is Nothing Then
            LogWarn "CHAIN: could not add '" & TOP_CFG & "' to " & StemOf(eff)
            If isChain And ALLOW_FALLBACK_SUPPRESS Then
                LogWarn "CHAIN: FALLBACK - suppressing fasteners inside '" & derive & "' of " & _
                        StemOf(eff) & " (design config is modified!)"
                If Len(derive) > 0 Then
                    If doc.ShowConfiguration2(derive) Then RewireAndSuppress doc, derive, False
                End If
            End If
        Else
            nCfgCreated = nCfgCreated + 1
            gChainOk.Item(oldStem) = True
            LogLine "CHAIN: added '" & TOP_CFG & "' to " & StemOf(eff) & " (derived from '" & derive & "')"
        End If
    End If

    If ConfigExists(doc, TOP_CFG) Then
        On Error Resume Next
        If doc.ShowConfiguration2(TOP_CFG) Then
            RewireAndSuppress doc, TOP_CFG, True
        Else
            LogWarn "CHAIN: cannot activate '" & TOP_CFG & "' on " & StemOf(eff)
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    ElseIf DRY_RUN And gChainOk.Exists(oldStem) Then
        ' dry preview: TOP_CFG is not on disk yet - scan the base config
        ' read-only and log what would be suppressed / rewired inside it
        If Len(derive) > 0 Then
            If doc.ShowConfiguration2(derive) Then RewireAndSuppress doc, derive, True
        End If
    End If

    If Len(origCfg) > 0 Then
        On Error Resume Next
        doc.ShowConfiguration2 origCfg
        On Error GoTo 0
    End If
End Sub

Private Function DeriveBaseConfig(ByVal doc As SldWorks.ModelDoc2, ByVal origCfg As String) As String
    Dim vNames As Variant
    DeriveBaseConfig = ""
    If ConfigExists(doc, "Default") Then
        DeriveBaseConfig = "Default"
        Exit Function
    End If
    On Error Resume Next
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        If UBound(vNames) >= LBound(vNames) Then DeriveBaseConfig = CStr(vNames(LBound(vNames)))
    End If
    On Error GoTo 0
    If Len(DeriveBaseConfig) = 0 Then DeriveBaseConfig = origCfg
End Function

' useCfg must be the ACTIVE configuration when this runs.
' doRewire=True  -> chain children get ReferencedConfiguration = TOP_CFG
Private Sub RewireAndSuppress(ByVal doc As SldWorks.ModelDoc2, ByVal useCfg As String, ByVal doRewire As Boolean)
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    Dim cStem As String, cfid As Long, dStem As String
    dStem = StemOf(doc.GetPathName)
    On Error GoTo EH
    Set swAssy = doc
    vComps = swAssy.GetComponents(True)
    If Not IsArray(vComps) Then Exit Sub
    For Each v In vComps
        Set swComp = v
        cfid = FileIdByPath(swComp.GetPathName)
        cStem = ChildOldStem(cfid, swComp.GetPathName)
        If IsFastenerStem(cStem) Then
            SuppressComp swComp, useCfg, dStem
        ElseIf doRewire Then
            If gChainOk.Exists(cStem) Then
                If StrComp(swComp.ReferencedConfiguration, TOP_CFG, vbTextCompare) <> 0 Then
                    If DRY_RUN Then
                        nRefRewired = nRefRewired + 1
                        LogLine "CHAIN [DRY]: would rewire " & swComp.Name2 & " (in " & dStem & ") -> '" & TOP_CFG & "'"
                    Else
                        On Error Resume Next
                        swComp.ReferencedConfiguration = TOP_CFG
                        If Err.Number <> 0 Then
                            LogWarn "CHAIN: rewire failed for " & swComp.Name2 & ": " & Err.Description
                            Err.Clear
                        ElseIf StrComp(swComp.ReferencedConfiguration, TOP_CFG, vbTextCompare) = 0 Then
                            nRefRewired = nRefRewired + 1
                            LogLine "CHAIN: rewired " & swComp.Name2 & " (in " & dStem & ") -> '" & TOP_CFG & "'"
                        Else
                            LogWarn "CHAIN: rewire did not stick for " & swComp.Name2
                        End If
                        On Error GoTo EH
                    End If
                End If
            ElseIf IsChainAssembly(cStem) Then
                LogWarn "CHAIN: child " & swComp.Name2 & " is a chain assembly WITHOUT '" & TOP_CFG & _
                        "' (phase 1 failed for it?) - reference not rewired"
            End If
        End If
    Next v
    Exit Sub
EH:
    LogWarn "CHAIN error in " & dStem & ": " & Err.Number & " " & Err.Description
End Sub

Private Sub SuppressComp(ByVal swComp As SldWorks.Component2, ByVal useCfg As String, ByVal dStem As String)
    Dim rc As Long
    On Error GoTo EH
    If swComp.SuppressionState = swComponentSuppressed Then
        nSuppressed = nSuppressed + 1
        LogLine "SUPPR: already suppressed: " & swComp.Name2 & " (" & dStem & ")"
        Exit Sub
    End If
    If DRY_RUN Then
        nSuppressed = nSuppressed + 1
        LogLine "SUPPR [DRY]: would suppress " & swComp.Name2 & " in '" & useCfg & "' of " & dStem
        Exit Sub
    End If
    rc = swComp.SetSuppression2(swComponentSuppressed)
    If swComp.SuppressionState = swComponentSuppressed Then
        nSuppressed = nSuppressed + 1
    Else
        nSuppFail = nSuppFail + 1
        LogWarn "SUPPR: failed for " & swComp.Name2 & " in " & dStem & " (rc=" & rc & ")"
    End If
    Exit Sub
EH:
    nSuppFail = nSuppFail + 1
    LogWarn "SUPPR error for " & swComp.Name2 & ": " & Err.Description
End Sub

' ---- component configuration fixes (GT2 pulley Bottom Base) ----------------
Private Sub ApplyFixesForDoc(ByVal fid As Long, ByVal doc As SldWorks.ModelDoc2)
    Dim i As Long, pStem As String
    If gNFixes < 1 Then Exit Sub
    If doc.GetType <> swDocASSEMBLY Then Exit Sub
    pStem = LCase(StemOf(Files(fid).OldPath))
    For i = 1 To gNFixes
        If LCase(Fixes(i).AssmStem) = pStem Then ApplyOneFixInDoc Fixes(i), doc
    Next i
End Sub

Private Sub ApplyOneFixInDoc(ByRef fx As FixRec, ByVal doc As SldWorks.ModelDoc2)
    Dim vNames As Variant, vN As Variant
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    Dim origCfg As String, cur As String
    Dim nCh As Long, cfid As Long, cStem As String
    On Error GoTo EH
    origCfg = doc.ConfigurationManager.ActiveConfiguration.Name
    vNames = doc.GetConfigurationNames
    For Each vN In vNames
        If doc.ShowConfiguration2(CStr(vN)) Then
            Set swAssy = doc
            vComps = swAssy.GetComponents(True)
            If IsArray(vComps) Then
                For Each v In vComps
                    Set swComp = v
                    cfid = FileIdByPath(swComp.GetPathName)
                    cStem = ChildOldStem(cfid, swComp.GetPathName)
                    If cStem = LCase(fx.CompStem) Then
                        cur = swComp.ReferencedConfiguration
                        If cur = fx.FromCfg Or cur = fx.FromCfgNew Then
                            If DRY_RUN Then
                                LogLine "FIX [DRY]: cfg '" & vN & "': would set '" & fx.CompStem & _
                                        "' ref '" & cur & "' -> '" & fx.ToCfgNew & "'"
                            Else
                                swComp.ReferencedConfiguration = fx.ToCfgNew
                                If swComp.ReferencedConfiguration = fx.ToCfgNew Then
                                    nCh = nCh + 1
                                    nFixApplied = nFixApplied + 1
                                    LogLine "FIX: cfg '" & vN & "': '" & fx.CompStem & "' '" & cur & "' -> '" & fx.ToCfgNew & "'"
                                Else
                                    LogWarn "FIX: could not set referenced config in '" & vN & "' of " & fx.AssmStem
                                End If
                            End If
                        End If
                    End If
                Next v
            End If
        End If
    Next vN
    On Error Resume Next
    doc.ShowConfiguration2 origCfg
    On Error GoTo 0
    LogLine "FIX: " & fx.AssmStem & " / " & fx.CompStem & ": " & nCh & " change(s)" & DryTag
    Exit Sub
EH:
    LogWarn "FIX error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    doc.ShowConfiguration2 origCfg
    On Error GoTo 0
End Sub

' ---- material + custom properties (per document) ---------------------------
Private Sub ApplyMaterial(ByRef f As FileRec, ByVal doc As SldWorks.ModelDoc2)
    Dim swPart As SldWorks.PartDoc
    Dim vNames As Variant, vN As Variant
    Dim got As String, libOut As String
    If f.MatAction <> "SET" Then Exit Sub
    If Len(f.MatTarget) = 0 Then Exit Sub
    If doc.GetType <> swDocPART Then
        LogLine "MAT: not a part - material skipped: " & StemOf(doc.GetPathName)
        Exit Sub
    End If
    Set swPart = doc
    On Error GoTo EH
    vNames = doc.GetConfigurationNames
    For Each vN In vNames
        If DRY_RUN Then
            LogLine "MAT [DRY]: would set '" & f.MatTarget & "' on " & StemOf(doc.GetPathName) & " [" & vN & "]"
        Else
            swPart.SetMaterialPropertyName2 CStr(vN), MAT_LIB, f.MatTarget
            got = ReadbackMaterial(doc, CStr(vN), libOut)
            If LCase(got) = LCase(f.MatTarget) Then
                nMatSet = nMatSet + 1
                LogLine "MAT: " & StemOf(doc.GetPathName) & " [" & vN & "] -> " & got
            Else
                nMatFail = nMatFail + 1
                LogWarn "MAT: " & StemOf(doc.GetPathName) & " [" & vN & "] expected '" & f.MatTarget & "', read back '" & got & "'"
            End If
        End If
    Next vN
    Exit Sub
EH:
    LogWarn "MAT error on " & StemOf(doc.GetPathName) & ": " & Err.Number & " " & Err.Description
End Sub

Private Function ReadbackMaterial(ByVal doc As SldWorks.ModelDoc2, ByVal cfgName As String, ByRef libOut As String) As String
    Dim oP As Object
    Set oP = doc
    On Error Resume Next
    ReadbackMaterial = oP.GetMaterialPropertyName2(cfgName, libOut)
    If Err.Number <> 0 Then
        ReadbackMaterial = "<readback failed: " & Err.Description & ">"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub ApplyProps(ByRef f As FileRec, ByVal doc As SldWorks.ModelDoc2)
    Dim cpm As Object
    Dim j As Long, rc As Long
    Dim vVal As Variant, bRes As Boolean, sLink As String
    On Error GoTo EH
    Set cpm = doc.Extension.CustomPropertyManager("")
    For j = 1 To 6
        If Len(f.PropNames(j)) > 0 And Len(f.PropValues(j)) > 0 Then
            If DRY_RUN Then
                LogLine "PROP [DRY]: would set " & f.PropNames(j) & " = '" & f.PropValues(j) & "' on " & StemOf(doc.GetPathName)
            Else
                rc = cpm.Add3(f.PropNames(j), swCustomInfoText, f.PropValues(j), swCustomPropertyDeleteAndAdd)
                vVal = Empty: bRes = False: sLink = ""
                On Error Resume Next
                rc = cpm.Get5(f.PropNames(j), False, vVal, bRes, sLink)
                If Err.Number <> 0 Then
                    LogWarn "PROP: readback failed for " & f.PropNames(j) & " on " & StemOf(doc.GetPathName) & ": " & Err.Description
                    Err.Clear
                    vVal = f.PropValues(j)   ' trust Add3
                End If
                On Error GoTo EH
                If CStr(vVal) = f.PropValues(j) Then
                    nPropSet = nPropSet + 1
                Else
                    LogWarn "PROP: readback mismatch " & f.PropNames(j) & " on " & StemOf(doc.GetPathName) & ": got '" & CStr(vVal) & "'"
                End If
            End If
        End If
    Next j
    Exit Sub
EH:
    LogWarn "PROP error on " & StemOf(doc.GetPathName) & ": " & Err.Number & " " & Err.Description
End Sub

Private Sub SaveDoc(ByVal doc As SldWorks.ModelDoc2, ByVal tag As String)
    Dim st As Boolean, e As Long, w As Long
    On Error Resume Next
    st = doc.Save3(swSaveAsOptions_Silent, e, w)
    If Err.Number <> 0 Then
        LogWarn tag & "Save3 raised " & Err.Number & ": " & Err.Description
        Err.Clear
        st = False
    End If
    On Error GoTo 0
    If st Then
        LogLine tag & "saved: " & doc.GetPathName
    Else
        LogWarn tag & "SAVE FAILED for " & doc.GetPathName & " (errors=" & e & ", warnings=" & w & _
                ") - run the macro again later to redo this file"
    End If
End Sub

' ------------------------------- PHASE 2 ------------------------------------
' Renames happen in small "windows": one assembly open (lightweight) at a
' time; its direct child files are renamed with the official flow and the
' window is saved (SaveReferenced) and closed before the next one starts.
Private Sub Phase2_RenameWindows()
    LogLine "--- Phase 2: renaming files in per-assembly windows (deepest first) ---"
    Dim wi As Long
    For wi = 1 To gNWins
        WindowProc wi
        CloseAllDocs
        Pace
    Next wi
    If Not DRY_RUN Then CleanupOldFiles
    LogLine "Phase 2 done: windows=" & nWindowsDone & " / " & gNWins & ", files renamed=" & nRenamed & _
            " (failed " & nRenameFail & "), refs repaired=" & nRefRepaired & DryTag
End Sub

Private Sub WindowProc(ByVal wi As Long)
    Dim fid As Long, eff As String, tag As String
    Dim W As SldWorks.ModelDoc2
    Dim changed As Boolean, saved As Boolean
    On Error GoTo EH
    fid = Wins(wi)
    tag = "[win " & wi & "/" & gNWins & " " & StemOf(Files(fid).OldPath) & "] "
    gWinOk.Item(fid) = False
    Set gRenamedInWindow = New Collection
    Set gTmpKills = New Collection
    eff = EffectivePath(Files(fid).OldPath, Files(fid).NewPath)
    If Len(eff) = 0 Then
        LogWarn tag & "assembly not found on disk (old and new missing) - window skipped"
        Exit Sub
    End If
    Set W = OpenSmart(eff)
    If W Is Nothing Then
        LogWarn tag & "cannot open assembly - window skipped: " & eff
        Exit Sub
    End If
    ActivateDoc W
    LogLine tag & "window open (" & Format(Now, "hh:nn:ss") & "): " & eff
    changed = RepairRefsInWindow(W, fid, tag)
    If RenameChildrenInWindow(W, fid, tag) Then changed = True
    If DRY_RUN Then
        If changed Then LogLine tag & "DRY: window changes would be saved here - closing without saving"
        gWinOk.Item(fid) = True
    ElseIf changed Then
        saved = SaveWindow(W, tag)
        gWinOk.Item(fid) = saved
    Else
        LogLine tag & "nothing to change in this window"
        gWinOk.Item(fid) = True
    End If
    nWindowsDone = nWindowsDone + 1
    CloseAllDocs
    If Not DRY_RUN Then
        If CBool(gWinOk.Item(fid)) Then MakeStubsForSharedChildren fid, tag
    End If
    Exit Sub
EH:
    LogWarn tag & "window error " & Err.Number & ": " & Err.Description & " - continuing with next window"
    Err.Clear
    RestorePendingTmp
    CloseAllDocs
End Sub

' Walk every configuration of the open window assembly and repair the
' referenced configuration of each direct child:
'   (a) chain children inside TOP_CFG  -> TOP_CFG
'   (b) component fixes (GT2 Bottom Base)
'   (c) references to renamed (localized) configuration names -> new names
Private Function RepairRefsInWindow(ByVal W As SldWorks.ModelDoc2, ByVal fid As Long, ByVal tag As String) As Boolean
    Dim vNames As Variant, vN As Variant, cfgN As String, firstCfg As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    Dim cfid As Long, cur As String, nw As String, parStem As String, cStem As String
    Dim isChainPar As Boolean
    On Error GoTo EH
    parStem = LCase(StemOf(Files(fid).OldPath))
    isChainPar = IsChainAssembly(parStem) Or (fid = gTopFid)
    vNames = W.GetConfigurationNames
    If Not IsArray(vNames) Then Exit Function
    If UBound(vNames) >= LBound(vNames) Then firstCfg = CStr(vNames(LBound(vNames)))
    For Each vN In vNames
        cfgN = CStr(vN)
        If W.ShowConfiguration2(cfgN) Then
            Set swAssy = W
            vComps = swAssy.GetComponents(True)
            If IsArray(vComps) Then
                For Each v In vComps
                    Set swComp = v
                    cfid = FileIdByPath(swComp.GetPathName)
                    If cfid > 0 Then
                        cStem = LCase(StemOf(Files(cfid).OldPath))
                        cur = ""
                        On Error Resume Next
                        cur = swComp.ReferencedConfiguration
                        On Error GoTo EH
                        If isChainPar And StrComp(cfgN, TOP_CFG, vbTextCompare) = 0 And gChainOk.Exists(cStem) Then
                            If StrComp(cur, TOP_CFG, vbTextCompare) <> 0 Then
                                If SetCompRef(swComp, TOP_CFG, tag & "chain-rewire " & swComp.Name2 & " '" & cur & "' -> '" & TOP_CFG & "'") Then
                                    RepairRefsInWindow = True
                                End If
                            End If
                        Else
                            If ApplyFixRefHere(swComp, cfid, fid, cur, tag) Then
                                RepairRefsInWindow = True
                            Else
                                nw = NewCfgName(cfid, cur)
                                If Len(nw) > 0 And StrComp(nw, cur, vbBinaryCompare) <> 0 Then
                                    If SetCompRef(swComp, nw, tag & "cfg-repair " & swComp.Name2 & " '" & cur & "' -> '" & nw & "'") Then
                                        RepairRefsInWindow = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next v
            End If
        End If
    Next vN
    If Len(firstCfg) > 0 Then
        On Error Resume Next
        W.ShowConfiguration2 firstCfg
        On Error GoTo 0
    End If
    Exit Function
EH:
    LogWarn tag & "ref-repair error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Function SetCompRef(ByVal swComp As SldWorks.Component2, ByVal newName As String, ByVal logTxt As String) As Boolean
    On Error Resume Next
    If DRY_RUN Then
        LogLine "[DRY] would set ref: " & logTxt
        SetCompRef = True
        Exit Function
    End If
    swComp.ReferencedConfiguration = newName
    If Err.Number <> 0 Then
        LogWarn "REF set failed (" & logTxt & "): " & Err.Description
        Err.Clear
    ElseIf StrComp(swComp.ReferencedConfiguration, newName, vbTextCompare) = 0 Then
        SetCompRef = True
        nRefRepaired = nRefRepaired + 1
        LogLine "REF: " & logTxt
    Else
        LogWarn "REF did not stick: " & logTxt
    End If
    On Error GoTo 0
End Function

Private Function ApplyFixRefHere(ByVal swComp As SldWorks.Component2, ByVal cfid As Long, _
                                 ByVal parFid As Long, ByVal cur As String, ByVal tag As String) As Boolean
    Dim j As Long, cStem As String, pStem As String
    cStem = LCase(StemOf(Files(cfid).OldPath))
    pStem = LCase(StemOf(Files(parFid).OldPath))
    For j = 1 To gNFixes
        If LCase(Fixes(j).AssmStem) = pStem And LCase(Fixes(j).CompStem) = cStem Then
            If Len(cur) > 0 Then
                If StrComp(cur, Fixes(j).FromCfg, vbTextCompare) = 0 Or StrComp(cur, Fixes(j).FromCfgNew, vbTextCompare) = 0 Then
                    If DRY_RUN Then
                        LogLine tag & "[DRY] FIX would set '" & Fixes(j).CompStem & "' ref '" & cur & "' -> '" & Fixes(j).ToCfgNew & "'"
                        ApplyFixRefHere = True
                    Else
                        ApplyFixRefHere = SetCompRef(swComp, Fixes(j).ToCfgNew, _
                                tag & "fix " & Fixes(j).CompStem & " '" & cur & "' -> '" & Fixes(j).ToCfgNew & "'")
                    End If
                End If
            End If
        End If
    Next j
End Function

Private Function RenameChildrenInWindow(ByVal W As SldWorks.ModelDoc2, ByVal fid As Long, ByVal tag As String) As Boolean
    Dim pi As Long, cid As Long
    Dim done As Object
    Set done = CreateObject("Scripting.Dictionary")
    For pi = 1 To gNPairs
        If PairPar(pi) = fid Then
            cid = PairChild(pi)
            If Files(cid).Action = "RENAME" Then
                If Not done.Exists(cid) Then
                    done.Add cid, True
                    If RenameOneChild(W, cid, tag) Then RenameChildrenInWindow = True
                End If
            End If
        End If
    Next pi
End Function

Private Function RenameOneChild(ByVal W As SldWorks.ModelDoc2, ByVal cid As Long, ByVal tag As String) As Boolean
    Dim oldP As String, newP As String, compName As String, compNameNew As String
    Dim oldE As Boolean, newE As Boolean, tmpP As String, okRen As Boolean
    On Error GoTo EH
    oldP = Files(cid).OldPath
    newP = Files(cid).NewPath
    oldE = FileExists(oldP)
    newE = FileExists(newP)
    compName = FindCompNameByPath(W, LCase(oldP))
    compNameNew = FindCompNameByPath(W, LCase(newP))

    If DRY_RUN Then
        If newE And Not oldE Then
            Files(cid).Renamed = True
            LogLine tag & "[DRY] child already renamed on disk: " & StemOf(newP)
        ElseIf oldE Then
            nRenamed = nRenamed + 1
            RenameOneChild = True
            LogLine tag & "[DRY] would rename: " & StemOf(oldP) & "  ->  " & StemOf(newP) & _
                    IIf(newE, "   (target exists - temporary-copy procedure)", "")
        Else
            LogWarn tag & "[DRY] child file missing on disk: " & oldP
        End If
        Exit Function
    End If

    ' ---- resume awareness -------------------------------------------------
    If newE And Not oldE Then
        If Len(compNameNew) > 0 Then
            Files(cid).Renamed = True
            LogLine tag & "child already renamed & referenced: " & StemOf(newP)
        Else
            nRenameFail = nRenameFail + 1
            LogWarn tag & "PROBLEM: '" & StemOf(newP) & "' exists on disk but this assembly has no component for it " & _
                    "(old reference broken?). Run the macro again; if it persists, restore this assembly from backup."
        End If
        Exit Function
    End If
    If Not oldE And Not newE Then
        nRenameFail = nRenameFail + 1
        LogWarn tag & "child file missing on disk (old and new): " & oldP
        Exit Function
    End If
    If Len(compName) = 0 Then
        If newE Then
            Files(cid).Renamed = True
            LogLine tag & "child already renamed: " & StemOf(newP)
        Else
            nRenameFail = nRenameFail + 1
            LogWarn tag & "no component found for '" & StemOf(oldP) & "' in this window - skipped"
        End If
        Exit Function
    End If

    ' ---- shared child: target file already exists -> move it aside ---------
    tmpP = ""
    If newE Then
        tmpP = newP & ".rntmp"
        On Error Resume Next
        If FileExists(tmpP) Then Kill tmpP
        Name newP As tmpP
        If Err.Number <> 0 Then
            LogWarn tag & "cannot move existing target aside (" & Err.Description & ") - child skipped: " & StemOf(newP)
            Err.Clear
            On Error GoTo 0
            nRenameFail = nRenameFail + 1
            Exit Function
        End If
        On Error GoTo 0
        LogLine tag & "shared child: existing target moved aside temporarily: " & StemOf(newP)
    End If

    ' ---- official rename flow (with fallbacks) -----------------------------
    ' order matters: the SaveAs fallback is tried BEFORE resolving a
    ' lightweight component, because resolving can temporarily use a lot
    ' of memory on big sub-assemblies.
    okRen = TryRenameDocument(W, compName, StemOf(newP), tag)
    If Not okRen Then okRen = TryRenameDocument(W, compName, newP, tag)
    If Not okRen Then okRen = TrySaveAsFallback(cid, tag)
    If Not okRen Then
        LogLine tag & "last resort: resolving component (may briefly use more memory): " & compName
        ResolveCompByName W, compName
        okRen = TryRenameDocument(W, compName, StemOf(newP), tag)
    End If

    If okRen Then
        Files(cid).Renamed = True
        nRenamed = nRenamed + 1
        RenameOneChild = True
        gRenamedInWindow.Add cid
        If Len(tmpP) > 0 Then gTmpKills.Add tmpP
        LogLine tag & "RENAMED: " & StemOf(oldP) & "  ->  " & StemOf(newP) & "  (" & Format(Now, "hh:nn:ss") & ")"
    Else
        nRenameFail = nRenameFail + 1
        LogWarn tag & "rename FAILED for " & oldP
        If Len(tmpP) > 0 Then RestoreOneTmp tmpP
    End If
    Exit Function
EH:
    nRenameFail = nRenameFail + 1
    LogWarn tag & "rename error " & Err.Number & ": " & Err.Description & " (" & Files(cid).OldPath & ")"
    Err.Clear
    If Len(tmpP) > 0 Then RestoreOneTmp tmpP
End Function

Private Function TryRenameDocument(ByVal W As SldWorks.ModelDoc2, ByVal compName As String, _
                                   ByVal newName As String, ByVal tag As String) As Boolean
    Dim renErr As Long, selOk As Boolean
    On Error GoTo EH
    W.ClearSelection2 True
    selOk = W.Extension.SelectByID2(compName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    If Not selOk Then
        LogLine tag & "SelectByID2 failed for '" & compName & "'"
        Exit Function
    End If
    renErr = -1
    On Error Resume Next
    renErr = W.Extension.RenameDocument(newName)
    If Err.Number <> 0 Then
        LogLine tag & "RenameDocument('" & newName & "') raised " & Err.Number & ": " & Err.Description
        Err.Clear
        renErr = -2
    End If
    On Error GoTo EH
    If renErr = 0 Then
        TryRenameDocument = True
        Exit Function
    End If
    LogLine tag & "RenameDocument('" & newName & "') returned " & renErr
    Exit Function
EH:
    LogLine tag & "RenameDocument error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Sub ResolveCompByName(ByVal W As SldWorks.ModelDoc2, ByVal compName As String)
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    On Error Resume Next
    Set swAssy = W
    vComps = swAssy.GetComponents(True)
    If IsArray(vComps) Then
        For Each v In vComps
            Set swComp = v
            If swComp.Name2 = compName Then
                If swComp.SuppressionState = swComponentLightweight Then
                    LogLine "REN: resolving lightweight component before rename: " & compName
                    swComp.SetSuppressionState2 swComponentFullyResolved
                    DoEvents
                    If PAUSE_MS > 0 Then ApiSleep PAUSE_MS
                End If
                Exit For
            End If
        Next v
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function TrySaveAsFallback(ByVal cid As Long, ByVal tag As String) As Boolean
    Dim childDoc As SldWorks.ModelDoc2, bRet As Boolean, e As Long, w2 As Long
    On Error Resume Next
    Set childDoc = OpenSmart(Files(cid).OldPath)
    If childDoc Is Nothing Then
        LogWarn tag & "SaveAs fallback: cannot open child: " & Files(cid).OldPath
        Exit Function
    End If
    bRet = childDoc.Extension.SaveAs(Files(cid).NewPath, swSaveAsCurrentVersion, _
                                     swSaveAsOptions_Silent, Nothing, e, w2)
    If Err.Number <> 0 Then
        LogWarn tag & "SaveAs fallback raised " & Err.Number & ": " & Err.Description
        Err.Clear
        bRet = False
    End If
    On Error GoTo 0
    If bRet Then
        Files(cid).UsedFallback = True
        LogLine tag & "SaveAs fallback OK: " & StemOf(Files(cid).OldPath) & " -> " & StemOf(Files(cid).NewPath) & _
                " (old file may remain on disk; cleanup will remove it)"
        TrySaveAsFallback = True
    Else
        LogWarn tag & "SaveAs fallback FAILED (errors=" & e & ", warnings=" & w2 & ")"
    End If
End Function

Private Function SaveWindow(ByVal W As SldWorks.ModelDoc2, ByVal tag As String) As Boolean
    Dim st As Boolean, e As Long, w2 As Long
    On Error Resume Next
    st = W.Save3(swSaveAsOptions_Silent + swSaveAsOptions_SaveReferenced, e, w2)
    If Err.Number <> 0 Then
        LogWarn tag & "Save3(SaveReferenced) raised " & Err.Number & ": " & Err.Description
        Err.Clear
        st = False
    End If
    On Error GoTo 0
    If st Then
        LogLine tag & "window saved (renames committed): " & W.GetPathName
        KillPendingTmp
        SaveWindow = True
        Exit Function
    End If
    If gRenamedInWindow.Count > 0 Then
        ' never save the window alone while renames are pending - that could
        ' write references to files that were not written. Discard instead:
        LogWarn tag & "SAVE FAILED with pending renames - window discarded (nothing written to disk). Run the macro again."
        RestorePendingTmp
        SaveWindow = False
        Exit Function
    End If
    LogWarn tag & "SaveReferenced failed (errors=" & e & ", warnings=" & w2 & ") - retrying plain save (ref repairs only)"
    On Error Resume Next
    st = W.Save3(swSaveAsOptions_Silent, e, w2)
    If Err.Number <> 0 Then
        LogWarn tag & "plain save raised " & Err.Number & ": " & Err.Description
        Err.Clear
        st = False
    End If
    On Error GoTo 0
    If st Then
        LogLine tag & "plain save OK: " & W.GetPathName
        SaveWindow = True
    Else
        LogWarn tag & "SAVE FAILED - window discarded. Run the macro again."
        SaveWindow = False
    End If
End Function

Private Sub KillPendingTmp()
    Dim v As Variant
    On Error Resume Next
    For Each v In gTmpKills
        If FileExists(CStr(v)) Then
            Kill CStr(v)
            If Err.Number <> 0 Then
                LogWarn "TMP: could not delete " & CStr(v) & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next v
    On Error GoTo 0
    Set gTmpKills = New Collection
End Sub

Private Sub RestorePendingTmp()
    Dim i As Long
    If gTmpKills Is Nothing Then Exit Sub
    For i = gTmpKills.Count To 1 Step -1
        RestoreOneTmp CStr(gTmpKills(i))
    Next i
End Sub

Private Sub RestoreOneTmp(ByVal t As String)
    Dim orig As String, i As Long
    If Len(t) = 0 Then Exit Sub
    On Error Resume Next
    orig = Left$(t, Len(t) - Len(".rntmp"))
    If FileExists(t) And Not FileExists(orig) Then
        Name t As orig
        If Err.Number <> 0 Then
            LogWarn "TMP: could not restore " & orig & ": " & Err.Description & " - rename it back manually!"
            Err.Clear
        Else
            LogLine "TMP: restored " & orig
        End If
    End If
    If Not gTmpKills Is Nothing Then
        For i = gTmpKills.Count To 1 Step -1
            If StrComp(CStr(gTmpKills(i)), t, vbTextCompare) = 0 Then gTmpKills.Remove i
        Next i
    End If
    On Error GoTo 0
End Sub

' A child used by TWO parent assemblies is renamed inside the first window;
' a temporary copy under the old name lets the second parent open cleanly
' and consume (re-rename) it in its own window.
Private Sub MakeStubsForSharedChildren(ByVal fid As Long, ByVal tag As String)
    Dim v As Variant, cid As Long
    If gRenamedInWindow Is Nothing Then Exit Sub
    For Each v In gRenamedInWindow
        cid = CLng(v)
        If CountParents(cid) > 1 Then
            If FileExists(Files(cid).NewPath) And Not FileExists(Files(cid).OldPath) Then
                On Error Resume Next
                FileCopy Files(cid).NewPath, Files(cid).OldPath
                If Err.Number <> 0 Then
                    LogWarn tag & "STUB: could not create temporary copy for the other parent(s): " & Err.Description
                    Err.Clear
                Else
                    LogLine tag & "STUB: temporary copy left for the other parent(s): " & StemOf(Files(cid).OldPath)
                End If
                On Error GoTo 0
            End If
        End If
    Next v
End Sub

Private Function CountParents(ByVal cid As Long) As Long
    Dim pi As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For pi = 1 To gNPairs
        If PairChild(pi) = cid Then
            If Not seen.Exists(PairPar(pi)) Then seen.Add PairPar(pi), True
        End If
    Next pi
    CountParents = seen.Count
End Function

Private Sub CleanupOldFiles()
    LogLine "--- Phase 2b: cleanup of leftover old-name files ---"
    Dim i As Long, f As String, pd As String
    For i = 1 To gNFiles
        If Files(i).Action = "RENAME" Then
            If FileExists(Files(i).NewPath) Then
                If FileExists(Files(i).OldPath) Then
                    If CLEAN_LEFTOVER_OLD And AllParentsOk(i) Then
                        On Error Resume Next
                        Kill Files(i).OldPath
                        If Err.Number <> 0 Then
                            LogWarn "CLEAN: could not delete " & Files(i).OldPath & ": " & Err.Description
                            Err.Clear
                            nOldLeft = nOldLeft + 1
                        Else
                            LogLine "CLEAN: removed leftover old file: " & StemOf(Files(i).OldPath)
                            nStubsCleaned = nStubsCleaned + 1
                        End If
                        On Error GoTo 0
                    Else
                        nOldLeft = nOldLeft + 1
                        LogLine "NOTE: old file KEPT for review (a parent window failed, or CLEAN_LEFTOVER_OLD=False): " & Files(i).OldPath
                    End If
                End If
            Else
                LogWarn "CLEAN: renamed file missing on disk: " & Files(i).NewPath
            End If
        End If
    Next i
    pd = ProjDir()
    On Error Resume Next
    f = Dir(pd & "\*.rntmp")
    Do While Len(f) > 0
        Kill pd & "\" & f
        LogLine "CLEAN: removed temp file: " & f
        f = Dir()
    Loop
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Function AllParentsOk(ByVal cid As Long) As Boolean
    Dim pi As Long
    AllParentsOk = True
    For pi = 1 To gNPairs
        If PairChild(pi) = cid Then
            If Not gWinOk.Exists(PairPar(pi)) Then
                AllParentsOk = False
                Exit Function
            End If
            If Not CBool(gWinOk.Item(PairPar(pi))) Then
                AllParentsOk = False
                Exit Function
            End If
        End If
    Next pi
End Function

' ------------------------------- PHASE 3 ------------------------------------
Private Sub Phase3_VerifyAndReport()
    LogLine "--- Phase 3: verification + CSV report ---"
    If DRY_RUN Then
        LogLine "VERIFY [DRY]: skipped - nothing was changed on disk."
        Exit Sub
    End If
    Dim i As Long, okRen As Long
    For i = 1 To gNFiles
        If Files(i).Action = "RENAME" Then
            If FileExists(Files(i).NewPath) Then
                okRen = okRen + 1
            Else
                LogWarn "VERIFY: renamed file missing on disk: " & Files(i).NewPath
            End If
        End If
    Next i
    LogLine "VERIFY: renamed files present on disk: " & okRen & " / " & EXP_RENAMES

    Dim csvLines As Collection
    Set csvLines = New Collection
    csvLines.Add "ParentAssembly;ScanConfig;Instance;ModelName;ModelPath;RefConfig;State;Fastener;Status"
    gVerifyCount = 0: nMismatch = 0: nSuppVerified = 0
    For i = 1 To gNFiles
        If Files(i).DocType = "Assembly" And Files(i).Action <> "SKIP_TOOLBOX" Then
            VerifyOneAssembly i, csvLines
            CloseAllDocs
            Pace
        End If
    Next i
    LogLine "VERIFY: instances scanned=" & gVerifyCount & " (audit total " & EXP_INSTANCES & _
            "), mismatches=" & nMismatch & ", suppressed fasteners=" & nSuppVerified & _
            " (expected " & EXP_FASTENERS & ")"
    If nSuppVerified <> EXP_FASTENERS Then
        LogWarn "suppressed fastener count differs from the plan (" & EXP_FASTENERS & ")"
    End If
    WriteCsv csvLines
End Sub

Private Sub VerifyOneAssembly(ByVal fid As Long, ByVal csvLines As Collection)
    Dim eff As String, doc As SldWorks.ModelDoc2, scanCfg As String, firstCfg As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant, vNames As Variant
    Dim swComp As SldWorks.Component2
    Dim p As String, cfid As Long, cStem As String, ref As String, stat As String
    Dim supp As Boolean, isChainPar As Boolean, parStem As String
    Dim j As Long
    On Error GoTo EH
    eff = Files(fid).OldPath
    If Files(fid).Action = "RENAME" Then
        If FileExists(Files(fid).NewPath) Then
            eff = Files(fid).NewPath
        End If
    End If
    If Not FileExists(eff) Then
        LogWarn "VERIFY: assembly missing on disk: " & eff
        Exit Sub
    End If
    Set doc = OpenSmart(eff)
    If doc Is Nothing Then
        LogWarn "VERIFY: cannot open " & eff
        Exit Sub
    End If
    ActivateDoc doc
    parStem = LCase(StemOf(Files(fid).OldPath))
    isChainPar = IsChainAssembly(parStem) Or (fid = gTopFid)
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        If UBound(vNames) >= LBound(vNames) Then firstCfg = CStr(vNames(LBound(vNames)))
    End If
    scanCfg = firstCfg
    If ConfigExists(doc, TOP_CFG) Then scanCfg = TOP_CFG
    On Error Resume Next
    doc.ShowConfiguration2 scanCfg
    On Error GoTo EH
    Set swAssy = doc
    vComps = swAssy.GetComponents(True)
    If Not IsArray(vComps) Then Exit Sub
    For Each v In vComps
        Set swComp = v
        p = swComp.GetPathName
        cfid = FileIdByPath(p)
        cStem = ChildOldStem(cfid, p)
        supp = False
        ref = ""
        On Error Resume Next
        supp = (swComp.SuppressionState = swComponentSuppressed)
        ref = swComp.ReferencedConfiguration
        On Error GoTo EH
        stat = "OK"
        gVerifyCount = gVerifyCount + 1
        If supp Then
            If IsFastenerStem(cStem) Then nSuppVerified = nSuppVerified + 1
        End If
        If cfid > 0 Then
            If Files(cfid).Action = "RENAME" And LCase(p) <> LCase(Files(cfid).NewPath) Then
                stat = "OLD_PATH"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: still references OLD file: " & p & "  (expected " & Files(cfid).NewPath & ")"
            ElseIf Len(NewCfgName(cfid, ref)) > 0 Then
                stat = "OLD_CFG_REF"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: component references OLD config name '" & ref & "': " & p
            End If
        Else
            If Len(p) > 0 Then
                If Not mToolbox.Exists(LCase(p)) Then
                    stat = "NOT_IN_PLAN"
                    nMismatch = nMismatch + 1
                    LogWarn "VERIFY: unexpected reference: " & p
                End If
            End If
        End If
        If StrComp(scanCfg, TOP_CFG, vbTextCompare) = 0 And isChainPar Then
            If Not supp And IsFastenerStem(cStem) And stat = "OK" Then
                stat = "FASTENER_NOT_SUPPRESSED"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: fastener NOT suppressed in '" & TOP_CFG & "': " & swComp.Name2
            End If
            If Not supp And gChainOk.Exists(cStem) And stat = "OK" Then
                If StrComp(ref, TOP_CFG, vbTextCompare) <> 0 Then
                    stat = "CHAIN_REF_MISSING"
                    nMismatch = nMismatch + 1
                    LogWarn "VERIFY: chain child not on '" & TOP_CFG & "': " & swComp.Name2 & " (ref=" & ref & ")"
                End If
            End If
        End If
        For j = 1 To gNFixes
            If LCase(Fixes(j).AssmStem) = parStem And LCase(Fixes(j).CompStem) = cStem Then
                If stat = "OK" Then
                    If StrComp(ref, Fixes(j).ToCfgNew, vbTextCompare) <> 0 And StrComp(ref, Fixes(j).ToCfg, vbTextCompare) <> 0 Then
                        stat = "FIX_MISSING"
                        nMismatch = nMismatch + 1
                        LogWarn "VERIFY: fix not applied on " & swComp.Name2 & " (ref=" & ref & ", expected " & Fixes(j).ToCfgNew & ")"
                    End If
                End If
            End If
        Next j
        csvLines.Add CsvQ(StemOf(Files(fid).OldPath)) & ";" & CsvQ(scanCfg) & ";" & CsvQ(swComp.Name2) & ";" & _
                     CsvQ(StemOf(p)) & ";" & CsvQ(p) & ";" & CsvQ(ref) & ";" & _
                     CsvQ(IIf(supp, "SUPPRESSED", "SHOWN")) & ";" & _
                     CsvQ(IIf(IsFastenerStem(cStem), "FASTENER", "")) & ";" & CsvQ(stat)
    Next v
    Exit Sub
EH:
    LogWarn "VERIFY error in " & StemOf(eff) & ": " & Err.Number & " " & Err.Description
End Sub

Private Sub WriteCsv(ByVal lines As Collection)
    Dim fn As Integer, i As Long
    On Error GoTo EH
    fn = FreeFile
    Open REPORT_CSV For Output As #fn
    For i = 1 To lines.Count
        Print #fn, lines(i)
    Next i
    Close #fn
    LogLine "CSV report written: " & REPORT_CSV & " (" & (lines.Count - 1) & " component instances)"
    Exit Sub
EH:
    LogWarn "CSV write failed: " & Err.Description
End Sub

' -------------------------------- SUMMARY -----------------------------------
Private Sub WriteSummary()
    Dim s As String
    LogLine "=================== SUMMARY (macro v3) ==================="
    LogLine "Documents processed (phase 1) : " & nDocsEdited & " / " & gNOrder
    LogLine "Configuration renames applied : " & nCfgRenamed & " (skipped: " & nCfgSkipped & ")"
    LogLine "Component config fixes        : " & nFixApplied
    LogLine "Material assignments verified : " & nMatSet & " (failed: " & nMatFail & ")"
    LogLine "Custom properties set         : " & nPropSet
    LogLine TOP_CFG & " configs created : " & nCfgCreated
    LogLine "Chain refs rewired (phase 1)  : " & nRefRewired
    LogLine "Refs repaired (phase 2)       : " & nRefRepaired
    LogLine "Fasteners suppressed          : " & nSuppressed & " (failed: " & nSuppFail & ")"
    LogLine "Rename windows processed      : " & nWindowsDone & " / " & gNWins
    LogLine "Files renamed                 : " & nRenamed & " (failed: " & nRenameFail & ")"
    If Not DRY_RUN Then
        LogLine "Leftover old files deleted    : " & nStubsCleaned & " (kept for review: " & nOldLeft & ")"
        LogLine "Verify mismatches             : " & nMismatch
    End If
    LogLine "Warnings                      : " & nWarnings
    LogLine "Log                           : " & LOG_PATH
    If nRenameFail > 0 Or nMismatch > 0 Or nSuppFail > 0 Or nMatFail > 0 Then
        LogLine ">>> Not perfect: simply RUN THE MACRO AGAIN - it resumes and retries only what is missing."
    End If
    If DRY_RUN Then
        LogLine "*** DRY RUN - no changes saved. Review this log, then set DRY_RUN = False and run again. ***"
    End If

    s = "MODUMILL Simscape prep v3" & IIf(DRY_RUN, " - DRY RUN (nothing changed)", " - COMPLETE") & vbCrLf & vbCrLf & _
        "Documents processed: " & nDocsEdited & " / " & gNOrder & vbCrLf & _
        "Config renames: " & nCfgRenamed & " (skipped " & nCfgSkipped & ")" & vbCrLf & _
        "Component fixes: " & nFixApplied & vbCrLf & _
        "Materials set: " & nMatSet & " (failed " & nMatFail & ")" & vbCrLf & _
        "Custom properties: " & nPropSet & vbCrLf & _
        TOP_CFG & " created: " & nCfgCreated & ", chain refs: " & nRefRewired & ", refs repaired: " & nRefRepaired & vbCrLf & _
        "Fasteners suppressed: " & nSuppressed & " (failed " & nSuppFail & ")" & vbCrLf & _
        "Files renamed: " & nRenamed & " (failed " & nRenameFail & ")" & vbCrLf & _
        "Warnings: " & nWarnings & vbCrLf & vbCrLf & "Log: " & LOG_PATH
    If Not DRY_RUN Then
        s = s & vbCrLf & "CSV: " & REPORT_CSV
        If nRenameFail > 0 Or nMismatch > 0 Then
            s = s & vbCrLf & vbCrLf & "Not perfect? Run the macro AGAIN - it resumes where it stopped."
        End If
    End If
    MsgBox s, vbInformation, "MODUMILL Simscape prep v3"
End Sub

' -------------------------------- HELPERS -----------------------------------
Private Sub Pace()
    On Error Resume Next
    DoEvents
    If PAUSE_MS > 0 Then ApiSleep PAUSE_MS
    DoEvents
    On Error GoTo 0
End Sub

Private Function OpenSmart(ByVal p As String) As SldWorks.ModelDoc2
    Dim t As Long, opts As Long, e As Long, w As Long
    t = DocTypeFromExt(p)
    If t = 0 Then Exit Function
    opts = swOpenDocOptions_Silent
    If OPEN_LIGHTWEIGHT And t = swDocASSEMBLY Then
        opts = swOpenDocOptions_Silent Or swOpenDocOptions_OverrideDefaultLoadLightweight Or swOpenDocOptions_LoadLightweight
    End If
    On Error Resume Next
    Set OpenSmart = swApp.OpenDoc6(p, t, opts, "", e, w)
    If Err.Number <> 0 Then
        LogWarn "OpenDoc6 raised for " & p & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If OpenSmart Is Nothing And opts <> swOpenDocOptions_Silent Then
        LogLine "OPEN: lightweight open failed for " & p & " (errors=" & e & ") - retrying normal open"
        On Error Resume Next
        Set OpenSmart = swApp.OpenDoc6(p, t, swOpenDocOptions_Silent, "", e, w)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    If OpenSmart Is Nothing Then
        LogWarn "OPEN: failed: " & p & " (errors=" & e & ", warnings=" & w & ")"
    End If
End Function

Private Sub ActivateDoc(ByVal doc As SldWorks.ModelDoc2)
    Dim e As Long
    On Error Resume Next
    swApp.ActivateDoc2 doc.GetTitle, False, e
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CloseAllDocs()
    Dim vDocs As Variant, m As Variant
    Dim md As SldWorks.ModelDoc2
    Dim t As String
    If swApp Is Nothing Then Exit Sub
    On Error Resume Next
    vDocs = swApp.GetDocuments
    If IsArray(vDocs) Then
        For Each m In vDocs
            Set md = m
            t = md.GetTitle
            swApp.CloseDoc t
            Set md = Nothing
        Next m
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function EffectivePath(ByVal oldP As String, ByVal newP As String) As String
    If FileExists(oldP) Then
        EffectivePath = oldP
        Exit Function
    End If
    If Len(newP) > 0 Then
        If FileExists(newP) Then
            EffectivePath = newP
            Exit Function
        End If
    End If
    EffectivePath = ""
End Function

Private Function FileIdByPath(ByVal p As String) As Long
    Dim lk As String
    FileIdByPath = 0
    If Len(p) = 0 Then Exit Function
    lk = LCase(p)
    If mIdByOld.Exists(lk) Then
        FileIdByPath = CLng(mIdByOld.Item(lk))
        Exit Function
    End If
    If mIdByNew.Exists(lk) Then FileIdByPath = CLng(mIdByNew.Item(lk))
End Function

Private Function ChildOldStem(ByVal cfid As Long, ByVal p As String) As String
    If cfid > 0 Then
        ChildOldStem = LCase(StemOf(Files(cfid).OldPath))
    Else
        ChildOldStem = LCase(StemOf(p))
    End If
End Function

Private Function NewCfgName(ByVal cfid As Long, ByVal cur As String) As String
    Dim d As Object
    NewCfgName = ""
    If Len(cur) = 0 Then Exit Function
    If mCfgMap Is Nothing Then Exit Function
    If mCfgMap.Exists(cfid) Then
        Set d = mCfgMap.Item(cfid)
        If d.Exists(LCase(cur)) Then NewCfgName = CStr(d.Item(LCase(cur)))
    End If
End Function

Private Function FindCompNameByPath(ByVal W As SldWorks.ModelDoc2, ByVal pathLower As String) As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    FindCompNameByPath = ""
    On Error GoTo EH
    Set swAssy = W
    vComps = swAssy.GetComponents(True)
    If Not IsArray(vComps) Then Exit Function
    For Each v In vComps
        Set swComp = v
        If LCase(swComp.GetPathName) = pathLower Then
            FindCompNameByPath = swComp.Name2
            Exit Function
        End If
    Next v
    Exit Function
EH:
    Err.Clear
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(p) <> "")
    If Err.Number <> 0 Then
        FileExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ProjDir() As String
    Dim s As String, i As Long
    s = TOP_ASM
    i = InStrRev(s, "\")
    If i > 0 Then s = Left$(s, i - 1)
    ProjDir = s
End Function

Private Function DryTag() As String
    DryTag = IIf(DRY_RUN, " (dry run - nothing changed)", "")
End Function

Private Function StemOf(ByVal p As String) As String
    Dim s As String, i As Long
    s = p
    i = InStrRev(s, "\")
    If i > 0 Then s = Mid(s, i + 1)
    i = InStrRev(s, ".")
    If i > 0 Then s = Left(s, i - 1)
    StemOf = s
End Function

Private Function DocTypeFromExt(ByVal p As String) As Long
    Dim e As String
    If Len(p) >= 7 Then e = LCase(Right(p, 7))
    If e = ".sldprt" Then
        DocTypeFromExt = swDocPART
    ElseIf e = ".sldasm" Then
        DocTypeFromExt = swDocASSEMBLY
    Else
        DocTypeFromExt = 0
    End If
End Function

Private Function DocHasDesignTable(ByVal doc As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    DocHasDesignTable = doc.Extension.HasDesignTable
    If Err.Number <> 0 Then
        DocHasDesignTable = False
        LogWarn "HasDesignTable failed (" & Err.Description & ") - assuming no design table"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ConfigExists(ByVal doc As SldWorks.ModelDoc2, ByVal name As String) As Boolean
    Dim vNames As Variant, v As Variant
    ConfigExists = False
    On Error Resume Next
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        For Each v In vNames
            If StrComp(CStr(v), name, vbTextCompare) = 0 Then
                ConfigExists = True
                Exit For
            End If
        Next v
    End If
    On Error GoTo 0
End Function

Private Function IsFastenerStem(ByVal s As String) As Boolean
    Dim v As Variant
    IsFastenerStem = False
    If Fasteners Is Nothing Then Exit Function
    For Each v In Fasteners
        If StrComp(CStr(v), s, vbTextCompare) = 0 Then
            IsFastenerStem = True
            Exit Function
        End If
    Next v
End Function

Private Function IsChainAssembly(ByVal s As String) As Boolean
    Dim v As Variant
    IsChainAssembly = False
    If ChainAsmys Is Nothing Then Exit Function
    For Each v In ChainAsmys
        If StrComp(CStr(v), s, vbTextCompare) = 0 Then
            IsChainAssembly = True
            Exit Function
        End If
    Next v
End Function

Private Function CsvQ(ByVal s As String) As String
    CsvQ = """" & Replace(s, """", """""") & """"
End Function

Private Sub OpenLog()
    On Error Resume Next
    gLogNum = FreeFile
    Open LOG_PATH For Append As #gLogNum
    If Err.Number <> 0 Then
        gLogNum = 0
        Err.Clear
        Debug.Print "Cannot open log file: " & LOG_PATH
    End If
    On Error GoTo 0
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    If gLogNum <> 0 Then
        On Error Resume Next
        Print #gLogNum, s
        On Error GoTo 0
    End If
End Sub

Private Sub LogWarn(ByVal s As String)
    nWarnings = nWarnings + 1
    LogLine "WARNING: " & s
End Sub

Private Sub CloseLog()
    On Error Resume Next
    If gLogNum <> 0 Then Close #gLogNum
    gLogNum = 0
    On Error GoTo 0
End Sub

' ============================ PLAN DATA (generated) ==========================
' Everything below comes from rename_and_material_map.json (approved plan)
' plus the component audit.  You can edit values here directly (e.g. a
' proposed name you want to change), or edit the JSON and re-run make_macro.py.
' =============================================================================
Private Sub LoadData()
    gNFiles = 71
    gNCfgs = 172
    gNFixes = 1
    gNOrder = 65
    gNWins = 11
    gNPairs = 68
    ReDim Files(1 To IIf(gNFiles > 0, gNFiles, 1))
    ReDim Cfgs(1 To IIf(gNCfgs > 0, gNCfgs, 1))
    ReDim Fixes(1 To IIf(gNFixes > 0, gNFixes, 1))
    ReDim OrderIds(1 To IIf(gNOrder > 0, gNOrder, 1))
    ReDim Wins(1 To IIf(gNWins > 0, gNWins, 1))
    ReDim PairPar(1 To IIf(gNPairs > 0, gNPairs, 1))
    ReDim PairChild(1 To IIf(gNPairs > 0, gNPairs, 1))
    Set Fasteners = New Collection
    Set ChainAsmys = New Collection

    ' --- files ---
    AddFile 1, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\ER16 d6 Collet.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_ER16_D6_Collet.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", False, False, False, _
            "MODUMILL", "COTS_SPN_ER16_D6_Collet", "ER16 collet 6 mm", "COTS", "A", "Alloy Steel"
    AddFile 2, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\GDK80-24Z_1.5.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_GDK80_24Z_Spindle_Body.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", True, False, False, _
            "MODUMILL", "COTS_SPN_GDK80_24Z_Spindle_Body", "GDK80-24Z spindle body", "COTS", "A", "Alloy Steel"
    AddFile 3, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\M3.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_GDK80_M3_Part.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", True, False, False, _
            "MODUMILL", "COTS_SPN_GDK80_M3_Part", "GDK80 spindle part 'M3' (to be confirmed)", "COTS", "A", "Alloy Steel"
    AddFile 4, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\Mini Clamp ER16.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_ER16_Clamp_Nut.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", False, False, False, _
            "MODUMILL", "COTS_SPN_ER16_Clamp_Nut", "ER16 mini clamp nut", "COTS", "A", "Alloy Steel"
    AddFile 5, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\Nozzle.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_GDK80_Nozzle.SLDPRT", _
            "SET", "Brass", "RENAME", True, False, False, _
            "MODUMILL", "COTS_SPN_GDK80_Nozzle", "GDK80 spindle nozzle", "COTS", "A", "Brass"
    AddFile 6, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\WS16J4Z.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_WS16J4Z_Part.SLDPRT", _
            "SET", "6061 Alloy", "RENAME", True, False, False, _
            "MODUMILL", "COTS_SPN_WS16J4Z_Part", "WS16J4Z spindle part (to be confirmed)", "COTS", "A", "6061 Alloy"
    AddFile 7, "RENAME", "Part", 5, "C:\Users\User\Desktop\Project\SOLIDWORKS\" & ChrW(1060) & ChrW(1088) & ChrW(1077) & ChrW(1079) & ChrW(1072) & ".SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_TOOL_EndMill_Cutter.SLDPRT", _
            "KEEP", "1.2083 (X42Cr13)", "RENAME", False, False, False, _
            "MODUMILL", "COTS_TOOL_EndMill_Cutter", "End mill cutter", "COTS", "A", "1.2083 (X42Cr13)"
    AddFile 8, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\80-2.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20_101_Spindle_Clamp_D80.SLDPRT", _
            "KEEP", "1060 Alloy", "RENAME", True, False, False, _
            "MODUMILL", "LM_Z_20_101_Spindle_Clamp_D80", "Spindle clamp D80 (legacy '80-2')", "FAB", "A", "1060 Alloy"
    AddFile 9, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P1.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P1.SLDPRT", _
            "KEEP", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P1", "CSMT-04BR1ANT3 component P1", "COTS", "A", "6061 Alloy"
    AddFile 10, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P2.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P2.SLDPRT", _
            "KEEP", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P2", "CSMT-04BR1ANT3 component P2", "COTS", "A", "6061 Alloy"
    AddFile 11, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P3.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P3.SLDPRT", _
            "KEEP", "Stainless Steel (ferritic)", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P3", "CSMT-04BR1ANT3 component P3", "COTS", "A", "Stainless Steel (ferritic)"
    AddFile 12, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P4.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P4.SLDPRT", _
            "KEEP", "AISI 1020", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P4", "CSMT-04BR1ANT3 component P4", "COTS", "A", "AISI 1020"
    AddFile 13, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P5.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P5.SLDPRT", _
            "KEEP", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P5", "CSMT-04BR1ANT3 component P5", "COTS", "A", "6061 Alloy"
    AddFile 14, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT-04BR1ANT3-P6.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_P6.SLDPRT", _
            "KEEP", "Stainless Steel (ferritic)", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_P6", "CSMT-04BR1ANT3 component P6", "COTS", "A", "Stainless Steel (ferritic)"
    AddFile 15, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\HGH20CAZ0C, BLOCK OIL FILLING PART.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_HGH20CA_Oil_Fitting.SLDPRT", _
            "SET", "Alloy Steel", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_HGH20CA_Oil_Fitting", "HGH20CA guide block - oil fitting", "COTS", "A", "Alloy Steel"
    AddFile 16, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\HGH20CAZ0C, BLOCK, MIDSECTION.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_HGH20CA_Block_Body.SLDPRT", _
            "KEEP", "Stainless Steel (ferritic)", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_HGH20CA_Block_Body", "HGH20CA guide block - body", "COTS", "A", "Stainless Steel (ferritic)"
    AddFile 17, "RENAME", "Part", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\HGH20CAZ0C, SCREW HEAD.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_HGH20CA_Endcap_Screw.SLDPRT", _
            "SET", "Alloy Steel", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_HGH20CA_Endcap_Screw", "HGH20CA guide block - end-cap screw", "COTS", "A", "Alloy Steel"
    AddFile 18, "SKIP_TOOLBOX", "Part", 4, "c:\solidworks data\browser\ansi metric\bolts and screws\self tapping screws\pan head tapping screw_am.sldprt", "c:\solidworks data\browser\ansi metric\bolts and screws\self tapping screws\pan head tapping screw_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "pan head tapping screw_am", "pan head tapping screw_am", "COTS", "A", ""
    AddFile 19, "SKIP_TOOLBOX", "Part", 4, "c:\solidworks data\browser\ansi metric\bolts and screws\socket head screws\socket head cap screw_am.sldprt", "c:\solidworks data\browser\ansi metric\bolts and screws\socket head screws\socket head cap screw_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "socket head cap screw_am", "socket head cap screw_am", "COTS", "A", ""
    AddFile 20, "RENAME", "Assembly", 4, "C:\Users\User\Desktop\Project\SOLIDWORKS\GDK80-24Z_1.5.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_SPN_GDK80_24Z_Spindle_Assy.SLDASM", _
            "NONE", "", "RENAME", False, False, False, _
            "MODUMILL", "COTS_SPN_GDK80_24Z_Spindle_Assy", "GDK80-24Z spindle assembly", "COTS", "A", ""
    AddFile 21, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\_bf10.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_BF10_Floating_Support.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", False, False, False, _
            "MODUMILL", "COTS_LIN_BF10_Floating_Support", "BF10 ball screw floating-side support", "COTS", "A", "Alloy Steel"
    AddFile 22, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\_bk10.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_BK10_Fixed_Support.SLDPRT", _
            "SET", "Alloy Steel", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_BK10_Fixed_Support", "BK10 ball screw fixed-side support", "COTS", "A", "Alloy Steel"
    AddFile 23, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\bast kochak.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_005_Small_Clamp_Bracket.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_005_Small_Clamp_Bracket", "Small clamp bracket (legacy 'bast kochak')", "FAB", "A", "6061 Alloy"
    AddFile 24, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Body, GT2 Timing Pulley, Configurable.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_PUL_GT2_Pulley_Body.SLDPRT", _
            "KEEP", "6061-T6 (SS)", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_PUL_GT2_Pulley_Body", "GT2 timing pulley - body", "COTS", "A", "6061-T6 (SS)"
    AddFile 25, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\bosh y.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_003_Spacer_Bushing.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_003_Spacer_Bushing", "Spacer bushing (legacy 'bosh y')", "FAB", "A", "6061 Alloy"
    AddFile 26, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Bottom Base, GT2 Timing Pulley, Configurable.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_PUL_GT2_Pulley_Bottom_Base.SLDPRT", _
            "KEEP", "6061-T6 (SS)", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_PUL_GT2_Pulley_Bottom_Base", "GT2 timing pulley - bottom base", "COTS", "A", "6061-T6 (SS)"
    AddFile 27, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\CHINA_Housing D22mm for ballscrew nut SFU1204.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_SFU1204_Nut_Housing_D22.SLDPRT", _
            "KEEP", "1023 Carbon Steel Sheet (SS)", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_SFU1204_Nut_Housing_D22", "Ball screw nut housing D22 for SFU1204", "COTS", "A", "1023 Carbon Steel Sheet (SS)"
    AddFile 28, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Copy of Master_Layout_s.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20_000_Master_Layout_Skeleton.SLDPRT", _
            "NONE", "", "NONE", False, False, True, _
            "MODUMILL", "LM_Z_20_000_Master_Layout_Skeleton", "Layout skeleton - Z spindle carriage (no mass, no BOM)", "FAB", "A", ""
    AddFile 29, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\CPM-_ _ _ _-231_ _-Q _.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_MOT_ClearPath_CPM_231.SLDPRT", _
            "SET", "Alloy Steel", "NONE", False, False, False, _
            "MODUMILL", "COTS_MOT_ClearPath_CPM_231", "Teknic ClearPath CPM 231 servo motor", "COTS", "A", "Alloy Steel"
    AddFile 30, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\fadkslkdfa.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20_002_Carriage_Spacer_Block.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_20_002_Carriage_Spacer_Block", "Carriage spacer block (legacy 'fadkslkdfa')", "FAB", "A", "Ductile Iron"
    AddFile 31, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\LM-Z-01-001_Z_Base_Mounting_Plate.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_001_Base_Mounting_Plate.SLDPRT", _
            "KEEP", "Alloy Steel", "RENAME", False, False, False, _
            "MODUMILL", "LM_Z_10_001_Base_Mounting_Plate", "Z base mounting plate", "FAB", "A", "Alloy Steel"
    AddFile 32, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\LM-Z-01-002_Z_Carriage_Plate.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20_001_Carriage_Plate.SLDPRT", _
            "SET", "6061 Alloy", "NONE", False, False, False, _
            "MODUMILL", "LM_Z_20_001_Carriage_Plate", "Z spindle carriage plate", "FAB", "A", "6061 Alloy"
    AddFile 33, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Master_Layout_s.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_000_Master_Layout_Skeleton.SLDPRT", _
            "NONE", "", "NONE", False, False, True, _
            "MODUMILL", "LM_Z_10_000_Master_Layout_Skeleton", "Layout skeleton - Z base saddle (no mass, no BOM)", "FAB", "A", ""
    AddFile 34, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\MDA2.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_008_Hollow_Frame_Bracket.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_008_Hollow_Frame_Bracket", "Hollow frame bracket (legacy MDA2)", "FAB", "A", "6061 Alloy"
    AddFile 35, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\MGN12 LINEAR GUIDE RAIL, CONFIGURABLE.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_MGN12_Rail.SLDPRT", _
            "KEEP", "Plain Carbon Steel", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_LIN_MGN12_Rail", "MGN12 linear guide rail (configurable length)", "COTS", "A", "Plain Carbon Steel"
    AddFile 36, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\MGN12H, LINEAR GUIDE BLOCK.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_MGN12H_Block.SLDPRT", _
            "KEEP", "Stainless Steel (ferritic)", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_MGN12H_Block", "MGN12H linear guide block", "COTS", "A", "Stainless Steel (ferritic)"
    AddFile 37, "SKIP_TOOLBOX", "Part", 3, "c:\solidworks data\browser\ansi metric\bolts and screws\machine screws\pan cross head_am.sldprt", "c:\solidworks data\browser\ansi metric\bolts and screws\machine screws\pan cross head_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "pan cross head_am", "pan cross head_am", "COTS", "A", ""
    AddFile 38, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Part1.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_002_Bracket_Plate_A.SLDPRT", _
            "KEEP", "1060 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_002_Bracket_Plate_A", "Bracket plate A (legacy Part1)", "FAB", "A", "1060 Alloy"
    AddFile 39, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Part4^MODUMILL.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_004_Bracket_Plate_B.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_004_Bracket_Plate_B", "Bracket plate B (legacy Part4, shared with gantry)", "FAB", "A", "6061 Alloy"
    AddFile 40, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Part5.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_009_Spacer_Block.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_009_Spacer_Block", "Spacer block (legacy Part5)", "FAB", "A", "6061 Alloy"
    AddFile 41, "SKIP_TOOLBOX", "Part", 3, "c:\solidworks data\browser\ansi metric\bolts and screws\set screws (socket)\socket set screw flat point_am.sldprt", "c:\solidworks data\browser\ansi metric\bolts and screws\set screws (socket)\socket set screw flat point_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "socket set screw flat point_am", "socket set screw flat point_am", "COTS", "A", ""
    AddFile 42, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Step motor holder NEMA 23.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_006_NEMA23_Motor_Holder.SLDPRT", _
            "SET", "6061 Alloy", "RENAME", False, False, False, _
            "MODUMILL", "LM_Z_10_006_NEMA23_Motor_Holder", "NEMA 23 stepper motor holder", "FAB", "A", "6061 Alloy"
    AddFile 43, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Top Flange, GT2 Timing Pulley, Configurable.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_PUL_GT2_Pulley_Top_Flange.SLDPRT", _
            "KEEP", "6061-T6 (SS)", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_PUL_GT2_Pulley_Top_Flange", "GT2 timing pulley - top flange", "COTS", "A", "6061-T6 (SS)"
    AddFile 44, "RENAME", "Part", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Z_Z_Base_C001.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10_007_Z_Base_Cast_Bracket.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", True, False, False, _
            "MODUMILL", "LM_Z_10_007_Z_Base_Cast_Bracket", "Z base cast bracket (legacy Z_Z_Base_C001)", "FAB", "A", "Ductile Iron"
    AddFile 45, "RENAME", "Assembly", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\Assem1.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20100_Spindle_Mount_Assy.SLDASM", _
            "NONE", "", "NONE", False, False, False, _
            "MODUMILL", "LM_Z_20100_Spindle_Mount_Assy", "Spindle mount assembly", "FAB", "A", ""
    AddFile 46, "RENAME", "Assembly", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\CSMT04BR1ANT3.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_BRK_CSMT04BR1ANT3_Assy.SLDASM", _
            "NONE", "", "NONE", True, False, False, _
            "MODUMILL", "COTS_BRK_CSMT04BR1ANT3_Assy", "CSMT-04BR1ANT3 assembly", "COTS", "A", ""
    AddFile 47, "RENAME", "Assembly", 3, "C:\Users\User\Desktop\Project\SOLIDWORKS\HGH20CAZ0C.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_HGH20CA_Block_Assy.SLDASM", _
            "NONE", "", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_HGH20CA_Block_Assy", "HGH20CA linear guide block assembly", "COTS", "A", ""
    AddFile 48, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\_m10sfunut.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_SFU1204_Lock_Nut_M10.SLDPRT", _
            "SET", "Alloy Steel", "NONE", True, False, False, _
            "MODUMILL", "COTS_LIN_SFU1204_Lock_Nut_M10", "M10 lock nut for SFU1204 ball screw", "COTS", "A", "Alloy Steel"
    AddFile 49, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\_sfu1204nut.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_SFU1204_Ballscrew_Nut.SLDPRT", _
            "KEEP", "1023 Carbon Steel Sheet (SS)", "NONE", False, False, False, _
            "MODUMILL", "COTS_LIN_SFU1204_Ballscrew_Nut", "Ball screw nut SFU1204", "COTS", "A", "1023 Carbon Steel Sheet (SS)"
    AddFile 50, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\_sfu1204x500.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_SFU1204_Ballscrew_L500.SLDPRT", _
            "SET", "Alloy Steel", "RENAME", False, False, False, _
            "MODUMILL", "COTS_LIN_SFU1204_Ballscrew_L500", "Ball screw SFU1204 L500", "COTS", "A", "Alloy Steel"
    AddFile 51, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\CopyofCopyofPart5.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_003_Pin_Small.SLDPRT", _
            "SET", "6061 Alloy", "NONE", True, False, False, _
            "MODUMILL", "LM_X_00_003_Pin_Small", "Small pin (legacy CopyofCopyofPart5)", "FAB", "A", "6061 Alloy"
    AddFile 52, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\CopyofPart3.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_004_Gantry_Plate.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", True, False, False, _
            "MODUMILL", "LM_X_00_004_Gantry_Plate", "Gantry plate (legacy CopyofPart3)", "FAB", "A", "Ductile Iron"
    AddFile 53, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\ewrewq.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_001_Spacer_Small.SLDPRT", _
            "SET", "Alloy Steel", "NONE", True, False, False, _
            "MODUMILL", "LM_X_00_001_Spacer_Small", "Small spacer (legacy 'ewrewq')", "FAB", "A", "Alloy Steel"
    AddFile 54, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\frontnew.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_BED_01_001_Front_Plate.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", False, False, False, _
            "MODUMILL", "LM_BED_01_001_Front_Plate", "Bed front plate", "FAB", "A", "Ductile Iron"
    AddFile 55, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\gantry_ cover.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_006_Gantry_Cover.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", False, False, False, _
            "MODUMILL", "LM_X_00_006_Gantry_Cover", "Gantry cover", "FAB", "A", "Ductile Iron"
    AddFile 56, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\gfgfxsgfg.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_002_Gantry_Hollow_Bracket.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", True, False, False, _
            "MODUMILL", "LM_X_00_002_Gantry_Hollow_Bracket", "Gantry hollow bracket (legacy 'gfgfxsgfg')", "FAB", "A", "Ductile Iron"
    AddFile 57, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\GT260Tid8mmbw7mm.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_PUL_GT2_60T_B8_W7_Pulley.SLDPRT", _
            "KEEP", "6061-T6 (SS)", "RENAME", False, False, False, _
            "MODUMILL", "COTS_PUL_GT2_60T_B8_W7_Pulley", "GT2 60T pulley, 8 mm bore, 7 mm belt width", "COTS", "A", "6061-T6 (SS)"
    AddFile 58, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\HGR20RLINEARGUIDE RAIL,CONFIGURABLE.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_LIN_HGR20_Rail.SLDPRT", _
            "KEEP", "Stainless Steel (ferritic)", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_LIN_HGR20_Rail", "HGR20 linear guide rail (configurable length)", "COTS", "A", "Stainless Steel (ferritic)"
    AddFile 59, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\Plate.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_BED_01_003_Weldment_Frame.SLDPRT", _
            "SET", "Plain Carbon Steel", "NONE", True, False, False, _
            "MODUMILL", "LM_BED_01_003_Weldment_Frame", "Bed weldment frame plate", "FAB", "A", "Plain Carbon Steel"
    AddFile 60, "SKIP_TOOLBOX", "Part", 2, "c:\solidworks data\browser\ansi metric\power transmission\gears\rack spur rectangular_am.sldprt", "c:\solidworks data\browser\ansi metric\power transmission\gears\rack spur rectangular_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "rack spur rectangular_am", "rack spur rectangular_am", "COTS", "A", ""
    AddFile 61, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\side.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00_005_Gantry_Side_Plate.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", False, False, False, _
            "MODUMILL", "LM_X_00_005_Gantry_Side_Plate", "Gantry side plate", "FAB", "A", "Ductile Iron"
    AddFile 62, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\side_plate_of_bed.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_BED_01_002_Side_Plate.SLDPRT", _
            "KEEP", "Ductile Iron", "NONE", False, False, False, _
            "MODUMILL", "LM_BED_01_002_Side_Plate", "Bed side plate", "FAB", "A", "Ductile Iron"
    AddFile 63, "SKIP_TOOLBOX", "Part", 2, "c:\solidworks data\browser\ansi metric\power transmission\gears\spur gear_am.sldprt", "c:\solidworks data\browser\ansi metric\power transmission\gears\spur gear_am.sldprt", _
            "SKIP_TOOLBOX", "", "SKIP_TOOLBOX", False, True, False, _
            "MODUMILL", "spur gear_am", "spur gear_am", "COTS", "A", ""
    AddFile 64, "RENAME", "Part", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\Surface.SLDPRT", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_BED_01_004_Table_Surface.SLDPRT", _
            "KEEP", "1023 Carbon Steel Sheet (SS)", "NONE", True, False, False, _
            "MODUMILL", "LM_BED_01_004_Table_Surface", "Bed table surface", "FAB", "A", "1023 Carbon Steel Sheet (SS)"
    AddFile 65, "RENAME", "Assembly", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\GT2TimingPulleyConfigurable.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\COTS_PUL_GT2_Pulley_Assy.SLDASM", _
            "NONE", "", "SKIP_DESIGN_TABLE_SUSPECT", False, False, False, _
            "MODUMILL", "COTS_PUL_GT2_Pulley_Assy", "GT2 timing pulley assembly", "COTS", "A", ""
    AddFile 66, "RENAME", "Assembly", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\LMZ10000_Z_Base_Saddle_SubAssy.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_10000_Z_Base_Saddle_SubAssy.SLDASM", _
            "NONE", "", "RENAME", False, False, False, _
            "MODUMILL", "LM_Z_10000_Z_Base_Saddle_SubAssy", "Z base saddle sub-assembly", "FAB", "A", ""
    AddFile 67, "RENAME", "Assembly", 2, "C:\Users\User\Desktop\Project\SOLIDWORKS\LMZ20000_Z_Spindle_Carriage_SubAssy.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_20000_Z_Spindle_Carriage_SubAssy.SLDASM", _
            "NONE", "", "RENAME", False, False, False, _
            "MODUMILL", "LM_Z_20000_Z_Spindle_Carriage_SubAssy", "Z spindle carriage sub-assembly", "FAB", "A", ""
    AddFile 68, "RENAME", "Assembly", 1, "C:\Users\User\Desktop\Project\SOLIDWORKS\Assem2.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_X_00000_Gantry_Bridge_Assy.SLDASM", _
            "NONE", "", "NONE", False, False, False, _
            "MODUMILL", "LM_X_00000_Gantry_Bridge_Assy", "Gantry bridge assembly", "FAB", "A", ""
    AddFile 69, "RENAME", "Assembly", 1, "C:\Users\User\Desktop\Project\SOLIDWORKS\Bed_System.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_BED_00000_Bed_System_Assy.SLDASM", _
            "NONE", "", "NONE", False, False, False, _
            "MODUMILL", "LM_BED_00000_Bed_System_Assy", "Machine bed system assembly", "FAB", "A", ""
    AddFile 70, "RENAME", "Assembly", 1, "C:\Users\User\Desktop\Project\SOLIDWORKS\LM-Z-00-000_Z_Axis_Complete_Assy.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\LM_Z_00000_Z_Axis_Complete_Assy.SLDASM", _
            "NONE", "", "RENAME", False, False, False, _
            "MODUMILL", "LM_Z_00000_Z_Axis_Complete_Assy", "Z axis complete assembly", "FAB", "A", ""
    AddFile 71, "KEEP", "Assembly", 0, "C:\Users\User\Desktop\Project\SOLIDWORKS\MODUMILL.SLDASM", "C:\Users\User\Desktop\Project\SOLIDWORKS\MODUMILL.SLDASM", _
            "NONE", "", "NONE", False, False, False, _
            "MODUMILL", "MODUMILL", "MODUMILL top-level machine assembly", "FAB", "A", ""

    ' --- configuration renames (fileId, oldName, newName) ---
    AddCfg 1, 1, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 2, 2, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 3, 3, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 4, 4, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 5, 5, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 6, 6, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 7, 7, "Domy" & ChrW(347) & "lna", "Default"
    AddCfg 8, 8, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 9, 20, ChrW(1055) & ChrW(1086) & " " & ChrW(1091) & ChrW(1084) & ChrW(1086) & ChrW(1083) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1102), "Default"
    AddCfg 10, 21, "with cosmetic bearing", "with_cosmetic_bearing"
    AddCfg 11, 21, "body only", "body_only"
    AddCfg 12, 24, "16T W10 B5", "16T_W10_B5"
    AddCfg 13, 24, "16T W3 B3", "16T_W3_B3"
    AddCfg 14, 24, "16T W6 B5", "16T_W6_B5"
    AddCfg 15, 24, "16T W6 B6", "16T_W6_B6"
    AddCfg 16, 24, "20T W10 B5", "20T_W10_B5"
    AddCfg 17, 24, "20T W10 B8", "20T_W10_B8"
    AddCfg 18, 24, "20T W6 B4", "20T_W6_B4"
    AddCfg 19, 24, "20T W6 B5", "20T_W6_B5"
    AddCfg 20, 24, "20T W6 B8", "20T_W6_B8"
    AddCfg 21, 26, "16T W10 B5", "16T_W10_B5"
    AddCfg 22, 26, "16T W3 B3", "16T_W3_B3"
    AddCfg 23, 26, "16T W6 B5", "16T_W6_B5"
    AddCfg 24, 26, "16T W6 B6", "16T_W6_B6"
    AddCfg 25, 26, "20T W10 B5", "20T_W10_B5"
    AddCfg 26, 26, "20T W10 B8", "20T_W10_B8"
    AddCfg 27, 26, "20T W6 B4", "20T_W6_B4"
    AddCfg 28, 26, "20T W6 B5", "20T_W6_B5"
    AddCfg 29, 26, "20T W6 B8", "20T_W6_B8"
    AddCfg 30, 31, "CFG-02_SFU1204_BK10_Standard", "CFG_02_SFU1204_BK10_Standard"
    AddCfg 31, 35, "100 mm", "100_mm"
    AddCfg 32, 35, "150 mm", "150_mm"
    AddCfg 33, 35, "200 mm", "200_mm"
    AddCfg 34, 35, "250 mm", "250_mm"
    AddCfg 35, 35, "300 mm", "300_mm"
    AddCfg 36, 35, "350 mm", "350_mm"
    AddCfg 37, 35, "400 mm", "400_mm"
    AddCfg 38, 35, "450 mm", "450_mm"
    AddCfg 39, 35, "500 mm", "500_mm"
    AddCfg 40, 35, "550 mm", "550_mm"
    AddCfg 41, 35, "600 mm", "600_mm"
    AddCfg 42, 35, "650 mm", "650_mm"
    AddCfg 43, 35, "700 mm", "700_mm"
    AddCfg 44, 35, "750 mm", "750_mm"
    AddCfg 45, 35, "800 mm", "800_mm"
    AddCfg 46, 35, "850 mm", "850_mm"
    AddCfg 47, 35, "900 mm", "900_mm"
    AddCfg 48, 35, "950 mm", "950_mm"
    AddCfg 49, 35, "1000 mm", "1000_mm"
    AddCfg 50, 35, "1050 mm", "1050_mm"
    AddCfg 51, 35, "1100 mm", "1100_mm"
    AddCfg 52, 35, "1150 mm", "1150_mm"
    AddCfg 53, 35, "1200 mm", "1200_mm"
    AddCfg 54, 35, "1250 mm", "1250_mm"
    AddCfg 55, 35, "1300 mm", "1300_mm"
    AddCfg 56, 35, "1350 mm", "1350_mm"
    AddCfg 57, 35, "1400 mm", "1400_mm"
    AddCfg 58, 35, "1450 mm", "1450_mm"
    AddCfg 59, 35, "1500 mm", "1500_mm"
    AddCfg 60, 35, "1550 mm", "1550_mm"
    AddCfg 61, 35, "1600 mm", "1600_mm"
    AddCfg 62, 35, "1650 mm", "1650_mm"
    AddCfg 63, 35, "1700 mm", "1700_mm"
    AddCfg 64, 35, "1750 mm", "1750_mm"
    AddCfg 65, 35, "1800 mm", "1800_mm"
    AddCfg 66, 35, "1850 mm", "1850_mm"
    AddCfg 67, 35, "1900 mm", "1900_mm"
    AddCfg 68, 35, "1950 mm", "1950_mm"
    AddCfg 69, 35, "2000 mm", "2000_mm"
    AddCfg 70, 42, "Domy" & ChrW(347) & "lna", "Default"
    AddCfg 71, 43, "16T W10 B5", "16T_W10_B5"
    AddCfg 72, 43, "16T W3 B3", "16T_W3_B3"
    AddCfg 73, 43, "16T W6 B5", "16T_W6_B5"
    AddCfg 74, 43, "16T W6 B6", "16T_W6_B6"
    AddCfg 75, 43, "20T W10 B5", "20T_W10_B5"
    AddCfg 76, 43, "20T W10 B8", "20T_W10_B8"
    AddCfg 77, 43, "20T W6 B4", "20T_W6_B4"
    AddCfg 78, 43, "20T W6 B5", "20T_W6_B5"
    AddCfg 79, 43, "20T W6 B8", "20T_W6_B8"
    AddCfg 80, 50, "cosmetic threaded", "cosmetic_threaded"
    AddCfg 81, 57, "V" & ChrW(253) & "choz" & ChrW(237), "Default"
    AddCfg 82, 58, "1000 mm", "1000_mm"
    AddCfg 83, 58, "1500 mm", "1500_mm"
    AddCfg 84, 58, "150 mm", "150_mm"
    AddCfg 85, 58, "200 mm", "200_mm"
    AddCfg 86, 58, "250 mm", "250_mm"
    AddCfg 87, 58, "300 mm", "300_mm"
    AddCfg 88, 58, "350 mm", "350_mm"
    AddCfg 89, 58, "400 mm", "400_mm"
    AddCfg 90, 58, "450 mm", "450_mm"
    AddCfg 91, 58, "500 mm", "500_mm"
    AddCfg 92, 58, "550 mm", "550_mm"
    AddCfg 93, 58, "600 mm", "600_mm"
    AddCfg 94, 58, "650 mm", "650_mm"
    AddCfg 95, 58, "700 mm", "700_mm"
    AddCfg 96, 58, "750 mm", "750_mm"
    AddCfg 97, 58, "800 mm", "800_mm"
    AddCfg 98, 58, "850 mm", "850_mm"
    AddCfg 99, 58, "900 mm", "900_mm"
    AddCfg 100, 58, "950 mm", "950_mm"
    AddCfg 101, 58, "1050 mm", "1050_mm"
    AddCfg 102, 58, "1100 mm", "1100_mm"
    AddCfg 103, 58, "1150 mm", "1150_mm"
    AddCfg 104, 58, "1200 mm", "1200_mm"
    AddCfg 105, 58, "1250 mm", "1250_mm"
    AddCfg 106, 58, "1300 mm", "1300_mm"
    AddCfg 107, 58, "1350 mm", "1350_mm"
    AddCfg 108, 58, "1400 mm", "1400_mm"
    AddCfg 109, 58, "1450 mm", "1450_mm"
    AddCfg 110, 58, "1550 mm", "1550_mm"
    AddCfg 111, 58, "1600 mm", "1600_mm"
    AddCfg 112, 58, "1650 mm", "1650_mm"
    AddCfg 113, 58, "1700 mm", "1700_mm"
    AddCfg 114, 58, "1750 mm", "1750_mm"
    AddCfg 115, 58, "1800 mm", "1800_mm"
    AddCfg 116, 58, "1850 mm", "1850_mm"
    AddCfg 117, 58, "1900 mm", "1900_mm"
    AddCfg 118, 58, "2000 mm", "2000_mm"
    AddCfg 119, 58, "1950 mm", "1950_mm"
    AddCfg 120, 58, "2050 mm", "2050_mm"
    AddCfg 121, 58, "2100 mm", "2100_mm"
    AddCfg 122, 58, "2150 mm", "2150_mm"
    AddCfg 123, 58, "2200 mm", "2200_mm"
    AddCfg 124, 58, "2250 mm", "2250_mm"
    AddCfg 125, 58, "2300 mm", "2300_mm"
    AddCfg 126, 58, "2350 mm", "2350_mm"
    AddCfg 127, 58, "2400 mm", "2400_mm"
    AddCfg 128, 58, "2450 mm", "2450_mm"
    AddCfg 129, 58, "2500 mm", "2500_mm"
    AddCfg 130, 58, "2550 mm", "2550_mm"
    AddCfg 131, 58, "2600 mm", "2600_mm"
    AddCfg 132, 58, "2650 mm", "2650_mm"
    AddCfg 133, 58, "2700 mm", "2700_mm"
    AddCfg 134, 58, "2750 mm", "2750_mm"
    AddCfg 135, 58, "2800 mm", "2800_mm"
    AddCfg 136, 58, "2850 mm", "2850_mm"
    AddCfg 137, 58, "2900 mm", "2900_mm"
    AddCfg 138, 58, "2950 mm", "2950_mm"
    AddCfg 139, 58, "3000 mm", "3000_mm"
    AddCfg 140, 58, "3050 mm", "3050_mm"
    AddCfg 141, 58, "3100 mm", "3100_mm"
    AddCfg 142, 58, "3150 mm", "3150_mm"
    AddCfg 143, 58, "3200 mm", "3200_mm"
    AddCfg 144, 58, "3250 mm", "3250_mm"
    AddCfg 145, 58, "3300 mm", "3300_mm"
    AddCfg 146, 58, "3350 mm", "3350_mm"
    AddCfg 147, 58, "3400 mm", "3400_mm"
    AddCfg 148, 58, "3450 mm", "3450_mm"
    AddCfg 149, 58, "3500 mm", "3500_mm"
    AddCfg 150, 58, "3550 mm", "3550_mm"
    AddCfg 151, 58, "3600 mm", "3600_mm"
    AddCfg 152, 58, "3650 mm", "3650_mm"
    AddCfg 153, 58, "3700 mm", "3700_mm"
    AddCfg 154, 58, "3750 mm", "3750_mm"
    AddCfg 155, 58, "3800 mm", "3800_mm"
    AddCfg 156, 58, "3850 mm", "3850_mm"
    AddCfg 157, 58, "3900 mm", "3900_mm"
    AddCfg 158, 58, "3950 mm", "3950_mm"
    AddCfg 159, 58, "4000 mm", "4000_mm"
    AddCfg 160, 65, "16T W10 B5", "16T_W10_B5"
    AddCfg 161, 65, "16T W3 B3", "16T_W3_B3"
    AddCfg 162, 65, "16T W6 B5", "16T_W6_B5"
    AddCfg 163, 65, "16T W6 B6", "16T_W6_B6"
    AddCfg 164, 65, "20T W10 B5", "20T_W10_B5"
    AddCfg 165, 65, "20T W10 B8", "20T_W10_B8"
    AddCfg 166, 65, "20T W6 B4", "20T_W6_B4"
    AddCfg 167, 65, "20T W6 B5", "20T_W6_B5"
    AddCfg 168, 65, "20T W6 B8", "20T_W6_B8"
    AddCfg 169, 66, "CFG-01_SFU1604_BK12_HeavyDuty", "CFG_01_SFU1604_BK12_HeavyDuty"
    AddCfg 170, 66, "CFG-02_SFU1204_BK10_Standard", "CFG_02_SFU1204_BK10_Standard"
    AddCfg 171, 67, "2.2k", "2_2k"
    AddCfg 172, 70, "Without advanced mate", "Without_advanced_mate"

    ' --- component configuration fixes ---
    AddFix 1, "GT2TimingPulleyConfigurable", "Bottom Base, GT2 Timing Pulley, Configurable", "20T W10 B8", "20T W6 B8", "20T_W10_B8", "20T_W6_B8"

    ' --- fasteners to suppress in the for_matlab chain (model stems) ---
    Fasteners.Add "pan cross head_am"
    Fasteners.Add "pan head tapping screw_am"
    Fasteners.Add "socket head cap screw_am"
    Fasteners.Add "socket set screw flat point_am"

    ' --- sub-assemblies that need the for_matlab config chain ---
    ChainAsmys.Add "assem2"
    ChainAsmys.Add "csmt04br1ant3"
    ChainAsmys.Add "gt2timingpulleyconfigurable"
    ChainAsmys.Add "lm-z-00-000_z_axis_complete_assy"
    ChainAsmys.Add "lmz10000_z_base_saddle_subassy"
    ChainAsmys.Add "lmz20000_z_spindle_carriage_subassy"

    ' --- phase 1 processing order (children before parents) ---
    AddOrder 1, 1
    AddOrder 2, 2
    AddOrder 3, 3
    AddOrder 4, 4
    AddOrder 5, 5
    AddOrder 6, 6
    AddOrder 7, 7
    AddOrder 8, 8
    AddOrder 9, 9
    AddOrder 10, 10
    AddOrder 11, 11
    AddOrder 12, 12
    AddOrder 13, 13
    AddOrder 14, 14
    AddOrder 15, 15
    AddOrder 16, 16
    AddOrder 17, 17
    AddOrder 18, 20
    AddOrder 19, 21
    AddOrder 20, 22
    AddOrder 21, 23
    AddOrder 22, 24
    AddOrder 23, 25
    AddOrder 24, 26
    AddOrder 25, 27
    AddOrder 26, 28
    AddOrder 27, 29
    AddOrder 28, 30
    AddOrder 29, 31
    AddOrder 30, 32
    AddOrder 31, 33
    AddOrder 32, 34
    AddOrder 33, 35
    AddOrder 34, 36
    AddOrder 35, 38
    AddOrder 36, 39
    AddOrder 37, 40
    AddOrder 38, 42
    AddOrder 39, 43
    AddOrder 40, 44
    AddOrder 41, 45
    AddOrder 42, 46
    AddOrder 43, 47
    AddOrder 44, 48
    AddOrder 45, 49
    AddOrder 46, 50
    AddOrder 47, 51
    AddOrder 48, 52
    AddOrder 49, 53
    AddOrder 50, 54
    AddOrder 51, 55
    AddOrder 52, 56
    AddOrder 53, 57
    AddOrder 54, 58
    AddOrder 55, 59
    AddOrder 56, 61
    AddOrder 57, 62
    AddOrder 58, 64
    AddOrder 59, 65
    AddOrder 60, 66
    AddOrder 61, 67
    AddOrder 62, 68
    AddOrder 63, 69
    AddOrder 64, 70
    AddOrder 65, 71

    ' --- phase 2 rename windows (assemblies, deepest first) ---
    AddWin 1, 20
    AddWin 2, 45
    AddWin 3, 46
    AddWin 4, 47
    AddWin 5, 65
    AddWin 6, 66
    AddWin 7, 67
    AddWin 8, 68
    AddWin 9, 69
    AddWin 10, 70
    AddWin 11, 71

    ' --- parent/child pairs for files that must be renamed ---
    AddPair 1, 20, 1
    AddPair 2, 20, 2
    AddPair 3, 20, 3
    AddPair 4, 20, 4
    AddPair 5, 20, 5
    AddPair 6, 20, 6
    AddPair 7, 20, 7
    AddPair 8, 45, 8
    AddPair 9, 45, 20
    AddPair 10, 46, 9
    AddPair 11, 46, 10
    AddPair 12, 46, 11
    AddPair 13, 46, 12
    AddPair 14, 46, 13
    AddPair 15, 46, 14
    AddPair 16, 47, 15
    AddPair 17, 47, 16
    AddPair 18, 47, 17
    AddPair 19, 65, 24
    AddPair 20, 65, 26
    AddPair 21, 65, 43
    AddPair 22, 66, 21
    AddPair 23, 66, 22
    AddPair 24, 66, 23
    AddPair 25, 66, 25
    AddPair 26, 66, 29
    AddPair 27, 66, 31
    AddPair 28, 66, 33
    AddPair 29, 66, 34
    AddPair 30, 66, 35
    AddPair 31, 66, 38
    AddPair 32, 66, 39
    AddPair 33, 66, 40
    AddPair 34, 66, 42
    AddPair 35, 66, 44
    AddPair 36, 66, 46
    AddPair 37, 66, 47
    AddPair 38, 67, 27
    AddPair 39, 67, 28
    AddPair 40, 67, 30
    AddPair 41, 67, 32
    AddPair 42, 67, 36
    AddPair 43, 67, 45
    AddPair 44, 68, 39
    AddPair 45, 68, 46
    AddPair 46, 68, 47
    AddPair 47, 68, 51
    AddPair 48, 68, 52
    AddPair 49, 68, 53
    AddPair 50, 68, 55
    AddPair 51, 68, 56
    AddPair 52, 68, 58
    AddPair 53, 68, 61
    AddPair 54, 69, 54
    AddPair 55, 69, 58
    AddPair 56, 69, 59
    AddPair 57, 69, 62
    AddPair 58, 69, 64
    AddPair 59, 70, 48
    AddPair 60, 70, 49
    AddPair 61, 70, 50
    AddPair 62, 70, 57
    AddPair 63, 70, 65
    AddPair 64, 70, 66
    AddPair 65, 70, 67
    AddPair 66, 71, 68
    AddPair 67, 71, 69
    AddPair 68, 71, 70
End Sub

Private Sub AddFile(ByVal id As Long, ByVal act As String, ByVal dt As String, ByVal depth As Long, _
                    ByVal oldP As String, ByVal newP As String, ByVal ma As String, ByVal mt As String, _
                    ByVal cp As String, ByVal nr As Boolean, ByVal tb As Boolean, ByVal ly As Boolean, _
                    ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, _
                    ByVal p4 As String, ByVal p5 As String, ByVal p6 As String)
    Files(id).Id = id
    Files(id).Action = act
    Files(id).DocType = dt
    Files(id).Depth = depth
    Files(id).OldPath = oldP
    Files(id).NewPath = newP
    Files(id).MatAction = ma
    Files(id).MatTarget = mt
    Files(id).CfgPolicy = cp
    Files(id).NeedsReview = nr
    Files(id).IsToolbox = tb
    Files(id).IsLayout = ly
    Files(id).PropNames(1) = "Project": Files(id).PropValues(1) = p1
    Files(id).PropNames(2) = "PartNo": Files(id).PropValues(2) = p2
    Files(id).PropNames(3) = "Description": Files(id).PropValues(3) = p3
    Files(id).PropNames(4) = "Category": Files(id).PropValues(4) = p4
    Files(id).PropNames(5) = "Revision": Files(id).PropValues(5) = p5
    Files(id).PropNames(6) = "Material": Files(id).PropValues(6) = p6
    Files(id).Renamed = False
    Files(id).UsedFallback = False
End Sub

Private Sub AddCfg(ByVal id As Long, ByVal fileId As Long, ByVal oldC As String, ByVal newC As String)
    Cfgs(id).FileId = fileId
    Cfgs(id).OldCfg = oldC
    Cfgs(id).NewCfg = newC
    Cfgs(id).Done = False
    Cfgs(id).Skipped = False
End Sub

Private Sub AddFix(ByVal id As Long, ByVal assm As String, ByVal comp As String, _
                   ByVal fromC As String, ByVal toC As String, ByVal fromN As String, ByVal toN As String)
    Fixes(id).AssmStem = assm
    Fixes(id).CompStem = comp
    Fixes(id).FromCfg = fromC
    Fixes(id).ToCfg = toC
    Fixes(id).FromCfgNew = fromN
    Fixes(id).ToCfgNew = toN
End Sub

Private Sub AddOrder(ByVal n As Long, ByVal fid As Long)
    OrderIds(n) = fid
End Sub

Private Sub AddWin(ByVal n As Long, ByVal fid As Long)
    Wins(n) = fid
End Sub

Private Sub AddPair(ByVal n As Long, ByVal parId As Long, ByVal childId As Long)
    PairPar(n) = parId
    PairChild(n) = childId
End Sub
