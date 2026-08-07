'==============================================================================
' Remove All Appearances
'
' Strips every appearance from the active display state.
'
' In a part, that means colours applied to faces, features, bodies and the part
' itself - including any applied by hand. Nothing is preserved.
'
' In an assembly, only the component-level appearances held by the assembly are
' removed. The part files the assembly references are never modified, so their
' own colours survive. That makes this identical to Remove Body and Component
' Appearances when run on an assembly, which is deliberate: reaching any further
' would mean editing files the user never opened.
'
' Only the active display state is affected. Other display states keep their
' appearances.
'
' This cannot be undone, so it asks before running.
'
' To use, open a part or assembly document and run the macro.
'
'   Version   0.1.0
'   Date      2026-08-07
'   Author    James Debono
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' There are no user settings.
'
' The companion macro, Remove Body and Component Appearances, shares this
' structure. The only real difference is what counts as ours to clear: there,
' bodies and the faces carrying a body's colour; here, everything a part holds.
' Fixes to one usually belong in the other.
'
' Removal works through the render material API, not RemoveMaterialProperty,
' which is scoped by *configuration* - and since several display states live
' under one configuration, that would clear every display state at once.
' GetRenderMaterials2 and AddDisplayStateSpecificRenderMaterial take a
' swDisplayStateOpts_e instead and can be pointed at the active display state.
'
' An appearance is a render material plus the list of entities it is attached to.
' There is no call to detach one entity, so the sequence is: read the entity
' list, drop all of them, add back any being kept, and write the result to this
' display state. An appearance left with no entities is gone. In a part nothing
' is ever kept, so nothing is ever written back.
'
' EditRebuild3 at the end is not optional. Without it the viewport keeps drawing
' the old colours until the display state is switched away and back.
'
' Nothing here may touch a referenced part file. In an assembly that means
' component entities only. Do not extend this to faces or bodies in an assembly,
' however much the macro's name suggests it should - those belong to the part
' documents, and editing them would dirty files the user never opened, including
' ones shared with other assemblies.
Const SHOW_DIAGNOSTICS As Boolean = True

Const MACRO_VERSION As String = "0.1.0"

Dim swApp As SldWorks.SldWorks

'--- Entry point --------------------------------------------------------------

Sub main()
    On Error GoTo mainError

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly document.", vbCritical
        Exit Sub
    End If

    Dim docType As Long
    docType = swModel.GetType()

    Dim isAssembly As Boolean
    If docType = swDocumentTypes_e.swDocPART Then
        isAssembly = False
    ElseIf docType = swDocumentTypes_e.swDocASSEMBLY Then
        isAssembly = True
    Else
        MsgBox "This macro only works on part or assembly documents.", vbCritical
        Exit Sub
    End If

    If Not Confirmed(isAssembly) Then Exit Sub

    StripAppearances swModel, isAssembly
    Exit Sub

mainError:
    MsgBox "An error occurred in main(): " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical
End Sub

' Asks before running. Unlike applying colours, this destroys work that cannot be
' recreated - a hand-applied face colour is gone for good - and the macro leaves
' no undo record behind it.
Function Confirmed(ByVal isAssembly As Boolean) As Boolean

    Dim prompt As String

    If isAssembly Then
        prompt = "Remove all component appearances from the active display state?" & _
                 vbCrLf & vbCrLf & _
                 "Every colour the assembly applies to its components will be" & vbCrLf & _
                 "cleared. The part files themselves are not modified, so any" & vbCrLf & _
                 "colours inside them will show through afterwards." & vbCrLf & vbCrLf & _
                 "Other display states are not affected. This cannot be undone."
    Else
        prompt = "Remove all appearances from the active display state?" & _
                 vbCrLf & vbCrLf & _
                 "Every colour on faces, features, bodies and the part itself" & vbCrLf & _
                 "will be cleared, including any applied by hand." & vbCrLf & vbCrLf & _
                 "Other display states are not affected. This cannot be undone."
    End If

    Confirmed = (MsgBox(prompt, vbYesNo + vbExclamation, "Remove All Appearances") = vbYes)
End Function

'--- Core ---------------------------------------------------------------------

Sub StripAppearances(swModel As SldWorks.ModelDoc2, ByVal isAssembly As Boolean)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    Dim tStart As Single, tScanned As Single, tStripped As Single, tRefreshed As Single
    tStart = Timer

    currentStep = "Reading appearances for the active display state"
    Dim vAppearances As Variant
    vAppearances = swModel.Extension.GetRenderMaterials2( _
        swDisplayStateOpts_e.swThisDisplayState, Empty)

    tScanned = Timer

    If IsEmpty(vAppearances) Then
        MsgBox "There are no appearances in the active display state.", vbInformation
        Exit Sub
    End If

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    Dim examined As Long, strippedTotal As Long, keptTotal As Long
    Dim emptied As Long, partial As Long, untouched As Long
    Dim reattachFails As Long, writeBackFails As Long
    examined = 0: strippedTotal = 0: keptTotal = 0
    emptied = 0: partial = 0: untouched = 0
    reattachFails = 0: writeBackFails = 0

    Dim i As Long, j As Long
    Dim swAppearance As SldWorks.RenderMaterial
    Dim vEnts As Variant

    If SHOW_DIAGNOSTICS Then
        Debug.Print "=== Remove All Appearances " & MACRO_VERSION & " ==="
        Debug.Print "appearances in this display state: " & (UBound(vAppearances) + 1)
    End If

    currentStep = "Detaching entities"
    For i = 0 To UBound(vAppearances)
        Set swAppearance = vAppearances(i)
        examined = examined + 1

        vEnts = swAppearance.GetEntities

        If SHOW_DIAGNOSTICS Then
            Debug.Print "appearance " & i & ": " & DescribeEntities(vEnts)
        End If

        If Not IsEmpty(vEnts) Then

            Dim keep As Collection
            Set keep = New Collection
            Dim strippedHere As Long
            strippedHere = 0

            For j = 0 To UBound(vEnts)
                If ShouldStrip(vEnts(j), isAssembly) Then
                    strippedHere = strippedHere + 1
                Else
                    keep.Add vEnts(j)
                End If
            Next j

            If strippedHere > 0 Then
                strippedTotal = strippedTotal + strippedHere
                keptTotal = keptTotal + keep.Count

                If keep.Count = 0 Then
                    emptied = emptied + 1
                Else
                    partial = partial + 1
                End If

                If SHOW_DIAGNOSTICS Then
                    Debug.Print "   clearing " & strippedHere & ", keeping " & keep.Count
                End If

                swAppearance.RemoveAllEntities

                Dim k As Long
                For k = 1 To keep.Count
                    If swAppearance.AddEntity(keep(k)) = False Then
                        reattachFails = reattachFails + 1
                        If SHOW_DIAGNOSTICS Then
                            Debug.Print "   FAILED to re-attach entity " & k
                        End If
                    End If
                Next k

                If keep.Count > 0 Then
                    Dim matId1 As Long, matId2 As Long
                    swAppearance.GetMaterialIds matId1, matId2
                    If swModel.Extension.AddDisplayStateSpecificRenderMaterial( _
                        swAppearance, swDisplayStateOpts_e.swThisDisplayState, _
                        Empty, matId1, matId2) = False Then
                        writeBackFails = writeBackFails + 1
                        If SHOW_DIAGNOSTICS Then
                            Debug.Print "   FAILED to write back"
                        End If
                    End If
                End If
            Else
                untouched = untouched + 1
            End If
        Else
            untouched = untouched + 1
        End If
    Next i

    tStripped = Timer

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True

    currentStep = "Rebuilding"
    swModel.EditRebuild3
    tRefreshed = Timer

    Dim msg As String
    If isAssembly Then
        msg = "Removed all component appearances from the active display state" & vbCrLf & _
              "Components Cleared: " & strippedTotal & vbCrLf & vbCrLf & _
              "The referenced part files were not modified, so colours" & vbCrLf & _
              "held inside them will still show."
    Else
        msg = "Removed all appearances from the active display state" & vbCrLf & _
              "Entities Cleared: " & strippedTotal
    End If

    msg = msg & vbCrLf & vbCrLf & "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Appearances examined: " & examined & vbCrLf & _
              "  emptied completely: " & emptied & vbCrLf & _
              "  partly kept: " & partial & vbCrLf & _
              "  left alone: " & untouched & vbCrLf & _
              "Entities kept: " & keptTotal & vbCrLf & _
              "Re-attach failures: " & reattachFails & vbCrLf & _
              "Write-back failures: " & writeBackFails & vbCrLf & vbCrLf & _
              "Scan:    " & Format(tScanned - tStart, "0.00") & " s" & vbCrLf & _
              "Strip:   " & Format(tStripped - tScanned, "0.00") & " s" & vbCrLf & _
              "Rebuild: " & Format(tRefreshed - tStripped, "0.00") & " s" & vbCrLf & _
              "Total:   " & Format(tRefreshed - tStart, "0.00") & " s"

        Debug.Print "emptied=" & emptied & "  partial=" & partial & _
                    "  untouched=" & untouched & _
                    "  kept=" & keptTotal & _
                    "  reattachFails=" & reattachFails & _
                    "  writeBackFails=" & writeBackFails
        Debug.Print "=== end ==="
    End If

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in StripAppearances at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

' Which entities this macro clears.
'
' In a part, everything: faces, features, bodies and the part itself. In an
' assembly, components only - those are what the assembly owns. Anything else an
' assembly's appearances point at belongs to a referenced part file and is left
' strictly alone, which is why this behaves identically to the body and component
' macro on an assembly.
Function ShouldStrip(ByVal ent As Object, ByVal isAssembly As Boolean) As Boolean
    On Error GoTo NotOurs

    If isAssembly Then
        ShouldStrip = (TypeOf ent Is SldWorks.Component2)
    Else
        ShouldStrip = True
    End If
    Exit Function

NotOurs:
    ShouldStrip = False
End Function

' A readable summary of what an appearance is attached to, for diagnostics.
Function DescribeEntities(ByVal vEnts As Variant) As String
    On Error GoTo Failed

    If IsEmpty(vEnts) Then
        DescribeEntities = "no entities"
        Exit Function
    End If

    Dim bodies As Long, faces As Long, comps As Long
    Dim feats As Long, others As Long
    Dim i As Long

    For i = 0 To UBound(vEnts)
        If TypeOf vEnts(i) Is SldWorks.Body2 Then
            bodies = bodies + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Face2 Then
            faces = faces + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Component2 Then
            comps = comps + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Feature Then
            feats = feats + 1
        Else
            others = others + 1
        End If
    Next i

    DescribeEntities = (UBound(vEnts) + 1) & " entities [" & _
                       "bodies=" & bodies & _
                       " faces=" & faces & _
                       " components=" & comps & _
                       " features=" & feats & _
                       " other=" & others & "]"
    Exit Function

Failed:
    DescribeEntities = "could not be read (" & Err.Description & ")"
End Function
