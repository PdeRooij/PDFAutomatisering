Attribute VB_Name = "PDF_Functions"
Option Explicit

''=======================================================
'' Program:     InitializeAdobe
'' Desc:        Initializes Adobe modules and creates public variables for reference.
'' Called by:   ConvertAssignments
'' Call:        InitializeAdobe
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 03-08-2018   Pieter de Rooij     Simplified to pure initialisation
'' 09-08-2018   Pieter de Rooij     Removed form references
''=======================================================
Sub InitializeAdobe()
    ' Initialize app, AVDoc and FormApp
    Set gAcrobatApplication = CreateObject("AcroExch.App")
    Set gAcrobatAVDoc = CreateObject("AcroExch.AVDoc")
    
    ' Initialize global variables
    g_intPageNum = 0                ' Start spawning from page 0 onwards
    g_intAsCount = 0                ' Start with assignment 0
    
End Sub

''=======================================================
'' Program:     OpenAdobe
'' Desc:        Opens the PDF template in Adobe specified by the user.
'' Called by:   ConvertAssignments
'' Call:        OpenAdobe strTemplLoc:=TemplateLocation
'' Arguments:   TemplateLocation    - String of the template's path
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 03-08-2018   Pieter de Rooij     Opens PDF template after initialisation
'' 09-08-2018   Pieter de Rooij     Removed form references
''=======================================================
Sub OpenAdobe(ByVal strTemplLoc As String)
    ' Open PDF document
    If gAcrobatAVDoc.Open(strTemplLoc, "Toewijzingen") Then
        ' Succesfully opened
        Set gAcrobatPDDoc = gAcrobatAVDoc.GetPDDoc() ' Also store PDDoc
        
        ' With the PDDoc, it is now also possible to initialize the JScript bridge
        Set g_jso = gAcrobatPDDoc.GetJSObject
        
        ' Show Acrobat window
'        gAcrobatApplication.Show
    End If
    
End Sub

''=======================================================
'' Program:     SpawnAssignments
'' Desc:        Spawns an extra page of assignments from the template.
'' Called by:   ConvertAssignments, WriteAssignment
'' Call:        SpawnAssignments
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 07-07-2018   Pieter de Rooij     Proof of concept of using JScript to spawn
'' 03-08-2018   Pieter de Rooij     Now using public variables
'' 04-08-2018   Pieter de Rooij     Now spawns on consecutive pages instead of duplication on the first page
''=======================================================
Sub SpawnAssignments()
    ' Find template and spawn it
    Dim Template As Object
    Dim spawn As Object
    Set Template = g_jso.GetTemplate("Toewijzingen")
    Set spawn = Template.spawn(g_intPageNum, True, False)
    g_intPageNum = g_intPageNum + 1     ' Increment page number
    
End Sub

''=======================================================
'' Program:     WriteAssignment
'' Desc:        Fills out one assignment with provided information.
'' Called by:   ConvertRow
'' Call:        WriteAssignment(Name, Date, Type, [CounselPoint], [Assistant], [Concerns])
'' Arguments:   Name        - Name of the assignee
''              Date        - Date of the assignment
''              Type        - The type of assignment (Bible reading, initial call etc.)
''              CounselPoint - (Optional) Point of counsel for assignee
''              Assistant   - (Optional) Name of the assistant for the assignee
''              Concerns    - (Optional) Whether the current assignment concerns the assignee (1) or assistant (2).
'' Comments:    Uses the public variable g_intCounter to determine the assignment to write to.
''              Assumes that specific assignment is already (made) available!
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 04-08-2018   Pieter de Rooij     Spawn a new page if more assignments are required
'' 08-08-2018   Pieter de Rooij     Fills an entire assignment based on input
''=======================================================
Sub WriteAssignment(ByVal strName As String, ByVal dAsDate As Date, ByVal strAsType As String, Optional ByVal intCounselPoint As Integer = 0, Optional ByVal strAssistant As String = "", Optional ByVal intConcerns As Integer = 0)
    ' Next assignment is being written, increment counter
    g_intAsCount = g_intAsCount + 1
    ' Spawn a new page with assignments if needed
    If g_intPageNum * 4 < g_intAsCount Then
        SpawnAssignments
    End If
    
    ' Construct field name prefix and suffix
    Dim strPreFName As String
    Dim intSufFName As Integer
    ' All fields spawned from the "Toewijzingen" template contain the prefix P(num).Toewijzingen.
    strPreFName = "P" & g_intPageNum - 1 & ".Toewijzingen."
    ' Four assignments on one page are distinguished by a suffix 0 - 3
    intSufFName = (g_intAsCount - 1) Mod 4
    
    '' Fill all fields
    ' Always fill date field
    WriteAdobeField strPreFName & "Date" & intSufFName, dAsDate
    
    ' Next determine whether an assistant is involved
    If strAssistant = "" Then
        ' No assistant involved, so just fill name and counsel point
        WriteAdobeField strPreFName & "Name" & intSufFName, strName
        WriteAdobeField strPreFName & "CounselPoint" & intSufFName, intCounselPoint
    ElseIf intConcerns = 1 Then
        ' Assignment for the assignee, highlight name, fill assistant and counsel point
        WriteAdobeField strPreFName & "Name" & intSufFName, strName, True
        WriteAdobeField strPreFName & "Assistant" & intSufFName, strAssistant
        WriteAdobeField strPreFName & "CounselPoint" & intSufFName, intCounselPoint
    Else    ' intConcerns = 2
        ' Assignment for the assistant, fill name and highlight assistant (no counsel point)
        WriteAdobeField strPreFName & "Name" & intSufFName, strName
        WriteAdobeField strPreFName & "Assistant" & intSufFName, strAssistant, True
    End If
    
    ' Lastly, put a tick before the corresponding assignment type
    TickType strPreFName, intSufFName, strAsType
    
End Sub

''=======================================================
'' Program:     WriteAdobeField
'' Desc:        Writes a given value to a specified field in a PDF document.
'' Called by:   WriteAssignment
'' Call:        WriteAdobeField(Field, Value, [Fill])
'' Arguments:   Field   - Name of the field to write to
''              Value   - Value to write into that field
''              Fill    - (Optional) Whether or not the field needs to be filled with a colour
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 08-08-2018   Pieter de Rooij     Dynamically writes adobe fields based on input
'' 09-08-2018   Pieter de Rooij     Switched to using JScript bridge and colour is now defined here
''=======================================================
Sub WriteAdobeField(ByVal strField As String, ByVal strVal As String, Optional ByVal bFill As Boolean = False)
    ' Write specified value to specified field
    g_jso.getField(strField).Value = strVal
    
    If bFill Then
        ' Also fill if colour is given
        g_jso.getField(strField).fillColor = g_jso.Color.yellow     ' Fill with yellow
    End If
    
End Sub

''=======================================================
'' Program:     TickType
'' Desc:        Puts a tick in the assignment type based on a provided (Dutch) string.
'' Called by:   WriteAssignment
'' Call:        TickType(DutchType)
'' Arguments:   DutchType   - A string of the type of assignment in Dutch
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 08-08-2018   Pieter de Rooij     Initial version
'' 09-08-2018   Pieter de Rooij     Switched to using the JScript bridge
''=======================================================
Sub TickType(ByVal strPre As String, ByVal strSuf As String, ByVal strDutchType As String)
    ' Translate Dutch type into English counterpart
    Dim strEngType
    Select Case LCase(strDutchType)
        Case "bijbellezen"
            strEngType = "bibleReading"
        Case "eerste gesprek"
            strEngType = "initialCall"
        Case "eerste nabezoek"
            strEngType = "firstRV"
        Case "tweede nabezoek"
            strEngType = "secondRV"
        Case "derde nabezoek"
            strEngType = "thirdRV"
        Case "bijbelstudie"
            strEngType = "bibleStudy"
        Case "lezing"
            strEngType = "talk"
        Case "anders"
            strEngType = "other"
        Case Else
            MsgBox "Onbekend type toewijzing (" & strDutchType & ") ingevoerd!"
            Exit Sub
    End Select
    
    ' Tick field
    g_jso.getField(strPre & strEngType & strSuf).checkThisBox 0, True
    
End Sub

''=======================================================
'' Program:     SaveAdobe
'' Desc:        Saves the current document to a specified location.
'' Called by:   ConvertAssignments
'' Call:        SaveAdobe(FileLocation)
'' Arguments:   FileLocation    - String where the PDF should be saved
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 04-08-2018   Pieter de Rooij     Formed the stub
''=======================================================
Function SaveAdobe(ByVal strFLoc As String) As Boolean
    ' Try to save to specified file
    If gAcrobatPDDoc.Save(PDSaveFull, strFLoc) = False Then
        SaveAdobe = False
    Else
        SaveAdobe = True
    End If
    
End Function

''=======================================================
'' Program:     CloseAdobe
'' Desc:        Closes Adobe modules after use.
'' Called by:   ConvertAssignments
'' Call:        CloseAdobe
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 03-08-2018   Pieter de Rooij     Simplified for pure closing
'' 09-08-2018   Pieter de Rooij     Removed form references and made closing even cleaner
''=======================================================
Sub CloseAdobe()
    ' Delete references to Adobe modules
    Set g_jso = Nothing
    Set gAcrobatPDDoc = Nothing
    Set gAcrobatAVDoc = Nothing
    
    ' Close application
    gAcrobatApplication.CloseAllDocs
    gAcrobatApplication.Hide
    gAcrobatApplication.Exit
    Set gAcrobatApplication = Nothing
    
End Sub
