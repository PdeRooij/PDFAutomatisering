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
'' 13-6-18      Pieter de Rooij     Formed the stub
'' 3-8-18       Pieter de Rooij     Simplified to pure initialisation
''=======================================================
Sub InitializeAdobe()
    'Initialize app, AVDoc and FormApp
    Set gAcrobatApplication = CreateObject("AcroExch.App")
    Set gAcrobatAVDoc = CreateObject("AcroExch.AVDoc")
    Set gAFormApp = CreateObject("AFormAut.App")
    
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
'' 3-8-18       Pieter de Rooij     Opens PDF template after initialisation
''=======================================================
Sub OpenAdobe(ByVal strTemplLoc As String)
    ' Open PDF document
    If gAcrobatAVDoc.Open(strTemplLoc, "Toewijzingen") Then
        ' Succesfully opened
        Set gAcrobatPDDoc = gAcrobatAVDoc.GetPDDoc() ' Also store PDDoc
        
        ' With the PDDoc, it is now also possible to initialize the JScript bridge
        Set g_jso = gAcrobatPDDoc.GetJSObject
        g_intPageNum = 0                        ' Start spawning from page 0 onwards
        
        ' Show Acrobat window
        gAcrobatApplication.Show
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
'' 13-6-18      Pieter de Rooij     Formed the stub
'' 7-7-18       Pieter de Rooij     Proof of concept of using JScript to spawn
'' 3-8-18       Pieter de Rooij     Now using public variables
'' 4-8-18       Pieter de Rooij     Now spawns on consecutive pages instead of duplication on the first page
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
'' Called by:   ConvertAssignments
'' Call:        WriteAssignment(Name, Date, Type, [CounselPoint], [Assistant], [Concerns])
'' Arguments:   Name        - Name of the assignee
''              Date        - Date of the assignment
''              Type        - The type of assignment (Bible reading, initial call etc.)
''              CounselPoint - (Optional) Point of counsel for assignee
''              Assistant   - (Optional) Name of the assistant for the assignee
''              Concerns    - (Optional) Whether the current assignment concerns the assignee or assistant.
'' Comments:    Uses the public variable g_intCounter to determine the assignment to write to.
''              Assumes that specific assignment is already (made) available!
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-06-2018   Pieter de Rooij     Formed the stub
'' 04-08-2018   Pieter de Rooij     Spawn a new page if more assignments are required
''=======================================================
Function WriteAssignment(ByVal Name As String, ByVal AsDate As Date, ByVal AsType As String, Optional ByVal CounselPoint As Integer = 0, Optional ByVal Assistant As String = "", Optional ByVal Concerns As Integer = 0) As Boolean
    ' Next assignment is being written, increment counter
    g_intAsCount = g_intAsCount + 1
    ' Spawn a new page with assignments if needed
    If g_intPageNum * 4 < g_intAsCount Then
        SpawnAssignments
    End If
    
    ' Call function to fill fields
    
End Function

''=======================================================
'' Program:     WriteAdobeField
'' Desc:        Writes a given values to a specified field in a PDF document.
'' Called by:   WriteAssignment
'' Call:        WriteAdobeField(Field, Value, [FillColour])
'' Arguments:   Field       - Name of the field to write to
''              Value       - Value to write into that field
''              FillColour  - (Optional) What colour to fill specified field with
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Sub WriteAdobeField()
    Dim AcrobatApplication As Acrobat.CAcroApp
    Dim AcrobatDocument As Acrobat.CAcroAVDoc
    Dim fcount As Long
    Dim sFieldName As String
    
    Set AcrobatApplication = CreateObject("AcroExch.App")
    Set AcrobatDocument = CreateObject("AcroExch.AVDoc")
    
    If AcrobatDocument.Open("D:\Zaal\LTV\Toewijzingen formulier.pdf", "") Then
        AcrobatApplication.Show
        Set AcroForm = CreateObject("AFormAut.App")
        Set Fields = AcroForm.Fields
        fcount = Fields.Count
        
        Fields("Name0").Value = "Test"
        Fields("Name1").Value = "Test2"
        Fields("Name2").Value = "Test3"
        Fields("Name3").Value = "Test4"
    Else
        MsgBox "Failed to write field!"
    End If
    
    ' Neatly exit
    AcrobatApplication.Exit
    Set AcrobatApplication = Nothing
    Set AcrobatDocument = Nothing
    Set Field = Nothing
    Set Fields = Nothing
    Set AcroForm = Nothing
    
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
'' 4-8-18       Pieter de Rooij     Formed the stub
''=======================================================
Function SaveAdobe(ByVal strFLoc As String)
    
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
'' 13-6-18      Pieter de Rooij     Formed the stub
'' 3-8-18       Pieter de Rooij     Simplified for pure closing
''=======================================================
Sub CloseAdobe()
    ' Neatly exit Adobe modules
    gAcrobatApplication.Exit
    Set gAcrobatApplication = Nothing
    Set gAcrobatAVDoc = Nothing
    Set gAcrobatPDDoc = Nothing
    Set g_jso = Nothing
'    Set Field = Nothing
'    Set Fields = Nothing
    Set gAFormApp = Nothing
    
End Sub
