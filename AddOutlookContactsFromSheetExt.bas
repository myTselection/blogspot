Attribute VB_Name = "AddOutlookContactsFromSheet"
'Developed by myT, http://myTselection.blogspot.com
Option Explicit


Public Sub RefreshQueries()

  Dim wks As Worksheet
  Dim qt As QueryTable
  Dim lo As ListObject

  For Each wks In Worksheets
    For Each qt In wks.QueryTables
        qt.Refresh BackgroundQuery:=False
    Next qt

    For Each lo In wks.ListObjects
        lo.QueryTable.Refresh BackgroundQuery:=False
    Next lo

  Next wks

  Set qt = Nothing
  Set wks = Nothing
End Sub
'http://www.globaliconnect.com/excel/index.php?option=com_content&view=article&id=167:import-contacts-from-excel-to-outlook-automate-in-vba&catid=79&Itemid=475



Sub ExcelWorksheetDataAddToOutlookContacts()
    'Automating Outlook from Excel: This example uses the Items.Add Method to export data from an Excel Worksheet to the default Contacts folder.
    'Automate Outlook from Excel, using Late Binding. You need not add a reference to the Outlook library in Excel (your host application), in this case you will not be able to use the Outlook's predefined constants and will need to replace them by their numerical values in your code.
    
    
    
    Dim oApplOutlook As Object
    Dim oNsOutlook As Object
    Dim oCFolder As Object
    Dim oDelFolder As Object
    Dim oCItem As Object
    'Dim olItems As Outlook.Items
    Dim olItems As Object
    'Dim olContactItem As contactItem
    Dim olContactItem As Object
    Dim oDelItems As Object
    Dim lLastRow As Long, i As Long, n As Long, c As Long
    Dim firstRowToProcess As Integer, emailColumn As Integer, pictureColumn As Integer, processedColumn As Integer, itemsFound As Integer
    Dim sheetName As String, fullFilePath As String
    Dim updateExistingContacts As Boolean
    
    
    RefreshQueries
    
    'Config:
    sheetName = "Contacts"
    firstRowToProcess = 2
    updateExistingContacts = False
    
    updateExistingContacts = MsgBox("Update and overwrite fields of existing contacts? Pictures will always be updated. (Existing notes will be appended, not overwritten)", vbYesNo, "Overwrite")
    
    
   Application.ScreenUpdating = False
    ' turns off screen updating
    Application.DisplayStatusBar = True
    ' makes sure that the statusbar is visible
    Application.StatusBar = "Preparing export contacts to Outlook"
    
    'determine last data row in the worksheet:
    lLastRow = Sheets(sheetName).Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a new instance of the Outlook application, if an existing Outlook object is not available.
    'Set the Application object as follows:
    On Error Resume Next
    Set oApplOutlook = GetObject(, "Outlook.Application")
    'if an instance of an existing Outlook object is not available, an error will occur (Err.Number = 0 means no error):
    If Err.Number <> 0 Then
        Set oApplOutlook = CreateObject("Outlook.Application")
    End If
    'disable error handling:
    On Error GoTo 0
    
    'use the GetNameSpace method to instantiate (ie. create an instance) a NameSpace object variable, to access existing Outlook items. Set the NameSpace object as follows:
    Set oNsOutlook = oApplOutlook.GetNamespace("MAPI")
    
    '----------------------------
    
    'Empty the Deleted Items folder in Outlook so that when you quit the Outlook application you bypass the prompt: Are you sure you want to permanently delete all the items and subfolders in the "Deleted Items" folder?
    
    'set the default Deleted Items folder:
'    'The numerical value of olFolderDeletedItems is 3. The following code has replaced the Outlook's built-in constant olFolderDeletedItems by its numerical value 3.
'    Set oDelFolder = oNsOutlook.GetDefaultFolder(3)
'    'set the items collection:
'    Set oDelItems = oDelFolder.Items
'
'    'determine number of items in the collection:
'    c = oDelItems.Count
'    'start deleting from the last item:
'    For n = c To 1 Step -1
'        oDelItems(n).Delete
'    Next n
'
    '----------------------------
    
    'set reference to the default Contact Items folder:
    'The numerical value of olFolderContacts is 10. The following code has replaced the Outlook's built-in constant olFolderContacts by its numerical value 10.
    Set oCFolder = oNsOutlook.GetDefaultFolder(10)
'    Set olItems = oCFolder.Items.Restrict("[MessageClass]='IPM.Contact'")
    
    'Find contact to update, if not found add a new contact item
    
    
    'post each row's data on a separate contact item form:
    For i = firstRowToProcess To lLastRow
    'restrict info: https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/items-restrict-method-outlook
    'folder items info: https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/folder-items-property-outlook
        Set olItems = oCFolder.Items.Restrict("[MessageClass]='IPM.Contact' And [Email1Address] = '" & Sheets(sheetName).Cells(i, Sheets("References").Range("email").Value) & "'")
        
        itemsFound = 0
        For Each olContactItem In olItems
            
            Application.StatusBar = "Updating Outlook contact " + olContactItem.FullName & ", " & Sheets(sheetName).Cells(i, Sheets("References").Range("pnr").Value)
            'matching contact found in Outlook and Excel
			'web pictures should be stored in a local folder
            fullFilePath = Replace(Sheets(sheetName).Cells(i, Sheets("References").Range("pictureurl").Value), "http://someimageurl", "C:\Users\username\Documents\Pictures\profilephotos\")
            If FileExists(fullFilePath) Then
                olContactItem.AddPicture (fullFilePath)
            Else
                'MsgBox ("Missing picture: " + fullFilePath + ", for member: " + olContactItem.firstName + " " + olContactItem.lastname)
                Debug.Print ("Missing picture: " & fullFilePath & ", for member: " & olContactItem.firstName & " " & olContactItem.lastName)
            End If
            If updateExistingContacts = True Then
                With olContactItem
                    'update existing contact fields https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_properties.aspx
                    .firstName = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value)
                    .lastName = Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    If (Sheets(sheetName).Cells(i, Sheets("References").Range("birthdate").Value) <> "") Then
                        .Birthday = DateValue(Sheets(sheetName).Cells(i, Sheets("References").Range("birthdate").Value))
                    End If
                    .BusinessAddressStreet = "BusinessStreet 100"
                    .BusinessAddressCity = "BusinessCity"
                    .BusinessAddressCountry = "BusinessCountry"
                    .BusinessAddressPostalCode = "BusinessPostal"
                    .BusinessHomePage = "http://www.company.com"
                    'do not remove or duplicate existing categories, set some desired category when not yet set
                    If (InStr(.Categories, "Business") = 0) Then
                        .Categories = .Categories & ",Business"
                    End If
                    .CompanyName = "Company"
                    .Email1DisplayName = Sheets(sheetName).Cells(i, Sheets("References").Range("email").Value)
                    .FullName = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    .FileAs = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    '.gender = Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value)
'                    If (Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value) = "M") Then
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olMale
'                    ElseIf (Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value) = "F") Then
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olFemale
'                    Else
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olUnspecified
'                    End If
                    .HomeAddressCity = Sheets(sheetName).Cells(i, Sheets("References").Range("city").Value)
                    .HomeAddressCountry = "DefaultCountry"
                    .HomeAddressPostalCode = Sheets(sheetName).Cells(i, Sheets("References").Range("postcode").Value)
                    .HomeAddressStreet = Sheets(sheetName).Cells(i, Sheets("References").Range("street").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("street_number").Value)
                    .JobTitle = Sheets(sheetName).Cells(i, Sheets("References").Range("current_function").Value)
                    '.Language = Sheets(sheetName).Cells(i, Sheets("References").Range("speaking_language_info_list").Value)
                    .MobileTelephoneNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("mobile").Value)
                    .BusinessTelephoneNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("phone").Value)
                    .BusinessFaxNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("fax").Value)
                    .Department = Sheets(sheetName).Cells(i, Sheets("References").Range("division").Value)
                    .ManagerName = Sheets(sheetName).Cells(i, Sheets("References").Range("manager").Value)
                    .WebPage = "http://www.company.com"
                    If (InStr(.body, "Staff number: ") = 0) Then
                        .body = "Staff number: " & Sheets(sheetName).Cells(i, Sheets("References").Range("pnr").Value) & vbCrLf
                        .body = .body & "Recruitment date: " & Sheets(sheetName).Cells(i, Sheets("References").Range("recruitment_date").Value) & vbCrLf
                        .body = .body & "Education: " & Sheets(sheetName).Cells(i, Sheets("References").Range("educations").Value) & vbCrLf
                        .body = .body & "Languages: " & Sheets(sheetName).Cells(i, Sheets("References").Range("speaking_language_info_list").Value) & vbCrLf
                        .body = .body & "Specialty skills: " & Sheets(sheetName).Cells(i, Sheets("References").Range("specialty_skills").Value) & vbCrLf
'                    Else
' keeps original body note, but duplicates all data if run mulitple times
'                        originalBody = .body
'                        .body = "Staff number: " & Sheets(sheetName).Cells(i, Sheets("References").Range("pnr").Value) & vbCrLf
'                        .body = .body & "Recruitment date: " & Sheets(sheetName).Cells(i, Sheets("References").Range("recruitment_date").Value) & vbCrLf
'                        .body = .body & "Education: " & Sheets(sheetName).Cells(i, Sheets("References").Range("educations").Value) & vbCrLf
'                        .body = .body & "Languages: " & Sheets(sheetName).Cells(i, Sheets("References").Range("speaking_language_info_list").Value) & vbCrLf
'                        .body = .body & "Specialty skills: " & Sheets(sheetName).Cells(i, Sheets("References").Range("specialty_skills").Value) & vbCrLf
'                        .body = .body & vbCrLf & originalBody
                    End If
                End With
                                        
            End If
            Sheets(sheetName).Cells(i, Sheets("References").Range("processed").Value).Value = "OK"
            olContactItem.Save
            itemsFound = itemsFound + 1
        Next
        
        If itemsFound = 0 Then
            Application.StatusBar = "Adding Outlook contact " & Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value) & ", " & Sheets(sheetName).Cells(i, Sheets("References").Range("pnr").Value)
            'Using the Items.Add Method to create a new Outlook contact item in the default Contacts folder.
            Set oCItem = oCFolder.Items.Add
            'display the new contact item form:
            'oCItem.Display
            'set properties of the new contact item:
            With oCItem
                    'update existing contact fields https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_properties.aspx
                    .firstName = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value)
                    .lastName = Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    If (Sheets(sheetName).Cells(i, Sheets("References").Range("birthdate").Value) <> "") Then
                        .Birthday = DateValue(Sheets(sheetName).Cells(i, Sheets("References").Range("birthdate").Value))
                    End If
                    .BusinessAddressStreet = "BusinessStreet 100"
                    .BusinessAddressCity = "BusinessCity"
                    .BusinessAddressCountry = "BusinessCountry"
                    .BusinessAddressPostalCode = "BusinessPostal"
                    .BusinessHomePage = "http://www.company.com"
                    'do not remove or duplicate existing categories, set some desired category when not yet set
                    If (InStr(.Categories, "Business") = 0) Then
                        .Categories = .Categories & ",Business"
                    End If
                    .CompanyName = "Company"
                    .Email1DisplayName = Sheets(sheetName).Cells(i, Sheets("References").Range("email").Value)
                    .FullName = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    .FileAs = Sheets(sheetName).Cells(i, Sheets("References").Range("first_name").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("last_name").Value)
                    '.gender = Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value)
'                    If (Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value) = "M") Then
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olMale
'                    ElseIf (Sheets(sheetName).Cells(i, Sheets("References").Range("gender").Value) = "F") Then
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olFemale
'                    Else
'                        .gender = Microsoft.Office.Interop.Outlook.OlGender.olUnspecified
'                    End If
                    .HomeAddressCity = Sheets(sheetName).Cells(i, Sheets("References").Range("city").Value)
                    .HomeAddressCountry = "DefaultCountry"
                    .HomeAddressPostalCode = Sheets(sheetName).Cells(i, Sheets("References").Range("postcode").Value)
                    .HomeAddressStreet = Sheets(sheetName).Cells(i, Sheets("References").Range("street").Value) & " " & Sheets(sheetName).Cells(i, Sheets("References").Range("street_number").Value)
                    .JobTitle = Sheets(sheetName).Cells(i, Sheets("References").Range("current_function").Value)
                    '.Language = Sheets(sheetName).Cells(i, Sheets("References").Range("speaking_language_info_list").Value)
                    .MobileTelephoneNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("mobile").Value)
                    .BusinessTelephoneNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("phone").Value)
                    .BusinessFaxNumber = Sheets(sheetName).Cells(i, Sheets("References").Range("fax").Value)
                    .Department = Sheets(sheetName).Cells(i, Sheets("References").Range("division").Value)
                    .ManagerName = Sheets(sheetName).Cells(i, Sheets("References").Range("manager").Value)
                    .WebPage = "http://www.company.com"
                    If (InStr(.body, "Staff number: ") = 0) Then
                        .body = "Staff number: " & Sheets(sheetName).Cells(i, Sheets("References").Range("pnr").Value) & vbCrLf
                        .body = .body & "Recruitment date: " & Sheets(sheetName).Cells(i, Sheets("References").Range("recruitment_date").Value) & vbCrLf
                        .body = .body & "Education: " & Sheets(sheetName).Cells(i, Sheets("References").Range("educations").Value) & vbCrLf
                        .body = .body & "Languages: " & Sheets(sheetName).Cells(i, Sheets("References").Range("speaking_language_info_list").Value) & vbCrLf
                        .body = .body & "Specialty skills: " & Sheets(sheetName).Cells(i, Sheets("References").Range("specialty_skills").Value) & vbCrLf
                    End If
            End With
                
                'mark as done
                Sheets(sheetName).Cells(i, Sheets("References").Range("processed").Value).Value = "OK"
            fullFilePath = Replace(Sheets(sheetName).Cells(i, Sheets("References").Range("pictureurl").Value), "http://someimageurl", "C:\Users\username\Documents\Pictures\profilephotos\")
            If FileExists(fullFilePath) Then
                oCItem.AddPicture (fullFilePath)
            Else
                'MsgBox ("Missing picture: " + fullFilePath + ", for member: " + oCItem.firstName + " " + oCItem.lastname)
                Debug.Print ("Missing picture: " & fullFilePath + ", for member: " & oCItem.firstName & " " & oCItem.lastName)
            End If
            'close the new contact item form after saving:
            'The numerical value of olSave is 0. The following code has replaced the Outlook's built-in constant olSave by its numerical value 0.
            oCItem.Close 0
        End If
    Next i
    
    'quit the Oulook application:
    oApplOutlook.Quit
    
    Application.StatusBar = "Outlook " & lLastRow & " contacts exported"
    Application.ScreenUpdating = True
    
    'clear the variables:
    Set oApplOutlook = Nothing
    Set oNsOutlook = Nothing
    Set oCFolder = Nothing
    Set oDelFolder = Nothing
    Set oCItem = Nothing
    Set oDelItems = Nothing
    Set olContactItem = Nothing
    Set olItems = Nothing
    
    MsgBox "Successfully Exported Worksheet Data to the Default Outlook Contacts Folder."
    
     

End Sub

Function FileExists(fullFileName As String) As Boolean
    FileExists = VBA.Len(VBA.Dir(fullFileName)) > 0
End Function