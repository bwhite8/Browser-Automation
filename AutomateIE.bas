Attribute VB_Name = "AutomateIE"
'
' Title:      AutomateIE
' Purpose:    Automates IE to login to a website
' Dependency: Any Windows OS, IE
' Author:     Brian White
' Date:       12/2014
'

Sub AutomateIE()
  Dim appIE, CoN, UsN, Pw, Element, btnInput, ElementCol, Link As Object
  Dim sURL, strCountBody, TextIWant, jsLogin, jsShipTo As String
  Dim lStartPos, lEndPos As Long

  Set appIE = CreateObject("InternetExplorer.Application")

  ' Set base URL here
  sURL = "someurl"

  With appIE
     .Navigate sURL
     .Visible = True
  End With

  ' Wait for page to load
  Do While appIE.Busy
  Loop

  ' If input is required for login...
  Set CoN = appIE.Document.getElementsByName("SomeWebElement")
  ' Input Field 1
  If Not CoN Is Nothing Then
     CoN(0).Value = "SomeInput"
  End If

  Set UsN = appIE.Document.getElementsByName("SomeWebElement")
  ' Input Field 2
  If Not UsN Is Nothing Then
     UsN(0).Value = "SomeInput"
  End If

  Set Pw = appIE.Document.getElementsByName("SomeWebElement")
  ' Input Field 3
  If Not Pw Is Nothing Then
     Pw(0).Value = "SomeInput"
  End If

  ' Trigger Login Javascript
  jsLogin = "javascript:subDoLogonError()"
  SendKeys "%d"
  SendKeys jsLogin
  'SendKeys "{enter}"
  Application.Wait (Now + TimeValue("00:00:03"))

  ' Wait for page to load
  Do While appIE.Busy
  Loop


  Set appIE = Nothing
  Set CoN = Nothing
  Set UsN = Nothing
  Set Pw = Nothing

End Sub
