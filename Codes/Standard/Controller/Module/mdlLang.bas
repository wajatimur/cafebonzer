Attribute VB_Name = "MdlLanguage"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlLanguage
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public Crnc As String
Public SL(12) As String
Public VS(10, 10) As String
Public ST(10, 10) As String

Public Const CbMsgApp = "CafeBonzer"
Public Const CbMsgWarn = "CafeBonzer Warning !"


Sub LangLoad()
    Crnc = "RM"
    
    SL(1) = "Login"
    SL(2) = "Logout"
    SL(3) = "Locking %sa%"
    SL(4) = "Unlocking %pa1% without login customer !"
    SL(5) = "Acessing to Setting !"
    SL(6) = "Acessing to Statistic!"
    SL(7) = "Cancelling usage for customer %cn% on %sa% !"
    SL(8) = "Rebooting %sa%"
    SL(9) = "Shutdown %sa%"
    SL(10) = "Mass reboot"
    SL(11) = "Mass shutdown"
    SL(12) = "Closing server"
    
    
    VS(0, 0) = "Delete"
    VS(0, 1) = "None"
    VS(0, 2) = "Hour"
    VS(0, 3) = "Minute"
    VS(0, 4) = "Second"
    
    VS(1, 0) = "Default"
    VS(1, 1) = "Used"
    VS(1, 2) = "Unused"
    VS(1, 3) = "End"
    
    VS(2, 0) = "Terminal"
    VS(2, 1) = "NA"
    VS(2, 2) = "Quantity"
    
    ST(0, 0) = "Please enter number only !"
    ST(0, 1) = "Please enter your main password !"
    ST(0, 2) = "Please enter the same password !"
    ST(0, 3) = "The scheme is already exist, please select others name !"
    ST(0, 4) = "The overhead is already exist, please select others name !"

    ST(1, 0) = "Shutdown CafeBonzer ?"
    ST(1, 1) = "You dont have an access to this part, Please contact your administrator !"
    ST(1, 2) = LangPrcs("Shutdown CafeBonzer ?%nl% All connection will be terminate !")
    ST(1, 3) = "Module system is not found or not installed properly !"
    ST(1, 4) = "Access Denied"
    
    ST(2, 0) = "Terminal is already in used !"
    ST(2, 1) = "Cancel PC usage ?"
    ST(2, 2) = "Cannot cancel if used more than 10 seconds !"
    ST(2, 3) = "There's no transaction have been made, please add items !"
    ST(2, 4) = "connected at"

    ST(3, 0) = "The trial period is expired, please contact us for full version"
    ST(3, 1) = "Thanks for purchasing the CafeBonzer !"
    ST(3, 2) = LangPrcs("Are you sure to transfer your liscense ?%nl%If yes, please insert your DiskKey !")
    ST(3, 3) = "Unable to activate CafeBonzer. Application will exit now !"
End Sub

Function LangPrcs(Text As String) As String
    ' %sa% = Selected Agent
    ' %cn% = Customer Name
    
    ' %nl% = Newline
    ' %tb% = Tab
    ' %exdt% = Date
    ' %extm% = Time
    
    
    Text = Replace(Text, "%cn%", SelSubItm(2))
    Text = Replace(Text, "%sa%", SelText)
    Text = Replace(Text, "%nl%", vbNewLine)
    Text = Replace(Text, "%tb%", vbTab)
    Text = Replace(Text, "%exdt%", Date)
    LangPrcs = Replace(Text, "%extm%", Time)
End Function
