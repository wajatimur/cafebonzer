Attribute VB_Name = "mLang"
Public Crnc As String

Public SL(12) As String
Public VS(10) As String
Public MB(23) As String

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
    
    VS(1) = "None"
    VS(2) = "Normal"
    VS(3) = "Used"
    VS(4) = "Unused"
    VS(5) = "End"
    VS(6) = "ONLINE !"
    VS(7) = "Hour"
    VS(8) = "Minute"
    VS(9) = "connected at"
    VS(10) = "Please provide required information !"
    
    MB(1) = "The scheme is already exist, please select another name !"
    MB(2) = "Delete"
    MB(3) = "The trial period is expired, please contact us for full version"
    MB(4) = "Please enter number only !"
    MB(5) = LangPrcs("The ID number on this computer has been verified !%nl%Rollback Registration ? If yes, please insert your DiskKey !")
    MB(6) = "Thanks for purchasing the CafeBonzer !"
    MB(7) = "Wrong registration number !"
    MB(8) = "Please enter your main password !"
    MB(9) = "Please enter the same password !"
    MB(10) = "You dont have an access to this part, Please contact your administrator !"
    MB(11) = "Shutdown CafeBonzer ?"
    MB(12) = LangPrcs("Shutdown CafeBonzer ?%nl% All connection will be terminate !")
    MB(13) = "Trial period has expired, please contact us to purchase the full version !"
    MB(14) = "There's no transaction have been made, please add items !"
    MB(15) = "Please select an item to delete."
    MB(16) = "Groups doesn't exist !"
    MB(17) = "Please enter time value only !"
    MB(18) = "Are you sure to delete this group and all of it contents ?"
    MB(19) = "Please select a group to delete !"
    MB(20) = "Module ystem is not found or not installed properly !"
    MB(21) = "Terminal is already in used !"
    MB(22) = "Cancel PC usage ?"
    MB(23) = "Cannot cancel if used more than 10 seconds !"
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
    Text = Replace(Text, "%exdt%", Tarikh)
    LangPrcs = Replace(Text, "%extm%", Time)
End Function

Function LangParam(Text As String, ParamArray SeqParam()) As String

End Function
