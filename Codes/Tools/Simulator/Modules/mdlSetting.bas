Attribute VB_Name = "mdlSetting"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Simpan Setting] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub SetSave(NamaSetting As String, Nilai As String)
    Dim s_SetName As String, s_Nilai As String
    
    s_SetName = Crypt(NamaSetting, 6)
    s_Nilai = Crypt(Nilai, 6)
    'SaveSetting "h065fdc7s", "penalaan", Namasetting, nilai
    SaveString HKEY_CLASSES_ROOT, "odexstring\shell", s_SetName, s_Nilai
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function SetGet(NamaSetting As String, Optional Default As Variant = "") As Variant
    Dim s_SetName As String
        
    s_SetName = Crypt(NamaSetting, 6)
    'SetGet = GetSetting("h065fdc7s", "penalaan", Namasetting)
    SetGet = Crypt(GetString(HKEY_CLASSES_ROOT, "odexstring\shell", s_SetName), 6)
    If SetGet = "" Then SetGet = Default
End Function
