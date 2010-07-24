Attribute VB_Name = "Module1"
Public Function InitReg(Nama) As String
    a1 = Len(Nama)
    a2 = a1 * 5
    a3 = a2 * a1
    a4 = Left(Nama, 1)
    a5 = Right(Nama, 1)
    a6 = "0v10"
    
    genstr = a1 & a2 & a3 & a4 & a5 & a6
    
    InitReg = genstr
 End Function
