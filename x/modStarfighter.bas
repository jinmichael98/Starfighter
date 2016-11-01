Attribute VB_Name = "modStarfighter"
Type Enemy
    Hit As Boolean
    Speed As Integer
    Health As Integer

End Type

Function fnRndInt(Lowerbound As Integer, Upperbound As Integer) As Integer
 
    fnRndInt = Int(Rnd * (Upperbound - Lowerbound + 1) + Lowerbound)

End Function
