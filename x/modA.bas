Attribute VB_Name = "modA"
Global intPlayerMode As Integer
Global intDifficulty As Integer
Global pShip() As Player
Global intGoal As Integer
Global intOutcome As Integer
Global intShields As Integer

Type Enemy

    Destroyed As Boolean
    Passed As Boolean
    
    FireCD As Integer
    OnFireCD As Integer
    Health As Integer
    MaxHealth As Integer
    EnemyType As Integer
    EnemyIndex As Integer
    ExplosionCounter As Integer
    Speed As Integer
    
End Type

Type Player
    
    Destroyed As Boolean
    
    CannonType As Integer
    FireCD As Integer
    OnFireCD As Integer
    FiringCost As Integer
    FirePower As Integer
    Health As Integer
    Energy As Integer
    MissileSpeed As Integer
    XSpeed As Integer
    YSpeed As Integer
    Score As Integer
        
End Type

Function fnRndInt(Lowerbound As Integer, Upperbound As Integer) As Integer
 
    fnRndInt = Int(Rnd * (Upperbound - Lowerbound + 1) + Lowerbound)

End Function



