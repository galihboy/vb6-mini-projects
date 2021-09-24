Attribute VB_Name = "Module1"
Option Explicit
' Galih Hermawan @ http://galih.eu #3-8-2013

Function RandomPositive(Lowerbound As Long, Upperbound As Long)
Randomize
RandomPositive = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function
