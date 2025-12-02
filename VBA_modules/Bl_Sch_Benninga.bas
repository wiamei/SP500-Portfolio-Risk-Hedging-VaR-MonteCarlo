Attribute VB_Name = "Bl_Sch_Benninga"
Function dOne(Stock, Exercise, Time, interest, sigma)
    dOne = (Log(Stock / Exercise) + interest * Time) / (sigma * Sqr(Time)) _
        + 0.5 * sigma * Sqr(Time)
End Function
'This is the BS call option price
Function BSCall(Stock, Exercise, Time, interest, sigma)
    BSCall = Stock * Application.NormSDist(dOne(Stock, Exercise, _
        Time, interest, sigma)) - Exercise * Exp(-Time * interest) * _
     Application.NormSDist(dOne(Stock, Exercise, Time, interest, sigma) _
      - sigma * Sqr(Time))
End Function
'The BS put option price uses put-call parity
Function BSPut(Stock, Exercise, Time, interest, sigma)
    BSPut = BSCall(Stock, Exercise, Time, interest, sigma) + _
    Exercise * Exp(-Time * interest) - Stock
End Function


