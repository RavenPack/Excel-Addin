Attribute VB_Name = "UnitTests"

Public Function TestRPGetDailyValue()
    TestRPGetDailyValue = _
        RPGetDailyValue("", _
                        "1774047ADEC0FFD8DB435C6ADC6CA3B4", _
                        "228D42", "strength_91d", _
                        Date, "UTC")
End Function

Public Function TestRPMapEntity()
    TestRPMapEntity = RPMapEntity("", "Amazon", "COMP")
End Function

Public Function TestRPGetDailyEntitySentiment()
    TestRPGetDailyEntitySentiment = RPGetDailyEntitySentiment("", "228D42", Now())
End Function


