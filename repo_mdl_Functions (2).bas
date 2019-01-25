Attribute VB_Name = "mdl_Functions"
Option Explicit

Public Function KPI_HG(ByVal Actual As Double, ByVal Target As Double) As Double

    KPI_HG = (Actual / Target)

End Function

Public Function KPI_Points(ByVal KPI_HGcl As Double, ByVal Limit1 As Double, ByVal Limit2 As Double, ByVal Limit3 As Double, ByVal Points1 As Double, ByVal Points2 As Double, ByVal Points3 As Double) As Double

If KPI_HGcl >= Limit1 Then KPI_Points = Points1
    If KPI_HGcl < Limit1 And KPI_HGcl >= Limit2 Then KPI_Points = Points2
        If KPI_HGcl < Limit2 And KPI_HGcl >= Limit3 Then KPI_Points = Points3
            If KPI_HGcl < Limit3 Then KPI_Points = 0
End Function

Public Function KPI_LmtvsPnt(ByVal KPI_HGcl As Double, ByVal NoOfLimits As Double, limit As Variant, Points As Variant) As Double
' KPI_HGcl lookup Target Realization value for KPI, NoOfLimits how many points brackets are present excluding no points, limit is the array of minimum target realization limits, Points is the array of corresponding Points
Dim i As Double
For i = 1 To NoOfLimits
    If KPI_HGcl >= limit(i) Then Exit For
    If KPI_HGcl < limit(i) Then
    End If
Next i
KPI_LmtvsPnt = Points(i)