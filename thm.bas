Attribute VB_Name = "THM"
' Fast Approximate Transient Hyperbolic Model
' Copyright (C) 2018 David S. Fulford

' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.

' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.

' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301
' USA

Option Explicit

Public Type ThmParams
  qi As Double
  Di As Double
  bi As Double
  bf As Double
  telf As Double

  t1 As Double
  t2 As Double
  t3 As Double
  t_term As Double

  b1 As Double
  b2 As Double
  b3 As Double
  b_term As Double

  D1 As Double
  D2 As Double
  D3 As Double
  D_term As Double

  q1 As Double
  q2 As Double
  q3 As Double
  q_term As Double

  G1 As Double
  G2 As Double
  G3 As Double
  G_term As Double

End Type

Const e_ As Double = 2.71828182845905
Const DaysPerYear As Double = 365.25
Const DaysPerMonth As Double = 365.25 / 12#


Private Sub Precalculate(params As ThmParams)
  ' pre-calculate all segment initial conditions

  params.t1 = 0#
  params.t2 = params.telf * (e_ - 1#)
  params.t3 = params.telf * (e_ + 1#)

  params.b1 = params.bi
  params.b2 = params.bi - ((params.bi - params.bf) / e_)
  params.b3 = params.bf

  If Round(params.t_term, 5) < Round(params.t3, 5) Then
    'Not valid terminal segment parameters
    params.t_term = 0#
  End If

  params.D1 = ((1# - params.Di) ^ (-1# * params.bi) - 1#) / params.bi / DaysPerYear
  params.D2 = D_fn(params, params.t2)
  params.D3 = D_fn(params, params.t3)

  params.q1 = params.qi
  params.q2 = rate_fn(params, params.t2)
  params.q3 = rate_fn(params, params.t3)

  params.G1 = 0#
  params.G2 = N_fn(params, params.t2)
  params.G3 = N_fn(params, params.t3)

  If params.t_term > 0# And params.b_term < params.b3 Then
    params.D_term = D_fn(params, params.t_term)
    params.q_term = rate_fn(params, params.t_term)
    params.G_term = N_fn(params, params.t_term)
  End If

  If params.t_term = 0# And params.b_term > 0# Then
    params.D_term = -Log(1# - params.b_term) / DaysPerYear
    params.b_term = 0#
    params.t_term = params.t3 + (1# / params.D_term - 1# / params.D3) / params.b3
    params.q_term = rate_fn(params, params.t_term)
    params.G_term = N_fn(params, params.t_term)
  End If

End Sub


Public Function thm_array(time As Range, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Variant
  ' return an array of all parameters
  ' q(t) - D(t) - b(t) - N(t)

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  Dim i As Long
  Dim rowcount As Long
  Dim OutputArray() As Variant

  rowcount = time.Rows.Count

  ReDim OutputArray(1 To rowcount, 1 To 4)
  For i = 1 To rowcount
    OutputArray(i, 1) = rate_fn(params, time(i, 1))
    OutputArray(i, 2) = D_fn(params, time(i, 1))
    OutputArray(i, 3) = b_fn(params, time(i, 1))
    OutputArray(i, 4) = N_fn(params, time(i, 1))

  Next i

  thm_array = OutputArray

End Function


Public Function thm_b(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate hyperbolic parameter b(t) function

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  thm_b = b_fn(params, time)

End Function


Public Function thm_D(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate norminal decline D(t) function

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  thm_D = D_fn(params, time)

End Function


Public Function thm_D_eff(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
Dim b_calc As Double, D_calc As Double
  ' calculate secant effective D(t) function

  b_calc = thm_b(time, qi, Di, bi, bf, telf, b_term, t_term)
  D_calc = thm_D(time, qi, Di, bi, bf, telf, b_term, t_term)

  If b_calc = 0# Then
    thm_D_eff = 1# - Exp(-1# * D_calc * DaysPerYear)
  Else
    thm_D_eff = 1# - (1# + b_calc * D_calc * DaysPerYear) ^ (-1# / b_calc)
  End If

End Function


Public Function thm_rate(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate rate q(t) function

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  thm_rate = rate_fn(params, time)

End Function


Public Function thm_cum(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate cumulative volume N(t) function

  If time = 0# Then
    thm_cum = 0#
    Exit Function
  End If

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  thm_cum = N_fn(params, time)

End Function


Public Function thm_cum_qf(ratelimit As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate cumulative volume N(q) function given a minimimum rate cutoff
  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  Dim time As Double

  If ratelimit > params.q2 Then
    time = t_fn(ratelimit, params.q1, params.D1, params.b1)
  ElseIf ratelimit > params.q3 Then
    time = params.t2 + t_fn(ratelimit, params.q2, params.D2, params.b2)
  ElseIf t_term > 0# And ratelimit <= params.q_term Then
    time = params.t_term + t_fn(ratelimit, params.q_term, params.D_term, params.b_term)
  Else
    time = params.t3 + t_fn(ratelimit, params.q3, params.D3, params.b3)
  End If

  thm_cum_qf = N_fn(params, time)

End Function


Public Function thm_monthly_vol(time As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate the monthly volume N(t) - N(t - 1 month)

  Dim t_m1 As Double, dt As Double

  If time = 0# Then
    thm_monthly_vol = 0#
    Exit Function
  End If

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  t_m1 = time - DaysPerMonth
  If t_m1 < 0# Then t_m1 = 0#
  dt = (time - t_m1) / DaysPerMonth

  thm_monthly_vol = (N_fn(params, time) - N_fn(params, t_m1)) / dt

End Function


Public Function thm_eur_t(ecotime As Double, days_on As Double, wellcum As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate the well EUR accounting for actual produced volumes to a time limit

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  Dim forecast_eur As Double
  Dim produced As Double

  forecast_eur = thm_cum(ecotime, qi, Di, bi, bf, telf, b_term, t_term)
  produced = thm_cum(days_on, qi, Di, bi, bf, telf, b_term, t_term)

  thm_eur_t = wellcum + forecast_eur - produced

End Function


Public Function thm_eur_qf(ratelimit As Double, days_on As Double, wellcum As Double, qi As Double, Di As Double, bi As Double, bf As Double, _
  telf As Double, Optional b_term As Double, Optional t_term As Double) As Double
  ' calculate the well EUR accounting for actual produced volumes to a rate limit

  Dim params As ThmParams
  params.qi = qi
  params.Di = Di
  params.bi = bi
  params.bf = bf
  params.telf = telf
  params.t_term = t_term * DaysPerYear
  params.b_term = b_term

  Call Precalculate(params)

  Dim forecast_eur As Double
  Dim produced As Double

  forecast_eur = thm_cum_qf(ratelimit, qi, Di, bi, bf, telf, b_term, t_term)
  produced = thm_cum(days_on, qi, Di, bi, bf, telf, b_term, t_term)

  thm_eur_qf = wellcum + forecast_eur - produced

End Function


Private Function t_fn(ratelimit As Double, q As Double, D As Double, b As Double) As Double
  ' calculate the time to reach a rate limit t(q)

  If b = 0# Then
      t_fn = Log(ratelimit / q) / -D
  Else
      t_fn = ((ratelimit / q) ^ (-b) - 1#) / (D * b)
  End If

End Function


Private Function b_fn(ByRef params As ThmParams, time As Double) As Double
  ' calculate the hyperbolic parameter at a given time

  If params.t_term > 0# And time >= params.t_term Then
    b_fn = params.b_term
  ElseIf time >= params.t3 Then
    b_fn = params.b3
  ElseIf time >= params.t2 Then
    b_fn = params.b2
  Else
    b_fn = params.b1
  End If

End Function


Private Function D_fn(ByRef params As ThmParams, time As Double) As Double
  ' calculate the nominal decline at a given time

  If params.t_term > 0# And Round(time, 5) > Round(params.t_term, 5) Then
    D_fn = D_check(params.D_term, params.b_term, params.t_term, time)
  ElseIf time > params.t3 Then
    D_fn = D_check(params.D3, params.b3, params.t3, time)
  ElseIf time > params.t2 Then
    D_fn = D_check(params.D2, params.b2, params.t2, time)
  Else
    D_fn = D_check(params.D1, params.b1, 0#, time)
  End If

End Function


Private Function D_check(D As Double, b As Double, t0 As Double, time As Double) As Double
  ' handle various Arps cases

  On Error GoTo Err:
    D_check = 1# / (1# / D + b * (time - t0))
  Exit Function

Err:
  D_check = 0#
  On Error Resume Next

End Function


Private Function rate_fn(ByRef params As ThmParams, time As Double) As Double
  ' calculate the rate at a given time

  If params.t_term > 0# And Round(time, 5) > Round(params.t_term, 5) Then
    rate_fn = q_check(params.q_term, params.D_term, params.b_term, params.t_term, time)
  ElseIf time > params.t3 Then
    rate_fn = q_check(params.q3, params.D3, params.b3, params.t3, time)
  ElseIf time > params.t2 Then
    rate_fn = q_check(params.q2, params.D2, params.b2, params.t2, time)
  Else
    rate_fn = q_check(params.q1, params.D1, params.b1, 0#, time)
  End If

End Function


Private Function q_check(q As Double, D As Double, b As Double, t0 As Double, time As Double) As Double
  ' handle various Arps cases

  On Error GoTo Err:
  If D = 0# Then
    q_check = q
  ElseIf b = 0# Then
    q_check = q * Exp(-D * (time - t0))
  Else
    q_check = q / ((1# + b * D * (time - t0)) ^ (1# / b))
  End If
  Exit Function

Err:
  q_check = 0#
  On Error Resume Next

End Function


Private Function N_fn(ByRef params As ThmParams, time As Double) As Double
  ' calculate the cumulative volume function at a given time

  If params.t_term > 0# And Round(time, 5) > Round(params.t_term, 5) Then
    N_fn = params.G_term + N_check(params.q_term, params.D_term, params.b_term, params.t_term, time)
  ElseIf time > params.t3 Then
    N_fn = params.G3 + N_check(params.q3, params.D3, params.b3, params.t3, time)
  ElseIf time > params.t2 Then
    N_fn = params.G2 + N_check(params.q2, params.D2, params.b2, params.t2, time)
  Else
    N_fn = N_check(params.q1, params.D1, params.b1, 0#, time)
  End If

End Function


Private Function N_check(q As Double, D As Double, b As Double, t0 As Double, time As Double)
  ' handle various Arps cases

  On Error GoTo Err:
  If q < 0# Then
    N_check = 0#
  ElseIf D < 0# Then
    N_check = q * (time - t0) / 1000#
  ElseIf b = 0# Then
    N_check = q / D * -(Exp(-D * (time - t0)) - 1) / 1000#
  ElseIf (Abs(1# - b) = 0#) Then
    N_check = q / D * (Log(1# + D * (time - t0))) / 1000#
  Else
    N_check = q / ((1# - b) * D) * ((1# - (1# + b * D * (time - t0)) ^ (1# - (1# / b)))) / 1000#
  End If
  Exit Function

Err:
  N_check = 0#
  On Error Resume Next

End Function


