' SPDX-FileCopyrightText: 2023-present Lukas Neubert <lukas.neubert@proton.me>
' SPDX-License-Identifier: MPL-2.0

Option Explicit

Private Const RATIO_FOR_GREEN = 0.9

Private Const WHITE = 16777215
Private Const GREEN = 13561798
Private Const YELLOW = 10284031
Private Const RED = 13551615

Sub EvaluateSheet()
	Dim ws As Worksheet: Set ws = ActiveSheet
	Dim row As Range
	Dim cell As Range

	Dim total As Integer
	Dim nr_green As Integer
	Dim percent As Double

	Set row = ws.Range("D3")
	Do Until row.Value = ""

		For Each cell in row.Offset(1, 0).Resize(ws.Rows.Count - row.Row).Cells
			If cell.DisplayFormat.Interior.Color = WHITE Then
				Exit For
			End If

			If cell.DisplayFormat.Interior.Color = GREEN Then
				nr_green = nr_green + 1
			End If
		Next cell

		total = cell.Row - row.Row - 1
		cell.NumberFormat = "@"
		cell.Value = nr_green & "/" & total

		Set cell = cell.Offset(1, 0)
		percent = nr_green / total
		cell.NumberFormat = "0%"
		cell.Value = percent
		If percent >= RATIO_FOR_GREENRATIO_FOR_GREEN Then
			cell.Interior.Color = GREEN
		Else
			cell.Interior.Color = RED
		End If

		nr_green = 0
		Set row = row.Offset(0, 1)
	Loop
End Sub
