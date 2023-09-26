' This file is part of https://github.com/serkonda7/excel-cfr-color-evaluator.
' Copyright (c) 2023-present Lukas Neubert.
' This Source Code is subject to the terms of the Mozilla Public License 2.0.

Option Explicit

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
		percent = nr_green / total
		cell.Value = percent * 100 & "%" & " (" & nr_green & ")"

		Set row = row.Offset(0, 1)
	Loop
End Sub
