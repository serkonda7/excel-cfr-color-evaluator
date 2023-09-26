' Copyright (c) 2023-present Lukas Neubert.
' This Source Code is subject to the terms of the Mozilla Public License 2.0.

Option Explicit

Sub EvaluateSheet()
	Dim ws As Worksheet: Set ws = ActiveSheet
	Dim row As Range
	Dim cell As Range

	Set row = ws.Range("D3")

	Do Until row.Value = ""
		For Each cell in row.Offset(1, 0).Resize(ws.Rows.Count - row.Row).Cells
			If cell.DisplayFormat.Interior.Color = RGB(255, 255, 255) Then
				Exit For
			End If

			Debug.Print cell.Value
		Next cell

		Set row = row.Offset(0, 1)
	Loop
End Sub
