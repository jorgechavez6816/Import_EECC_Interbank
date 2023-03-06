'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()	'D:\RUC1\DATA\Archivos fuente.ILB\2022_ITB.pdf
	Call AppendField()	'G_ITB2022.IMD
	Call Summarization()	'G_ITB2022.IMD
	Client.CloseAll
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "G_ITB2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "G.1_Resumen_ITB.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "G_ITB2022.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\ITB_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_ITB.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("G_ITB2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PERIODO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """2022""+@RIGHT(FECHA_OP;2)"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function


' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("G_ITB2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToSummarize "CUENTA"
	task.AddFieldToInc "MONEDA"
	task.AddFieldToTotal "CARGO_ABONO"
	dbName = "G.1_Resumen_ITB.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

