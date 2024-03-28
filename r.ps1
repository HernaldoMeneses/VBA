# Caminho para o arquivo Excel (.xlsm)
$caminhoExcel = "L:\GitHub\VBA\Pasta1.xlsm"

# Iniciar o Excel
$excel = New-Object -ComObject Excel.Application

# Abrir o arquivo Excel
$workbook = $excel.Workbooks.Open($caminhoExcel)

# Executar o procedimento VBA do módulo externo
$excel.Run("HelloWorld")

# Fechar o arquivo Excel sem salvar alterações
$workbook.Close($false)

# Fechar o Excel
$excel.Quit()

# Limpar os objetos do Excel
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel


