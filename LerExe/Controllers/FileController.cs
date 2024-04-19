using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System;
using OfficeOpenXml; // Biblioteca EPPlus para manipulação de arquivos Excel

namespace Integracao.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class FileController : ControllerBase
    {
        private readonly ILogger<FileController> _logger;

        public FileController(ILogger<FileController> logger)
        {
            _logger = logger;
            // Defina o contexto de licença do EPPlus
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        [HttpPost]
        public IActionResult Post([FromHeader(Name = "FilePath")] string filePath)
        {
            // Verifica se o caminho do arquivo foi fornecido
            if (string.IsNullOrEmpty(filePath))
            {
                return BadRequest("Caminho do arquivo não fornecido.");
            }

            try
            {
                // Verifica se o arquivo existe no caminho fornecido
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound("Arquivo não encontrado.");
                }

                // Ler o arquivo Excel e converter para uma lista de objetos
                var data = ReadExcelContent(filePath);

                // Retorna uma resposta HTTP bem-sucedida com o conteúdo lido do arquivo Excel
                return Ok(new { resultados = data });
            }
            catch (Exception ex)
            {
                // Trata exceções que possam ocorrer durante o processamento da requisição
                return StatusCode(500, $"Erro interno do servidor: {ex.Message}");
            }
        }

        private List<List<Dictionary<string, string>>> ReadExcelContent(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = FindWorksheet(workbook, "Proposta_(Uso_Concessionária)");

                if (worksheet == null)
                {
                    worksheet = FindWorksheet(workbook, "Proposta_(Uso_Concessionária)2");

                    if (worksheet == null)
                    {
                        // Se não encontrar as planilhas específicas, lê todos os dados da primeira planilha
                        worksheet = workbook.Worksheets[0];
                    }
                }

                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                var data = new List<List<Dictionary<string, string>>>();

                for (int row = 1; row <= rowCount; row++)
                {
                    var rowData = new List<Dictionary<string, string>>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                        var columnName = ((char)('A' + col - 1)).ToString() + row;
                        if (rowData.Count > 0)
                        {
                            var lastColumnData = rowData[rowData.Count - 1];
                            lastColumnData[columnName] = cellValue;
                        }
                        else
                        {
                            var columnData = new Dictionary<string, string> { { columnName, cellValue } };
                            rowData.Add(columnData);
                        }
                    }
                    data.Add(rowData);
                }

                return data;
            }
        }

        private ExcelWorksheet FindWorksheet(ExcelWorkbook workbook, string worksheetName)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (worksheet.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }
            return null;
        }
    }
}