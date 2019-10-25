using ExportingDataExcel.Bean;
using ExportingDataExcel.Bean.Enumeracoes;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportingDataExcel.Methods
{
    public class ExportFile
    {
        // Método de exportar arquivo excel usando a dll Microsoft.Office.Interop.Excel
        public static void ExportToExcelUsingMicrosoft<T>(T item)
        {
            // Representa o programa excel instalado no PC do usuário
            Application app = new Application();
            // Representa um arquivo excel
            Workbook excelFile = app.Workbooks.Add();
            // Representa as abas dos excel
            Sheets sheets = excelFile.Worksheets;

            // Cria uma aba em específico
            Worksheet usuarioSheet = sheets.Add() as Worksheet;
            usuarioSheet.Name = "Usuários";

            PopulateSheet(usuarioSheet, item);

            // Se houver mais de uma aba, a de usuários terá o foco.
            usuarioSheet.Activate();

            usuarioSheet.Columns.AutoFit();

            app.Visible = true;

            SaveExcelFile(excelFile, "");
        }


        private static void PopulateSheet<T>(Worksheet worksheet, T item)
        {
            CreateHeader(worksheet, item);
            PopulateCellWithValues(worksheet, item);
        }

        private static void CreateHeader<T>(Worksheet sheet, T item)
        {
            PropertyDescriptorCollection listOfProperties = null;
            // Vetor que representa as colunas do meu arquivo excel;
            string[] columns = {"A", "B", "C", "D", "E", "F", "G", "H" };
            // Essa variável será usada para acessar o meu vetor columns dinamicamente.
            int columnIndex = 0;

            // Essa variável representa a primeira linha do arquivo excel
            int header = 1;

            // Verifica se o meu objeto genérico é uma lista de usuários
            if (item is List<Usuario>)
            {
                // Converte o meu objeto genérico para List<Usuario>
                List<Usuario> usuarios = item as List<Usuario>;

                // Obtém as propriedades do objeto genérico passado por parâmetro
                listOfProperties = TypeDescriptor.GetProperties(usuarios[0]);

                // Itera cada propriedade do meu objeto, pega o nome e seta na célula do excel
                foreach (PropertyDescriptor prop in listOfProperties)
                {
                    if (prop.Name.Equals("IdUsuario"))
                        continue;

                    sheet.Cells[header, columns[columnIndex]] = prop.DisplayName;
                    columnIndex++;
                }

            }
            else if (item is Usuario)
            {
                // Converte o meu objeto genérico para Usuario
                Usuario usuario = item as Usuario;

                // Obtém as propriedades do objeto genérico passado por parâmetro
                listOfProperties = TypeDescriptor.GetProperties(usuario);

                // Itera cada propriedade do meu objeto, pega o nome e seta na célula do excel
                foreach (PropertyDescriptor prop in listOfProperties)
                {
                    if (prop.Name.Equals("IdUsuario"))
                        continue;

                    sheet.Cells[header, columns[columnIndex]] = prop.DisplayName;
                    columnIndex++;
                }
            }
        }

        private static void PopulateCellWithValues<T>(Worksheet sheet, T item)
        {
            PropertyDescriptorCollection listOfProperties = null;
            // Vetor que representa as colunas do meu arquivo excel;
            string[] columns = { "A", "B", "C", "D", "E", "F", "G", "H" };
            // Essa variável será usada para acessar o meu vetor columns dinamicamente.
            int columnIndex = 0;
            // Essa variável representa cada linha do meu arquivo excel.
            // Acaba sendo 2, porque a linha 1 pertence ao cabeçalho.
            int rowIndex = 2;

            if (item is Usuario)
            {
                Usuario usuario = item as Usuario;
                listOfProperties = TypeDescriptor.GetProperties(usuario);

                foreach (PropertyDescriptor prop in listOfProperties)
                {
                    if (prop.Name.Equals("IdUsuario"))
                        continue;

                    sheet.Cells[rowIndex, columns[columnIndex]] = prop.GetValue(usuario);
                    columnIndex++;
                }            
            }
            else if (item is List<Usuario>)
            {
                List<Usuario> usuarios = item as List<Usuario>;
                listOfProperties = TypeDescriptor.GetProperties(usuarios[0]);

                foreach (Usuario user in usuarios)
                {
                    foreach (PropertyDescriptor prop in listOfProperties)
                    {
                        if (prop.Name.Equals("IdUsuario"))
                            continue;

                        if (!prop.Name.Equals("Sexo"))
                        {
                            sheet.Cells[rowIndex, columns[columnIndex]] = prop.GetValue(user);
                            columnIndex++;
                        }
                        else
                        {
                            if (user.Sexo == (char)TSexo.Masculino)
                                sheet.Cells[rowIndex, columns[columnIndex]] = TSexo.Masculino.ToString();
                            else if (user.Sexo == (char)TSexo.Feminino)
                                sheet.Cells[rowIndex, columns[columnIndex]] = TSexo.Feminino.ToString();

                            columnIndex++;
                        }                        
                    }

                    columnIndex = 0;
                    rowIndex++;
                }
            }
        }

        public static void SaveExcelFile(Workbook excelFile, string path)
        {

        }

        public static void ExportToExcelUsingTextFile<T>(List<T> usuarios)
        {
            
        }

        
    }
}
