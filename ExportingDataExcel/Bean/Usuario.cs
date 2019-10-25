using ExportingDataExcel.Bean.Enumeracoes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportingDataExcel.Bean
{
    public class Usuario
    {
        [DisplayName("ID Usuário")]
        public int IdUsuario { get; set; }
        public string Nome { get; set; }
        public string Idade { get; set; }
        public char Sexo { get; set; }
        [DisplayName("Salário")]
        public double Salario { get; set; }
    }
}
