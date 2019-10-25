using ExportingDataExcel.Bean;
using ExportingDataExcel.Bean.Enumeracoes;
using ExportingDataExcel.Business;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportingDataExcel.Views
{
    public partial class AlterarUsuarios : Form
    {
        public Usuario usuario { get; set; }

        public AlterarUsuarios()
        {
            InitializeComponent();
        }

        private void btnAlterar_Click(object sender, EventArgs e)
        {
            if (usuario != null)
            {
                usuario.Nome = txtNome.Text;
                usuario.Idade = txtIdade.Text;
                usuario.Salario = double.Parse(txtSalario.Text);
                usuario.Sexo = rdbMasculino.Checked ? (char)TSexo.Masculino : (char)TSexo.Feminino;

                new UsuarioBusiness().Update(usuario);

                MessageBox.Show("Usuário alterado com sucesso.", "Sucesso",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();

                this.DialogResult = DialogResult.OK;
            }
        }

        private void AlterarUsuarios_Load(object sender, EventArgs e)
        {
            PopulaFields();
        }

        private void PopulaFields()
        {
            if (usuario != null)
            {
                txtNome.Text = usuario.Nome;
                txtIdade.Text = usuario.Idade;
                txtSalario.Text = usuario.Salario.ToString();
                rdbMasculino.Checked = usuario.Sexo == (char)TSexo.Masculino;
                rdbFeminino.Checked = usuario.Sexo == (char)TSexo.Feminino;
            }
        }
    }
}
