using ExportingDataExcel.Bean;
using ExportingDataExcel.Bean.Enumeracoes;
using ExportingDataExcel.Business;
using ExportingDataExcel.Dao;
using ExportingDataExcel.Methods;
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
    public partial class CadastrarUsuarios : Form
    {
        public CadastrarUsuarios()
        {
            InitializeComponent();
        }

        private void btnIncluir_Click(object sender, EventArgs e)
        {
            UsuarioBusiness usuarioBusiness = new UsuarioBusiness();
            Usuario usuario = new Usuario();

            string nome = txtNome.Text;
            string idade = txtIdade.Text;
            char sexo;
            double salario = Convert.ToDouble(txtSalario.Text);

            if (rdbMasculino.Checked)
                sexo = (char)TSexo.Masculino;
            else
                sexo = (char)TSexo.Feminino;

            usuario.Nome = nome;
            usuario.Idade = idade;
            usuario.Sexo = sexo;
            usuario.Salario = salario;

            usuarioBusiness.Insert(usuario);

            ClearFields();
            PopulaGridView();
        }

        private void ClearFields()
        {
            txtNome.Text = string.Empty;
            txtIdade.Text = string.Empty;
            txtSalario.Text = string.Empty;
            rdbMasculino.Checked = true;

            txtNome.Focus();
        }

        private void PopulaGridView()
        {
            dgvUsuarios.DataSource = new UsuarioBusiness().GetAll();
            SetPropsInColumns();
        }

        private void SetPropsInColumns()
        {
            dgvUsuarios.Columns["IdUsuario"].Visible = false;
            dgvUsuarios.Columns["Salario"].DefaultCellStyle.Format = "C2";

        }

        private void CadastroUsuarios_Load(object sender, EventArgs e)
        {
            PopulaGridView();
        }

        private void txtPesquisa_TextChanged(object sender, EventArgs e)
        {
            PopulaGridView();

            if (txtPesquisa.Text.Length > 0)
            {
                string textoDigitado = txtPesquisa.Text;
                List<Usuario> listaUsuarios = dgvUsuarios.DataSource as List<Usuario>;

                dgvUsuarios.DataSource = listaUsuarios
                    .Where(user => user.Nome.ToLower().Contains(textoDigitado.ToLower())).ToList();
            }
        }

        private void dgvUsuarios_KeyUp(object sender, KeyEventArgs e)
        {
            if(Keys.F2 == e.KeyCode)
            {
                int idUsuario = int.Parse(dgvUsuarios.SelectedRows[0]
                .Cells["IdUsuario"]
                .Value.ToString());

                List<Usuario> listaUsuarios = dgvUsuarios.DataSource as List<Usuario>;

                Usuario user = listaUsuarios
                    .Where(u => u.IdUsuario.Equals(idUsuario)).First();

                AlterarUsuarios alterarUsuarios = new AlterarUsuarios
                {
                    usuario = user
                };

                DialogResult result = alterarUsuarios.ShowDialog();

                if (result == DialogResult.OK)
                    PopulaGridView();
            }

            else if (Keys.Delete == e.KeyCode)
            {
                int idUsuario = int.Parse(dgvUsuarios.SelectedRows[0]
                .Cells["IdUsuario"]
                .Value.ToString());

                new UsuarioBusiness().Delete(idUsuario);

                PopulaGridView();
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            List<Usuario> usuarios = dgvUsuarios.DataSource as List<Usuario>;

            if (usuarios == null || usuarios.Count == 0)
            {
                MessageBox.Show("Não há dados a serem exportados.", "Aviso",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (chkArquivoXlsx.Checked)
                    ExportFile.ExportToExcelUsingMicrosoft(usuarios);
                else if (chkArquivoTexto.Checked)
                    ExportFile.ExportToExcelUsingTextFile(usuarios);
            }
        }
    }
}
