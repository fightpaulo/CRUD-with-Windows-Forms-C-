using Dapper;
using ExportingDataExcel.Bean;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportingDataExcel.Dao
{
    public class UsuarioDao
    {
        private SqlConnection conn;

        public UsuarioDao()
        {
            conn = DataBaseHelper.GetConnection();
        }

        public void Insert(Usuario usuario)
        {
            StringBuilder sql = new StringBuilder();

            sql.Append("INSERT INTO usuario ");

            sql.Append("(txt_nome ");
            sql.Append(",txt_idade ");
            sql.Append(",txt_sexo ");
            sql.Append(",float_salario) ");

            sql.Append("VALUES ");

            sql.Append("(@nome ");
            sql.Append(",@idade ");
            sql.Append(",@sexo ");
            sql.Append(",@salario) ");

            conn.Execute(sql.ToString(), usuario);

            conn.Close();
        }

        public List<Usuario> GetAll()
        {
            StringBuilder sql = new StringBuilder();
            List<Usuario> listaUsuarios;

            sql.Append("SELECT ");

            sql.Append("int_user_id [IdUsuario] ");
            sql.Append(",txt_nome [Nome] ");
            sql.Append(",txt_idade [Idade] ");
            sql.Append(",txt_sexo [Sexo] ");
            sql.Append(",float_salario [Salario] ");

            sql.Append("FROM ");

            sql.Append("usuario");

            listaUsuarios = conn.Query<Usuario>(sql.ToString()).ToList();

            conn.Close();

            return listaUsuarios;
        }

        public Usuario GetUser(int idUser)
        {
            StringBuilder sql = new StringBuilder();
            Usuario usuario;

            sql.Append("SELECT ");

            sql.Append("int_user_id [IdUsuario] ");
            sql.Append(",txt_nome [Nome] ");
            sql.Append(",txt_idade [Idade] ");
            sql.Append(",txt_sexo [Sexo] ");
            sql.Append(",float_salario [Salario] ");

            sql.Append("FROM ");

            sql.Append("usuario ");

            sql.Append("WHERE int_user_id = @idUser");

            usuario = conn.QuerySingle<Usuario>(sql.ToString(), new { idUser });

            conn.Close();

            return usuario;
        }

        public void Update(Usuario usuario)
        {
            StringBuilder sql = new StringBuilder();

            sql.Append("UPDATE usuario ");
            sql.Append("SET ");
            sql.Append("txt_nome = @nome ");
            sql.Append(",txt_idade = @idade ");
            sql.Append(",txt_sexo = @sexo ");
            sql.Append(",float_salario = @salario ");

            sql.Append("WHERE int_user_id = @id ");

            conn.Execute(sql.ToString(), new
            {
                @id = usuario.IdUsuario,
                @nome = usuario.Nome,
                @idade = usuario.Idade,
                @sexo = usuario.Sexo,
                @salario = usuario.Salario
            });

            conn.Close();
        }

        public void Delete(int idUser)
        {
            StringBuilder sql = new StringBuilder();

            sql.Append("DELETE FROM usuario ");
            sql.Append("WHERE int_user_id = @id");

            conn.Execute(sql.ToString(), new { @id = idUser });
        }
    }
}
