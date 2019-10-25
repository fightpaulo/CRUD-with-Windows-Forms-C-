using ExportingDataExcel.Bean;
using ExportingDataExcel.Dao;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportingDataExcel.Business
{
    public class UsuarioBusiness
    {
        private UsuarioDao usuarioDao;

        public UsuarioBusiness()
        {
            usuarioDao = new UsuarioDao();
        }

        public void Insert(Usuario usuario)
        {
            try
            {
                usuarioDao.Insert(usuario);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public List<Usuario> GetAll()
        {
            try
            {
                return usuarioDao.GetAll();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public Usuario GetUser(int idUser)
        {
            try
            {
                return usuarioDao.GetUser(idUser);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void Update(Usuario usuario)
        {
            try
            {
                usuarioDao.Update(usuario);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void Delete(int idUser)
        {
            try
            {
                usuarioDao.Delete(idUser);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
