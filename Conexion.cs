namespace AccesoDatos
{
    public class Conexion
    {
        public static GDatos GDatos;
        public static bool IniciarSesion(string nombreServidor, string baseDatos, string usuario, string password)
        {
            GDatos = new SqlServer(nombreServidor, baseDatos, usuario, password);
            return GDatos.Autenticar();
        } //fin inicializa sesion
 
        public static void FinalizarSesion()
        {
            GDatos.CerrarConexion();
        } //fin FinalizaSesion
 
    }//end class util
}//end namespace