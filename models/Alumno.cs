namespace JCruzPorcel.Inscipcion_Colegio.Models.AlumnoData
{
    public class Alumno
    {
        public int NroInscripcion { get; }
        public string Nombre { get; }
        public int Edad { get; }
        public int Grado { get; }
        public string Preferencia { get; }

        public Alumno(int nroInscripcion, string nombre, int edad, int grado, string preferencia)
        {
            NroInscripcion = nroInscripcion;
            Nombre = nombre;
            Edad = edad;
            Grado = grado;
            Preferencia = preferencia;
        }
    }
}
