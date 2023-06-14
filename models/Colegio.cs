using JCruzPorcel.Inscipcion_Colegio.Models.AlumnoData;

namespace JCruzPorcel.Inscipcion_Colegio.Models.ColegioData
{
    public class Colegio
    {
        public string Nombre { get; }
        public List<Grado> Grados { get; }
        public List<Alumno> AlumnosAsignados { get; }

        public Colegio(string nombre)
        {
            Nombre = nombre;
            Grados = new List<Grado>();
            AlumnosAsignados = new List<Alumno>();
        }

        public void AgregarGrado(int grado, int vacantes)
        {
            Grado nuevoGrado = new Grado(grado, vacantes);
            Grados.Add(nuevoGrado);
        }

        public bool TieneVacantes(int grado)
        {
            Grado? gradoEncontrado = Grados.FirstOrDefault(g => g.GradoEsIgual(grado));
            return gradoEncontrado?.VacantesDisponibles > 0;
        }

        public void AgregarAlumno(Alumno alumno)
        {
            Grado? gradoEncontrado = Grados.FirstOrDefault(g => g.GradoEsIgual(alumno.Grado));
            if (gradoEncontrado != null)
            {
                gradoEncontrado.VacantesDisponibles--;
                AlumnosAsignados.Add(alumno);
            }
        }
    }

    public class Grado
    {
        public int GradoNumero { get; }
        public int VacantesDisponibles { get; set; }

        public Grado(int grado, int vacantes)
        {
            GradoNumero = grado;
            VacantesDisponibles = vacantes;
        }

        public bool GradoEsIgual(int grado)
        {
            return GradoNumero == grado;
        }
    }
}
