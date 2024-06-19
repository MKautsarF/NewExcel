using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewExcel
{
    internal class Model
    {
        public Model() { }
        public class Poin
        {
            public string Observasi { get; set; }
            public string Id { get; set; }
            public int Nilai { get; set; }
            public bool Disable { get; set; }
            public int Bobot { get; set; }
        }

        public class Data
        {
            public int Nomor { get; set; }
            public string LangkahKerja { get; set; }
            public bool Disable { get; set; }
            public int Bobot { get; set; }
            public List<Poin> Poin { get; set; }
        }

        public class Penilaian
        {
            public int Unit { get; set; }
            public string Judul { get; set; }
            public bool Disable { get; set; }
            public List<Data> Data { get; set; }
        }

        public class DataDiri
        {
            public string TrainType { get; set; }
            public string WaktuMulai { get; set; }
            public string WaktuSelesai { get; set; }
            public string Durasi { get; set; }
            public string Tanggal { get; set; }
            public string NamaCrew { get; set; }
            public string Kedudukan { get; set; }
            public string Usia { get; set; }
            public string KodeKedinasan { get; set; }
            public string NoKa { get; set; }
            public string Lintas { get; set; }
            public string NamaInstruktur { get; set; }
            public string Keterangan { get; set; }
            public List<Penilaian> Penilaian { get; set; }
            public int NilaiAkhir { get; set; }
        }
    }
}
