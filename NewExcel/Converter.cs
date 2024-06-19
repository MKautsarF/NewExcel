using JsontoExcel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewExcel
{
    internal class Converter
    {
        public Converter(string filename)
        {
            ConvertJsonToCsv(filename);     
        }


        public void ConvertJsonToCsv(string filename)
        {
            // Set license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read JSON data from file
            dynamic jsonData = ReadJSON.ReadJsonFile();
            int dataCount = 0;

            int totalPoin = 0;

            //// Iterasi melalui semua elemen "data" dalam array "penilaian" pertama
            //foreach (var data in jsonData.penilaian[1].data)
            //{
            //    // Menambahkan jumlah elemen "poin" dari setiap elemen "data"
            //    totalPoin += data.poin.Count;
            //}

            //string test = totalPoin.ToString();
            // Example data
            var data1 = new List<Model.DataDiri>
            {
                new Model.DataDiri
                {
                    //TrainType = test,
                    TrainType = jsonData.train_type,
                    WaktuMulai = jsonData.waktu_mulai,
                    WaktuSelesai = jsonData.waktu_selesai,
                    Durasi = jsonData.durasi,
                    Tanggal = jsonData.tanggal,
                    NamaCrew = jsonData.nama_crew,
                    Kedudukan = jsonData.kedudukan,
                    Usia = jsonData.usia,
                    KodeKedinasan = jsonData.kode_kedinasan,
                    NoKa = jsonData.no_ka,
                    Lintas = jsonData.lintas,
                    NamaInstruktur = jsonData.nama_instruktur,
                    Keterangan = jsonData.keterangan,
                    Penilaian = jsonData.penilaian.ToObject<List<Model.Penilaian>>(),
                    NilaiAkhir = jsonData.nilai_akhir,
                },
                // Add more records here
            };

            var data2 = new List<Model.Penilaian>();
            foreach (var item in jsonData.penilaian)
            {
                var penilaianItem = new Model.Penilaian
                {
                    Unit = item.unit,
                    Judul = item.judul,
                    Disable = item.disable,
                    Data = item.data.ToObject<List<Model.Data>>(),
                };
                data2.Add(penilaianItem);
            };

            var data3 = new List<Model.Data>();
            foreach (var item in jsonData.penilaian)
            {
                foreach (var item2 in item.data)
                {
                    var dataItem = new Model.Data
                    {
                        Nomor = item2.no,
                        LangkahKerja = item2.langkah_kerja,
                        Disable = item2.disable,
                        Bobot = item2.bobot,
                        Poin = item2.poin.ToObject<List<Model.Poin>>(),
                    };
                    data3.Add(dataItem);
                }
            }

            var data4 = new List<Model.Poin>();
            foreach (var item in jsonData.penilaian)
            {
                foreach (var item2 in item.data)
                {
                    foreach (var item3 in item2.poin)
                    {
                        var poinItem = new Model.Poin
                        {
                            Observasi = item3.observasi,
                            Id = item3.id,
                            Nilai = item3.nilai,
                            Disable = item3.disable,
                            Bobot = item3.bobot,
                        };
                        data4.Add(poinItem);
                    }
                }
            };

            //var directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"C:\Train Simulator\Data\penilaian\PDF\");
            string directoryPath = @"C:\Train Simulator\Data\penilaian\Excel\";
            Directory.CreateDirectory(directoryPath);

            // Assuming fileName contains the desired name of your PDF file (without extension)
            string filePath = Path.Combine(directoryPath, filename + ".xlsx");

            using (var package = new ExcelPackage())
            {
                // Add a worksheet
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");


                // data diri

                // Write header row spread across columns
                var headers = new List<string>
                {
                    "Tipe Kereta", "Waktu Mulai", "Waktu Selesai", "Durasi", "Tanggal",
                    "Nama Crew", "Kedudukan", "Usia", "Kode Kedinasan", "Nomor kereta Api", "Lintas",
                    "Nama Instruktur", "Keterangan", "Penilaian"
                    , "Nilai Akhir"
                    // Add more headers for additional properties
                };

                // Write header row spread across columns
                var headers2 = new List<string>
                {
                    "Unit", "Judul", "Disable", "Data"
                    // Add more headers for additional properties
                };

                // Write header row spread across columns
                var headers3 = new List<string>
                {
                    "Nomor", "Langkah Kerja", "Disable", "Bobot", "Poin"
                    // Add more headers for additional properties
                };

                // Write header row spread across columns
                var headers4 = new List<string>
                {
                    "ID", "Observasi", "Nilai", "Disable", "Bobot"
                    // Add more headers for additional properties
                };

                for (int i = 1; i <= headers.Count; i++)
                {
                    if (i != (headers.Count))
                    {
                        string valuekey = "";

                        switch (i )
                        {
                            case 1:
                                valuekey = "Tipe Kereta";
                                break;
                            case 2:
                                valuekey = "Waktu Mulai";
                                break;
                            case 3:
                                valuekey = "Waktu Selesai";
                                break;
                            case 4:
                                valuekey = "Durasi";
                                break;
                            case 5:
                                valuekey = "Tanggal";
                                break;
                            case 6:
                                valuekey = "Nama Crew";
                                break;
                            case 7:
                                valuekey = "Kedudukan";
                                break;
                            case 8:
                                valuekey = "Usia";
                                break;
                            case 9:
                                valuekey = "Kode Kedinasan";
                                break;
                            case 10:
                                valuekey = "Nomor KA";
                                break;
                            case 11:
                                valuekey = "Lintas";
                                break;
                            case 12:
                                valuekey = "Nama Instruktur";
                                break;
                            case 13:
                                valuekey = "Keterangan";
                                break;
                            // Add more cases as needed
                            default:
                                valuekey = "DefaultKey" + (i ); // You can modify this logic based on your requirements
                                break;
                        }
                        if (i == headers.Count - 1)
                        {
                            continue;
                        }
                        worksheet.Cells[1, i, 4, i].Merge = true;
                        worksheet.Cells[1, i].Value = valuekey;
                        worksheet.Cells[1, i, 4, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, i, 4, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // Add borders to the merged cells
                        var border = worksheet.Cells[1, i, 4, i].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;
                    }

                    else if (i == (headers.Count))
                    {
                        int dataSize = headers2.Count + headers3.Count + headers4.Count - 3;
                        //worksheet.Cells[1, i-1].Clear();
                        //worksheet.Cells[1, i].Clear();
                        worksheet.Cells[1, i - 1].Value = null;
                        worksheet.Cells[1, i].Value = null;

                        string test2 = headers[i - 2];

                        worksheet.Cells[1, i + 1, 1, i + 8].Merge = true;
                        worksheet.Cells[1, i+1].Value = test2;
                        worksheet.Cells[1, i+1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        int dataSize2 = headers2.Count + headers3.Count + headers4.Count - 3;
                        string test3 = headers[i - 1];
                        worksheet.Cells[1, i + dataSize2, 4, i + dataSize2].Merge = true;
                        worksheet.Cells[1, i+ dataSize2].Value = test3;
                        worksheet.Cells[1, i + dataSize2, 4, i + dataSize2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, i + dataSize2, 4, i + dataSize2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        string test4 = jsonData.nilai_akhir;
                        worksheet.Cells["Z5"].Value = test4;

                        // Add borders to the merged cells
                        var border = worksheet.Cells[1, i + dataSize2, 4, i + dataSize2].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;

                    }
                }

                // Write records starting from the second row for data1
                worksheet.Cells["A5"].LoadFromCollection(data1, false);


                // penilaian

                // Write data2 starting from the column after "Penilaian"
                var data2ColumnStart = headers.IndexOf("Penilaian") + 1;

                for (int i = 1; i <= headers2.Count; i++)
                {
                    if (i != (headers2.Count))
                    {
                        worksheet.Cells[2, data2ColumnStart, 4, data2ColumnStart].Merge = true;
                        worksheet.Cells[2, data2ColumnStart].Value = headers2[i - 1];
                        worksheet.Cells[2, data2ColumnStart, 4, data2ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, data2ColumnStart, 4, data2ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // Add borders to the merged cells
                        var border = worksheet.Cells[2, data2ColumnStart, 4, data2ColumnStart].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;

                        if (i == headers2.Count)
                        {
                            continue;
                        }
                    }
                    //worksheet.Cells[2, data2ColumnStart].Value = headers2[i - 1];

                    if (i == headers2.Count)
                    {
                        int dataSize = headers4.Count + headers3.Count - 2;
                        string test2 = headers2[i-1];
                        worksheet.Cells[2, data2ColumnStart, 2, data2ColumnStart + dataSize].Merge = true;
                        worksheet.Cells[2, data2ColumnStart].Value = test2;
                        worksheet.Cells[2, data2ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, data2ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        // Add borders to the merged cells
                        var border = worksheet.Cells[2, data2ColumnStart, 2, data2ColumnStart + dataSize].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;


                    }
                    data2ColumnStart++;
                }

                // Write data2 starting from the column after "Penilaian"
                data2ColumnStart = headers.IndexOf("Penilaian") + 1;

                int langkahKerjaList = 0;
                int poinList = 0;

                int penilaianCount = jsonData.penilaian.Count;
                int poinCount = 0;

                for (int i = 0; i < penilaianCount; i++)
                {
                    string test2 = (jsonData.penilaian[i].unit);
                    //string test2 = (penilaianCount).ToString();
                    string test3 = (jsonData.penilaian[i].judul);
                    string test4 = (jsonData.penilaian[i].disable);
                    dataCount = jsonData.penilaian[i].data.Count;

                    poinCount = 0;

                    // Iterasi melalui semua elemen "data" dalam array "penilaian" pertama
                    foreach (var data in jsonData.penilaian[i].data)
                    {
                        // Menambahkan jumlah elemen "poin" dari setiap elemen "data"
                        poinCount += data.poin.Count;
                    }

                    //poinCount = jsonData.penilaian[i].data[i].poin.Count;

                    worksheet.Cells[5 + poinList, data2ColumnStart].Value = test2;
                    worksheet.Cells[5 + poinList, data2ColumnStart + 1].Value = test3;
                    worksheet.Cells[5 + poinList, data2ColumnStart + 2].Value = test4;

                    langkahKerjaList = langkahKerjaList + dataCount;
                    poinList = poinList + poinCount;

                }

                // data

                // Write data2 starting from the column after "Penilaian"
                var data3ColumnStart = headers2.IndexOf("Data") + headers.IndexOf("Penilaian") + 1;

                for (int i = 1; i <= headers3.Count; i++)
                {
                    if (i != (headers3.Count))
                    {
                        worksheet.Cells[3, data3ColumnStart, 4, data3ColumnStart].Merge = true;
                        worksheet.Cells[3, data3ColumnStart].Value = headers3[i - 1];
                        worksheet.Cells[3, data3ColumnStart, 4, data3ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[3, data3ColumnStart, 4, data3ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // Add borders to the merged cells
                        var border = worksheet.Cells[3, data3ColumnStart, 4, data3ColumnStart].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;

                        if (i == headers3.Count)
                        {
                            continue;
                        }
                    }
                    //worksheet.Cells[3, data3ColumnStart].Value = headers3[i - 1];
                    if (i == headers3.Count)
                    {
                        int dataSize = headers4.Count -1;
                        string test2 = headers3[i - 1];
                        worksheet.Cells[3, data3ColumnStart, 3, data3ColumnStart + dataSize].Merge = true;
                        worksheet.Cells[3, data3ColumnStart].Value = test2;
                        worksheet.Cells[3, data3ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[3, data3ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        // Add borders to the merged cells
                        var border = worksheet.Cells[3, data3ColumnStart, 3, data3ColumnStart + dataSize].Style.Border;
                        border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;

                    }
                    data3ColumnStart++;
                }

                data3ColumnStart = headers2.IndexOf("Data") + headers.IndexOf("Penilaian") + 1;

                int dataLKCount = 0;
                poinList = 0;

                for (int i = 0; i < penilaianCount; i++)
                {
                    dataLKCount = jsonData.penilaian[i].data.Count;

                    for (int j = 0; j < dataLKCount; j++)
                    {
                        string test2 = (jsonData.penilaian[i].data[j].no);
                        string test3 = (jsonData.penilaian[i].data[j].langkah_kerja);
                        string test4 = (jsonData.penilaian[i].data[j].disable);
                        string test5 = (jsonData.penilaian[i].data[j].bobot);

                        poinCount = 0;

                        poinCount = jsonData.penilaian[i].data[j].poin.Count;

                        worksheet.Cells[5 + poinList, data3ColumnStart].Value = test2;
                        worksheet.Cells[5 + poinList, data3ColumnStart + 1].Value = test3;
                        worksheet.Cells[5 + poinList, data3ColumnStart + 2].Value = test4;
                        worksheet.Cells[5 + poinList, data3ColumnStart + 3].Value = test5;

                        poinList = poinList + poinCount;
                    }

                }

                // poin

                var data4ColumnStart = headers3.IndexOf("Poin") + headers2.IndexOf("Data") + headers.IndexOf("Penilaian") + 1;

                for (int i = 1;i <= headers4.Count; i++)
                {
                    worksheet.Cells[4, data4ColumnStart].Value = headers4[i - 1];
                    worksheet.Cells[4, data4ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[4, data4ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    data4ColumnStart++;

                    // Add borders to the merged cells
                    var border = worksheet.Cells[4, data4ColumnStart - 1].Style.Border;
                    border.Left.Style = border.Right.Style = border.Top.Style = border.Bottom.Style = ExcelBorderStyle.Thin;

                }

                data4ColumnStart = headers3.IndexOf("Poin") + headers2.IndexOf("Data") + headers.IndexOf("Penilaian") + 1;

                int poinLKCount = 0;
                poinList = 0;

                for (int i = 0; i < penilaianCount; i++)
                {
                    dataLKCount = jsonData.penilaian[i].data.Count;

                    for (int j = 0; j < dataLKCount; j++)
                    {
                        poinLKCount = jsonData.penilaian[i].data[j].poin.Count;
                        for (int k = 0; k < poinLKCount; k++)
                        {
                            string test2 = (jsonData.penilaian[i].data[j].poin[k].observasi);
                            string test3 = (jsonData.penilaian[i].data[j].poin[k].id);
                            string test4 = (jsonData.penilaian[i].data[j].poin[k].nilai);
                            string test5 = (jsonData.penilaian[i].data[j].poin[k].disable);
                            string test6 = (jsonData.penilaian[i].data[j].poin[k].bobot);

                            worksheet.Cells[5 + poinList, data4ColumnStart].Value = test3;
                            worksheet.Cells[5 + poinList, data4ColumnStart + 1].Value = test2;
                            worksheet.Cells[5 + poinList, data4ColumnStart + 2].Value = test4;
                            worksheet.Cells[5 + poinList, data4ColumnStart + 3].Value = test5;
                            worksheet.Cells[5 + poinList, data4ColumnStart + 4].Value = test6;

                            poinList++;

                        }
                    }

                }

                //for (int row = 1; row <= 4; row++)
                //{
                //    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                //    {
                //        var cell = worksheet.Cells[row, col];
                //        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //        cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                //    }
                //}

                // Set column width to auto outside the inner loop
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    int maxTextLength = 0;

                    for (int row = 1; row <= 4; row++)
                    {
                        var cell = worksheet.Cells[row, col];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        // Get the length of the text in the cell
                        int cellTextLength = cell.Text.Length;

                        // Update maxTextLength if the current cell has a longer text
                        if (cellTextLength > maxTextLength)
                        {
                            maxTextLength = cellTextLength;
                        }
                    }

                    // Set column width to auto based on the maximum text length in the column
                    worksheet.Column(col).Width = maxTextLength + 2; // Add some extra spaceint maxTextLength = 0;

                    for (int row = 1; row <= 4; row++)
                    {
                        var cell = worksheet.Cells[row, col];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        // Get the length of the text in the cell
                        int cellTextLength = cell.Text.Length;

                        // Update maxTextLength if the current cell has a longer text
                        if (cellTextLength > maxTextLength)
                        {
                            maxTextLength = cellTextLength;
                        }
                    }

                    // Set column width to auto based on the maximum text length in the column
                    worksheet.Column(col).Width = maxTextLength + 2; // Add some extra space
                    //worksheet.Column(col).AutoFit();
                }


                // Find the column index for "Nilai Akhir"
                int nilaiAkhirColumn = headers.IndexOf("Nilai Akhir") + 1;

                // Find the last column index
                int lastColumnIndex = headers.Count;

                // Merge N1 with the right side of the columns
                worksheet.Cells[1, nilaiAkhirColumn, 1, lastColumnIndex].Merge = true;

                // Save the Excel package to a file
                package.SaveAs(new FileInfo(filePath));
            }

            Console.WriteLine("XLSX file created successfully.");
        }

    }
}