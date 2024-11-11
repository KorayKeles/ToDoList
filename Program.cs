using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;

namespace ToDoListApp
{
    class Program
    {
        static List<ToDoItem> toDoList = new List<ToDoItem>();

        static void Main(string[] args)
        {
            string command = "";

            Console.WriteLine("To-Do List Uygulamasına Hoş Geldiniz!");
            Console.WriteLine("Komutlar: ekle, sil, listele, tamamla, kaydet, çık");

            while (command != "çık")
            {
                Console.Write("\nKomut girin: ");
                command = Console.ReadLine()?.ToLower();

                switch (command)
                {
                    case "ekle":
                        Console.Write("Görev açıklaması: ");
                        string description = Console.ReadLine();
                        toDoList.Add(new ToDoItem(description));
                        Console.WriteLine("Görev eklendi.");
                        break;

                    case "sil":
                        Console.Write("Silmek istediğiniz görevin numarasını girin: ");
                        if (int.TryParse(Console.ReadLine(), out int removeIndex) && removeIndex >= 0 && removeIndex < toDoList.Count)
                        {
                            toDoList.RemoveAt(removeIndex);
                            Console.WriteLine("Görev silindi.");
                        }
                        else
                        {
                            Console.WriteLine("Geçersiz numara.");
                        }
                        break;

                    case "listele":
                        Console.WriteLine("\nGörev Listesi:");
                        for (int i = 0; i < toDoList.Count; i++)
                        {
                            Console.WriteLine($"{i}. {toDoList[i].Description} - {(toDoList[i].IsCompleted ? "Tamamlandı" : "Tamamlanmadı")}");
                        }
                        break;

                    case "tamamla":
                        Console.Write("Tamamladığınız görevin numarasını girin: ");
                        if (int.TryParse(Console.ReadLine(), out int completeIndex) && completeIndex >= 0 && completeIndex < toDoList.Count)
                        {
                            toDoList[completeIndex].MarkAsCompleted();
                            Console.WriteLine("Görev tamamlandı.");
                        }
                        else
                        {
                            Console.WriteLine("Geçersiz numara.");
                        }
                        break;

                    case "kaydet":
                        SaveToExcel();
                        Console.WriteLine("Görevler Excel dosyasına kaydedildi.");
                        break;

                    case "çık":
                        Console.WriteLine("Uygulamadan çıkılıyor.");
                        break;

                    default:
                        Console.WriteLine("Geçersiz komut.");
                        break;
                }
            }
        }

        static void SaveToExcel()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "GörevListesi.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Görev Listesi");

                // Başlıkları ekle
                worksheet.Cell(1, 1).Value = "Görev No";
                worksheet.Cell(1, 2).Value = "Açıklama";
                worksheet.Cell(1, 3).Value = "Durum";

                // Görevleri ekle
                for (int i = 0; i < toDoList.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = i + 1;
                    worksheet.Cell(i + 2, 2).Value = toDoList[i].Description;
                    worksheet.Cell(i + 2, 3).Value = toDoList[i].IsCompleted ? "Tamamlandı" : "Tamamlanmadı";
                }

                // Sütun genişliklerini içeriğe göre ayarla
                worksheet.Columns().AdjustToContents();

                // Dosyayı masaüstüne kaydet
                workbook.SaveAs(filePath);
            }
        }


    }

    class ToDoItem
    {
        public string Description { get; }
        public bool IsCompleted { get; private set; }

        public ToDoItem(string description)
        {
            Description = description;
            IsCompleted = false;
        }

        public void MarkAsCompleted()
        {
            IsCompleted = true;
        }
    }
}
