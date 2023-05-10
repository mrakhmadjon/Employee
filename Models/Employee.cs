using DocumentFormat.OpenXml.Wordprocessing;
using Fingers10.ExcelExport.Attributes;
using System.ComponentModel.DataAnnotations;

namespace Employee.Models
{
    public class Employee
    {
        [IncludeInReport(Order = 1)]

        public int Id { get; set; }

        [Display(Name = "Telefon")]
        [Required]
        [IncludeInReport(Order = 2)]

        public string Phone { get; set; }

        [Display(Name = "FIO")]
        [Required]
        [IncludeInReport(Order = 3)]

        public string FIO { get; set; }

        [Display(Name = "Lavozim")]
        [Required]
        [IncludeInReport(Order = 4)]

        public string Title { get; set; }

        [Display(Name = "Adres")]
        [Required]
        [IncludeInReport(Order = 5)]

        public string Address { get; set; }

        [Display(Name = "Oylik")]
        [Required]
        [IncludeInReport(Order = 6)]

        public string Salary { get; set; }

        [Display(Name = "Ish kunlar")]
        [Required]
        [IncludeInReport(Order = 7)]

        public string WorkingDays { get; set; }

        [Display(Name = "Ish vaqti")]
        [Required]
        [IncludeInReport(Order = 8)]

        public string WorkingTime { get; set; }
    }
}
