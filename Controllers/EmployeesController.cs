using Fingers10.ExcelExport.ActionResults;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace BookStoreWithData.Controllers
{
    public class EmployeesController : Controller
    {

        public EmployeesController(IWebHostEnvironment webHostEnvironment)
        {
            this.webHostEnvironment = webHostEnvironment;
        }
        private readonly IWebHostEnvironment webHostEnvironment;
        private string excelPath = string.Empty;


        public IActionResult Download()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // get the current project directory
            // get the wwwroot directory path
            var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
            var employees = new List<Employee.Models.Employee>();
            // read data from the Excel file
            using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;
                //var employees = new List<Employee.Models.Employee>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var employee = new Employee.Models.Employee
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        FIO = worksheet.Cells[row, 2].Value.ToString(),
                        Title = worksheet.Cells[row, 3].Value.ToString(),
                        Salary = worksheet.Cells[row, 4].Value.ToString(),
                        Phone = worksheet.Cells[row, 5].Value.ToString(),
                        Address = worksheet.Cells[row, 6].Value.ToString(),
                        WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                        WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                    };

                    employees.Add(employee);
                }

            }

            return new ExcelResult<Employee.Models.Employee>(employees, "Sheet1", "Employees");
        }

        // GET: Employees
        public async Task<IActionResult> Index()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // get the current project directory
                // get the wwwroot directory path
                var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
                var employees = new List<Employee.Models.Employee>();
                // read data from the Excel file
                using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var startRow = worksheet.Dimension.Start.Row + 1;
                    var endRow = worksheet.Dimension.End.Row;
                    //var employees = new List<Employee.Models.Employee>();

                    for (int row = startRow; row <= endRow; row++)
                    {
                        var employee = new Employee.Models.Employee
                        {
                            Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            FIO = worksheet.Cells[row, 2].Value.ToString(),
                            Title = worksheet.Cells[row, 3].Value.ToString(),
                            Salary = worksheet.Cells[row, 4].Value.ToString(),
                            Phone = worksheet.Cells[row, 5].Value.ToString(),
                            Address = worksheet.Cells[row, 6].Value.ToString(),
                            WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                            WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                        };

                        employees.Add(employee);
                    }
                }
                return View(employees);

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        // GET: Employees/Details/5
        public async Task<IActionResult> Details(int id)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // get the current project directory
            // get the wwwroot directory path
            var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
            var employees = new List<Employee.Models.Employee>();
            // read data from the Excel file
            using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;
                //var employees = new List<Employee.Models.Employee>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var employee = new Employee.Models.Employee
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        FIO = worksheet.Cells[row, 2].Value.ToString(),
                        Title = worksheet.Cells[row, 3].Value.ToString(),
                        Salary = worksheet.Cells[row, 4].Value.ToString(),
                        Phone = worksheet.Cells[row, 5].Value.ToString(),
                        Address = worksheet.Cells[row, 6].Value.ToString(),
                        WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                        WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                    };

                    employees.Add(employee);
                }

            }
            var employe = employees.FirstOrDefault(p => p.Id == id);
            return View(employe);
        }

        // GET: Employees/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Employees/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,Phone,FIO,Title,Address,Salary,WorkingDays,WorkingTime")] Employee.Models.Employee employee)
        {
            try
            {

                if (ModelState.IsValid)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    // get the current project directory
                    // get the wwwroot directory path
                    var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
                    var employees = new List<Employee.Models.Employee>();
                    // read data from the Excel file
                    using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var startRow = worksheet.Dimension.Start.Row + 1;
                        var endRow = worksheet.Dimension.End.Row;
                        //var employees = new List<Employee.Models.Employee>();

                        for (int row = startRow; row <= endRow; row++)
                        {
                            var empl = new Employee.Models.Employee
                            {
                                Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                                FIO = worksheet.Cells[row, 2].Value.ToString(),
                                Title = worksheet.Cells[row, 3].Value.ToString(),
                                Salary = worksheet.Cells[row, 4].Value.ToString(),
                                Phone = worksheet.Cells[row, 5].Value.ToString(),
                                Address = worksheet.Cells[row, 6].Value.ToString(),
                                WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                                WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                            };

                            employees.Add(empl);
                        }

                        // add more employees to the list

                        // clear existing data from the worksheet
                        worksheet.Cells[startRow, 1, endRow, worksheet.Dimension.End.Column].Clear();

                        // add header row to the worksheet
                        worksheet.Cells["A1"].Value = "Id";
                        worksheet.Cells["B1"].Value = "FIO";
                        worksheet.Cells["C1"].Value = "Lavozim";
                        worksheet.Cells["D1"].Value = "Oylik";
                        worksheet.Cells["E1"].Value = "Telefon raqam";
                        worksheet.Cells["F1"].Value = "Mazil";
                        worksheet.Cells["G1"].Value = "Ish kunlar";
                        worksheet.Cells["H1"].Value = "Ish soat";
                        int? empId = employees.OrderByDescending(x => x.Id).FirstOrDefault()?.Id;
                        employee.Id = empId is null ? 1 : empId.Value + 1;
                        employees.Add(employee);
                        // insert new data into the worksheet
                        for (int i = 0; i < employees.Count; i++)
                        {
                            worksheet.Cells[startRow + i, 1].Value = employees[i].Id;
                            worksheet.Cells[startRow + i, 2].Value = employees[i].FIO;
                            worksheet.Cells[startRow + i, 3].Value = employees[i].Title;
                            worksheet.Cells[startRow + i, 4].Value = employees[i].Salary;
                            worksheet.Cells[startRow + i, 5].Value = employees[i].Phone;
                            worksheet.Cells[startRow + i, 6].Value = employees[i].Address;
                            worksheet.Cells[startRow + i, 7].Value = employees[i].WorkingDays;
                            worksheet.Cells[startRow + i, 8].Value = employees[i].WorkingTime;
                        }

                        // save the updated worksheet to the Excel file
                        package.Save();
                    }
                    return RedirectToAction("Index");
                }

                return View(employee);

            }
            catch (Exception)
            {

                throw;
            }
        }

        // GET: Employees/Edit/5
        public async Task<IActionResult> Edit(int id)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // get the current project directory
            // get the wwwroot directory path
            var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
            var employees = new List<Employee.Models.Employee>();
            // read data from the Excel file
            using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;
                //var employees = new List<Employee.Models.Employee>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var employee = new Employee.Models.Employee
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        FIO = worksheet.Cells[row, 2].Value.ToString(),
                        Title = worksheet.Cells[row, 3].Value.ToString(),
                        Salary = worksheet.Cells[row, 4].Value.ToString(),
                        Phone = worksheet.Cells[row, 5].Value.ToString(),
                        Address = worksheet.Cells[row, 6].Value.ToString(),
                        WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                        WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                    };

                    employees.Add(employee);
                }
                var employe = employees.FirstOrDefault(p => p.Id == id);
                return View(employe);
            }
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,Phone,FIO,Title,Address,Salary,WorkingDays,WorkingTime")] Employee.Models.Employee employee)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // get the current project directory
            // get the wwwroot directory path
            var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
            var employees = new List<Employee.Models.Employee>();
            // read data from the Excel file
            using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;
                //var employees = new List<Employee.Models.Employee>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var empl = new Employee.Models.Employee
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        FIO = worksheet.Cells[row, 2].Value.ToString(),
                        Title = worksheet.Cells[row, 3].Value.ToString(),
                        Salary = worksheet.Cells[row, 4].Value.ToString(),
                        Phone = worksheet.Cells[row, 5].Value.ToString(),
                        Address = worksheet.Cells[row, 6].Value.ToString(),
                        WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                        WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                    };

                    employees.Add(empl);
                }

                // add more employees to the list

                // clear existing data from the worksheet
                worksheet.Cells[startRow, 1, endRow, worksheet.Dimension.End.Column].Clear();

                // add header row to the worksheet
                worksheet.Cells["A1"].Value = "Id";
                worksheet.Cells["B1"].Value = "FIO";
                worksheet.Cells["C1"].Value = "Lavozim";
                worksheet.Cells["D1"].Value = "Oylik";
                worksheet.Cells["E1"].Value = "Telefon raqam";
                worksheet.Cells["F1"].Value = "Mazil";
                worksheet.Cells["G1"].Value = "Ish kunlar";
                worksheet.Cells["H1"].Value = "Ish soat";
                foreach (var emp in from emp in employees
                                    where emp.Id == id
                                    select emp)
                {
                    emp.FIO = employee.FIO;
                    emp.Address = employee.Address;
                    emp.Salary = employee.Salary;
                    emp.Phone = employee.Phone;
                    emp.Title = employee.Title;
                    emp.WorkingDays = employee.WorkingDays;
                    emp.WorkingTime = employee.WorkingTime;
                }
                // insert new data into the worksheet
                for (int i = 0; i < employees.Count; i++)
                {
                    worksheet.Cells[startRow + i, 1].Value = employees[i].Id;
                    worksheet.Cells[startRow + i, 2].Value = employees[i].FIO;
                    worksheet.Cells[startRow + i, 3].Value = employees[i].Title;
                    worksheet.Cells[startRow + i, 4].Value = employees[i].Salary;
                    worksheet.Cells[startRow + i, 5].Value = employees[i].Phone;
                    worksheet.Cells[startRow + i, 6].Value = employees[i].Address;
                    worksheet.Cells[startRow + i, 7].Value = employees[i].WorkingDays;
                    worksheet.Cells[startRow + i, 8].Value = employees[i].WorkingTime;
                }

                // save the updated worksheet to the Excel file
                package.Save();
            }
            return RedirectToAction("Index");
        }

        // GET: Employees/Delete/5
        public async Task<IActionResult> Delete(int id)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // get the current project directory
            // get the wwwroot directory path
            var wwwrootPath = webHostEnvironment.WebRootPath + "\\file.xlsx";
            var employees = new List<Employee.Models.Employee>();
            // read data from the Excel file
            using (var package = new ExcelPackage(new FileInfo(wwwrootPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;
                //var employees = new List<Employee.Models.Employee>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var employee = new Employee.Models.Employee
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        FIO = worksheet.Cells[row, 2].Value.ToString(),
                        Title = worksheet.Cells[row, 3].Value.ToString(),
                        Salary = worksheet.Cells[row, 4].Value.ToString(),
                        Phone = worksheet.Cells[row, 5].Value.ToString(),
                        Address = worksheet.Cells[row, 6].Value.ToString(),
                        WorkingDays = worksheet.Cells[row, 7].Value.ToString(),
                        WorkingTime = worksheet.Cells[row, 8].Value.ToString()
                    };

                    employees.Add(employee);
                }

                // add more employees to the list

                // clear existing data from the worksheet
                worksheet.Cells[startRow, 1, endRow, worksheet.Dimension.End.Column].Clear();

                // add header row to the worksheet
                worksheet.Cells["A1"].Value = "Id";
                worksheet.Cells["B1"].Value = "FIO";
                worksheet.Cells["C1"].Value = "Lavozim";
                worksheet.Cells["D1"].Value = "Oylik";
                worksheet.Cells["E1"].Value = "Telefon raqam";
                worksheet.Cells["F1"].Value = "Mazil";
                worksheet.Cells["G1"].Value = "Ish kunlar";
                worksheet.Cells["H1"].Value = "Ish soat";

                var empl = employees.FirstOrDefault(p => p.Id == id);
                if (empl is null)
                    return NotFound();

                employees.Remove(empl);

                // insert new data into the worksheet
                for (int i = 0; i < employees.Count; i++)
                {
                    worksheet.Cells[startRow + i, 1].Value = employees[i].Id;
                    worksheet.Cells[startRow + i, 2].Value = employees[i].FIO;
                    worksheet.Cells[startRow + i, 3].Value = employees[i].Title;
                    worksheet.Cells[startRow + i, 4].Value = employees[i].Salary;
                    worksheet.Cells[startRow + i, 5].Value = employees[i].Phone;
                    worksheet.Cells[startRow + i, 6].Value = employees[i].Address;
                    worksheet.Cells[startRow + i, 7].Value = employees[i].WorkingDays;
                    worksheet.Cells[startRow + i, 8].Value = employees[i].WorkingTime;
                }

                // save the updated worksheet to the Excel file
                package.Save();
            }
            return RedirectToAction("Index");
        }
    }
}
