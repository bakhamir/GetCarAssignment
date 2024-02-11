using GetCarAssignment.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Diagnostics;
using Dapper;
using ClosedXML;
using ClosedXML.Excel;

namespace GetCarAssignment.Controllers
{
    public class HomeController : Controller
    {

        private readonly string conStr = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=testdb;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        [HttpGet("json")]
        public IActionResult GetCarsJson()
        {
            using (var connection = new SqlConnection(conStr))
            {
                connection.Open();
                var cars = connection.Query<Car>("SELECT * FROM Car").ToList();
                return Ok(cars);
            }
        }

        [HttpGet("excel")]
        public IActionResult GetCarsExcel()
        {
            using (var connection = new SqlConnection(conStr))
            {
                connection.Open();
                var cars = connection.Query<Car>("SELECT * FROM Car").ToList();

                // Создание бинарного файла Excel с использованием библиотеки ClosedXML
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Cars");

                    worksheet.Cell(1, 1).Value = "Id";
                    worksheet.Cell(1, 2).Value = "Name";
                    worksheet.Cell(1, 3).Value = "Cost";
                    worksheet.Cell(1, 4).Value = "Model";

              
                    for (int i = 0; i < cars.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = cars[i].Id;
                        worksheet.Cell(i + 2, 2).Value = cars[i].Name;
                        worksheet.Cell(i + 2, 3).Value = cars[i].Cost;
                        worksheet.Cell(i + 2, 4).Value = cars[i].Model;
                    }

                    using (var stream = new System.IO.MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Seek(0, System.IO.SeekOrigin.Begin);


                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cars.xlsx");
                    }
                }
            }
        }

        [HttpGet("json/{id}")]
        public IActionResult GetCarJson(int id)
        {
            using (var connection = new SqlConnection(conStr))
            {
                connection.Open();
                var car = connection.QueryFirstOrDefault<Car>("SELECT * FROM Car WHERE Id = @Id", new { Id = id });
                if (car == null)
                {
                    return NotFound();
                }
                return Ok(car);

            }
        }
        [HttpPost]
        public IActionResult AddCar( Car newCar)
        {
            if (newCar == null)
            {
                return BadRequest("Invalid car data");
            }

            using (var connection = new SqlConnection(conStr))
            {
                connection.Execute("INSERT INTO Car (Name, Cost, Model) VALUES (@Name, @Cost, @Model)", newCar);
            }

            return Ok(newCar);
        }

        [HttpDelete("{id}")]
        public IActionResult DeleteCar(int id)
        {
            using (var connection = new SqlConnection(conStr))
            {
                var affectedRows = connection.Execute("DELETE FROM Car WHERE Id = @Id", new { Id = id });

                if (affectedRows > 0)
                {
                    return Ok($"Car with Id {id} has been deleted");
                }
                else
                {
                    return NotFound($"Car with Id {id} not found");
                }
            }
        }
    }
}