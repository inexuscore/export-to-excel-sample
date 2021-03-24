using ClosedXML.Excel;
using ExportToExcelSample.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportToExcelSample.Controllers
{
    public class HomeController : Controller
    {
        private readonly List<User> users = new()
        {
            new User
            {
                Id = 1,
                Username = "ArminZia",
                Email = "armin.zia@gmail.com",
                SerialNumber = "NX33-AZ47",
                JoinedOn = new DateTime(1988, 04, 20)
            },
            new User
            {
                Id = 2,
                Username = "DoloresAbernathy",
                Email = "dolores.abernathy@gmail.com",
                SerialNumber = "CH1D-4AK7",
                JoinedOn = new DateTime(2021, 03, 24)
            },
            new User
            {
                Id = 3,
                Username = "MaeveMillay",
                Email = "maeve.millay@live.com",
                SerialNumber = "A33B-0JM2",
                JoinedOn = new DateTime(2021, 03, 23)
            },
            new User
            {
                Id = 4,
                Username = "BernardLowe",
                Email = "bernard.lowe@hotmail.com",
                SerialNumber = "H98M-LIP5",
                JoinedOn = new DateTime(2021, 03, 10)
            },
            new User
            {
                Id = 5,
                Username = "ManInBlack",
                Email = "maininblack@gmail.com",
                SerialNumber = "XN01-UT6C",
                JoinedOn = new DateTime(2021, 03, 9)
            }
        };

        public IActionResult Index()
        {
            return View(users);
        }

        public IActionResult Csv()
        {
            var builder = new StringBuilder();

            builder.AppendLine("Id,Username,Email,JoinedOn,SerialNumber");

            foreach (var user in users)
            {
                builder.AppendLine($"{user.Id},{user.Username},{user.Email},{user.JoinedOn.ToShortDateString()},{user.SerialNumber}");
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
        }

        public IActionResult Excel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Users");
            var currentRow = 1;

            worksheet.Row(currentRow).Height = 25.0;
            worksheet.Row(currentRow).Style.Font.Bold = true;
            worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGray;
            worksheet.Row(currentRow).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Cell(currentRow, 1).Value = "Id";
            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 2).Value = "Username";
            worksheet.Cell(currentRow, 3).Value = "Email";

            worksheet.Cell(currentRow, 4).Value = "Serial Number";
            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 5).Value = "Joined On";
            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            foreach (var user in users)
            {
                currentRow++;

                worksheet.Row(currentRow).Height = 20.0;
                worksheet.Row(currentRow).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                worksheet.Cell(currentRow, 1).Value = user.Id;
                worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                worksheet.Cell(currentRow, 2).Value = user.Username;

                worksheet.Cell(currentRow, 3).Value = user.Email;
                worksheet.Cell(currentRow, 3).Hyperlink.ExternalAddress = new Uri($"mailto:{user.Email}");

                worksheet.Cell(currentRow, 4).Value = user.SerialNumber;
                worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 4).Style.Fill.BackgroundColor = XLColor.PersianBlue;
                worksheet.Cell(currentRow, 4).Style.Font.FontColor = XLColor.WhiteSmoke;

                worksheet.Cell(currentRow, 5).Value = user.JoinedOn.ToShortDateString();
                worksheet.Cell(currentRow, 5).DataType = XLDataType.DateTime;
                worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                worksheet.Columns().AdjustToContents();
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users.xlsx");
        }
    }
}
