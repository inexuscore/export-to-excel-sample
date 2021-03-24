using System;

namespace ExportToExcelSample.Models
{
    public class User
    {
        public int Id { get; set; }

        public string Username { get; set; }

        public string Email { get; set; }

        public string SerialNumber { get; set; }

        public DateTime JoinedOn { get; set; }
    }
}
