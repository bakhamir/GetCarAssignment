using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace GetCarAssignment.Models
{
    public class Car
    {
        public int Id { get; set; }


        public string Name { get; set; }


        public int Cost { get; set; }


        public string Model { get; set; }
    }
}
