using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MVC_CRUD.Models
{
    public class Employee
    {
        public int EmpCode { get; set; }
        [Required(ErrorMessage ="This field is required")]
        public String Name { get; set; }
        [Required(ErrorMessage = "This field is required")]

        public String Address { get; set; }
        [Required(ErrorMessage = "This field is required")]

        public String Designation { get; set; }
        public String Gender { get; set; }

      

    }

   
}