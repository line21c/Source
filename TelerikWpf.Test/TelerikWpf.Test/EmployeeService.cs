using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TelerikWpf.Test
{
    public class EmployeeService
    {
        public static ObservableCollection<Employee> GetEmployees()
        {
            ObservableCollection<Employee> employees = new ObservableCollection<Employee>();
            Employee employee = new Employee();
            employee.FirstName = "Maria";
            employee.LastName = "Anders";
            employee.Married = true;
            employee.Age = 24;
            employees.Add(employee);        //... 
            Thread.Sleep(employees.Count * 10000);
            return employees;
        }
    }

    public class Employee
    {
        public string FirstName
        {
            get;
            set;
        }
        public string LastName
        {
            get;
            set;
        }
        public int Age
        {
            get;
            set;
        }
        public bool Married
        {
            get;
            set;
        }
    }
}
