using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTransformation.Librairie;
using DataTransformation.Models;

namespace DataTransformation.Interface
{
    public interface IDataInterface
    {
        Task ReadXLSFile(string pathFile);
    }
}
